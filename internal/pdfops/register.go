// Package pdfops exposes PDF manipulation endpoints (/api/pdf/*) backed by pdfcpu.
//
// All endpoints take multipart/form-data with one or more "file" parts plus
// string fields, work entirely in memory, and answer with application/pdf
// unless documented otherwise. Every endpoint also accepts an optional
// "password" field used to open encrypted input files.
package pdfops

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"strconv"
	"strings"

	"github.com/pdfcpu/pdfcpu/pkg/api"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu/form"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu/model"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu/types"
)

const (
	maxMemory   = 200 << 20 // multipart parts kept in memory
	maxBodySize = 1 << 30   // hard request size limit
)

var baseConf *model.Configuration

func init() {
	// Never read or create a pdfcpu config dir in the user's profile.
	api.DisableConfigDir()
	c, err := api.LoadConfiguration(api.ConfigurationOptions{Mode: api.ConfigurationModeStateless})
	if err != nil || c == nil {
		c = model.NewDefaultConfiguration()
	}
	c.ValidationMode = model.ValidationRelaxed
	baseConf = c
}

// newConf returns a fresh relaxed configuration, optionally with passwords.
func newConf(password string) *model.Configuration {
	c := baseConf.Clone()
	c.ValidationMode = model.ValidationRelaxed
	if password != "" {
		c.UserPW = password
		c.OwnerPW = password
	}
	return c
}

// Register adds this package's /api routes to mux.
func Register(mux *http.ServeMux) {
	routes := map[string]func(*request) error{
		"merge":         opMerge,
		"extract":       opExtract,
		"delete":        opDelete,
		"rotate":        opRotate,
		"reorder":       opReorder,
		"watermark":     opWatermark,
		"encrypt":       opEncrypt,
		"decrypt":       opDecrypt,
		"compress":      opCompress,
		"split":         opSplit,
		"info":          opInfo,
		"images-to-pdf": opImagesToPDF,
		"form-fields":   opFormFields,
		"form-fill":     opFormFill,
		"annotate":      opAnnotate,
	}
	for name, fn := range routes {
		mux.HandleFunc("POST /api/pdf/"+name, handler(fn))
	}
}

// httpError carries an HTTP status with a user-facing message.
type httpError struct {
	status int
	msg    string
}

func (e *httpError) Error() string { return e.msg }

func badRequest(format string, a ...any) error {
	return &httpError{http.StatusBadRequest, fmt.Sprintf(format, a...)}
}

type request struct {
	w    http.ResponseWriter
	r    *http.Request
	ctx  context.Context
	form *multipart.Form
}

func (q *request) value(name string) string {
	if v := q.form.Value[name]; len(v) > 0 {
		return strings.TrimSpace(v[0])
	}
	return ""
}

func (q *request) conf() *model.Configuration { return newConf(q.r.FormValue("password")) }

// files returns the contents of all "file" parts in order.
func (q *request) files() ([][]byte, error) {
	fhs := q.form.File["file"]
	out := make([][]byte, 0, len(fhs))
	for _, fh := range fhs {
		f, err := fh.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(f)
		f.Close()
		if err != nil {
			return nil, err
		}
		out = append(out, b)
	}
	return out, nil
}

// file returns the single required "file" part.
func (q *request) file() (io.ReadSeeker, error) {
	ff, err := q.files()
	if err != nil {
		return nil, err
	}
	if len(ff) == 0 {
		return nil, badRequest("missing PDF file (form field \"file\")")
	}
	return bytes.NewReader(ff[0]), nil
}

func (q *request) sendPDF(b []byte) error {
	q.w.Header().Set("Content-Type", "application/pdf")
	q.w.Header().Set("Content-Length", strconv.Itoa(len(b)))
	_, err := q.w.Write(b)
	return err
}

func (q *request) sendJSON(v any) error {
	q.w.Header().Set("Content-Type", "application/json; charset=utf-8")
	return json.NewEncoder(q.w).Encode(v)
}

func handler(fn func(*request) error) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		r.Body = http.MaxBytesReader(w, r.Body, maxBodySize)
		if err := r.ParseMultipartForm(maxMemory); err != nil {
			http.Error(w, "invalid multipart request: "+err.Error(), http.StatusBadRequest)
			return
		}
		defer r.MultipartForm.RemoveAll()
		q := &request{w: w, r: r, ctx: r.Context(), form: r.MultipartForm}
		if err := fn(q); err != nil {
			status, msg := classify(err)
			http.Error(w, msg, status)
		}
	}
}

// classify maps an error to an HTTP status and readable message.
func classify(err error) (int, string) {
	var he *httpError
	if errors.As(err, &he) {
		return he.status, he.msg
	}
	msg := err.Error()
	low := strings.ToLower(msg)
	if strings.Contains(low, "password") || strings.Contains(low, "encrypt") && strings.Contains(low, "credentials") {
		return http.StatusBadRequest, "This PDF is password protected or the password is wrong: " + msg
	}
	if strings.Contains(low, "page selection") || strings.Contains(low, "invalid page") {
		return http.StatusBadRequest, msg
	}
	return http.StatusInternalServerError, "PDF operation failed: " + msg
}

// --- helpers ---

func pageCount(q *request, rs io.ReadSeeker) (int, error) {
	n, err := api.PageCount(q.ctx, rs, q.conf())
	if _, serr := rs.Seek(0, io.SeekStart); serr != nil && err == nil {
		err = serr
	}
	return n, err
}

// pageSelection validates a selection like "1-3,5" against the page count.
func pageSelection(q *request, rs io.ReadSeeker, field string, required bool) ([]string, error) {
	s := strings.ReplaceAll(q.value(field), " ", "")
	if s == "" {
		if required {
			return nil, badRequest("missing %q (e.g. 1-3,5)", field)
		}
		return nil, nil
	}
	sel, err := api.ParsePageSelection(s)
	if err != nil {
		return nil, badRequest("invalid %s %q: %v", field, s, err)
	}
	n, err := pageCount(q, rs)
	if err != nil {
		return nil, err
	}
	pages, err := api.PagesForPageSelection(n, sel, false, true)
	if err != nil {
		return nil, badRequest("invalid %s %q: %v", field, s, err)
	}
	if len(pages) == 0 {
		return nil, badRequest("%s %q selects no pages (document has %d)", field, s, n)
	}
	return sel, nil
}

// --- operations ---

func opMerge(q *request) error {
	ff, err := q.files()
	if err != nil {
		return err
	}
	if len(ff) < 2 {
		return badRequest("merge needs at least two PDF files")
	}
	rsc := make([]io.ReadSeeker, len(ff))
	for i, b := range ff {
		rsc[i] = bytes.NewReader(b)
	}
	var out bytes.Buffer
	if err := api.MergeRaw(q.ctx, rsc, &out, false, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opExtract(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	sel, err := pageSelection(q, rs, "pages", true)
	if err != nil {
		return err
	}
	var out bytes.Buffer
	if err := api.Trim(q.ctx, rs, &out, sel, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opDelete(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	sel, err := pageSelection(q, rs, "pages", true)
	if err != nil {
		return err
	}
	n, err := pageCount(q, rs)
	if err != nil {
		return err
	}
	del, err := api.PagesForPageSelection(n, sel, false, true)
	if err != nil {
		return badRequest("invalid pages: %v", err)
	}
	if len(del) >= n {
		return badRequest("cannot delete every page of the document")
	}
	var out bytes.Buffer
	if err := api.RemovePages(q.ctx, rs, &out, sel, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opRotate(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	deg, err := strconv.Atoi(q.value("degrees"))
	if err != nil || deg%90 != 0 || deg == 0 {
		return badRequest("degrees must be 90, 180 or 270")
	}
	deg = ((deg % 360) + 360) % 360
	sel, err := pageSelection(q, rs, "pages", false)
	if err != nil {
		return err
	}
	var out bytes.Buffer
	if err := api.Rotate(q.ctx, rs, &out, deg, sel, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opReorder(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	n, err := pageCount(q, rs)
	if err != nil {
		return err
	}
	raw := strings.ReplaceAll(q.value("order"), " ", "")
	if raw == "" {
		return badRequest("missing order (e.g. 3,1,2)")
	}
	var sel []string
	for _, p := range strings.Split(raw, ",") {
		v, err := strconv.Atoi(p)
		if err != nil || v < 1 || v > n {
			return badRequest("invalid page %q in order (document has %d pages)", p, n)
		}
		sel = append(sel, strconv.Itoa(v))
	}
	var out bytes.Buffer
	if err := api.Collect(q.ctx, rs, &out, sel, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opWatermark(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	text := q.value("text")
	if text == "" {
		return badRequest("missing watermark text")
	}
	opacity := 0.3
	if s := q.value("opacity"); s != "" {
		v, err := strconv.ParseFloat(s, 64)
		if err != nil || v < 0 || v > 1 {
			return badRequest("opacity must be between 0 and 1")
		}
		opacity = v
	}
	conf := q.conf()
	desc := fmt.Sprintf("fontname:Helvetica-Bold, scalefactor:0.7 rel, opacity:%.2f, fillcolor:#808080, diagonal:1", opacity)
	wm, err := api.TextWatermark(q.ctx, text, desc, true, false, types.POINTS, conf)
	if err != nil {
		return err
	}
	sel, err := pageSelection(q, rs, "pages", false)
	if err != nil {
		return err
	}
	var out bytes.Buffer
	if err := api.AddWatermarks(q.ctx, rs, &out, sel, wm, conf); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opEncrypt(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	pw := q.r.FormValue("password")
	if pw == "" {
		return badRequest("missing password")
	}
	owner := q.r.FormValue("ownerPassword")
	if owner == "" {
		owner = pw
	}
	conf := model.NewAESConfiguration(pw, owner, 256)
	conf.ValidationMode = model.ValidationRelaxed
	var out bytes.Buffer
	if err := api.Encrypt(q.ctx, rs, &out, conf); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opDecrypt(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	var out bytes.Buffer
	if err := api.Decrypt(q.ctx, rs, &out, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opCompress(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	var out bytes.Buffer
	if err := api.Optimize(q.ctx, rs, &out, q.conf(), nil); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opSplit(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	spans, err := api.SplitRaw(q.ctx, rs, 1, q.conf())
	if err != nil {
		return err
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	width := len(strconv.Itoa(len(spans)))
	if width < 3 {
		width = 3
	}
	for _, s := range spans {
		f, err := zw.Create(fmt.Sprintf("page-%0*d.pdf", width, s.From))
		if err != nil {
			return err
		}
		if _, err := io.Copy(f, s.Reader); err != nil {
			return err
		}
	}
	if err := zw.Close(); err != nil {
		return err
	}
	q.w.Header().Set("Content-Type", "application/zip")
	q.w.Header().Set("Content-Length", strconv.Itoa(buf.Len()))
	_, err = q.w.Write(buf.Bytes())
	return err
}

type pageSize struct {
	W float64 `json:"w"`
	H float64 `json:"h"`
}

type infoResult struct {
	Pages     int        `json:"pages"`
	Title     string     `json:"title"`
	Author    string     `json:"author"`
	Encrypted bool       `json:"encrypted"`
	PageSizes []pageSize `json:"pageSizes"`
}

func opInfo(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	conf := q.conf()
	info, err := api.PDFInfo(q.ctx, rs, "document.pdf", nil, false, conf)
	if err != nil {
		low := strings.ToLower(err.Error())
		if strings.Contains(low, "password") || strings.Contains(low, "encrypt") {
			// Report the document as locked instead of failing, so the UI can ask for a password.
			return q.sendJSON(infoResult{Encrypted: true, PageSizes: []pageSize{}})
		}
		return err
	}
	res := infoResult{Pages: info.PageCount, Title: info.Title, Author: info.Author, Encrypted: info.Encrypted, PageSizes: []pageSize{}}
	if _, err := rs.Seek(0, io.SeekStart); err != nil {
		return err
	}
	if dims, err := api.PageDims(q.ctx, rs, conf); err == nil {
		for _, d := range dims {
			res.PageSizes = append(res.PageSizes, pageSize{d.Width, d.Height})
		}
	}
	return q.sendJSON(res)
}

func opImagesToPDF(q *request) error {
	ff, err := q.files()
	if err != nil {
		return err
	}
	if len(ff) == 0 {
		return badRequest("no images supplied (form field \"file\")")
	}
	imgs := make([]io.Reader, len(ff))
	for i, b := range ff {
		imgs[i] = bytes.NewReader(b)
	}
	var imp *pdfcpu.Import
	switch ps := q.value("pageSize"); strings.ToLower(ps) {
	case "image", "full":
		imp = api.DefaultImportConfig() // page dimensions follow each image
	default:
		if ps == "" {
			ps = "A4"
		}
		imp, err = api.Import("formsize:"+ps+", position:c, scalefactor:0.95 rel", types.POINTS)
		if err != nil {
			return badRequest("invalid pageSize %q: %v", ps, err)
		}
	}
	var out bytes.Buffer
	if err := api.ImportImages(q.ctx, nil, &out, imgs, imp, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

type fieldResult struct {
	ID      string `json:"id"`
	Name    string `json:"name"`
	Type    string `json:"type"`
	Value   string `json:"value"`
	Options string `json:"options,omitempty"`
	Pages   []int  `json:"pages,omitempty"`
	Locked  bool   `json:"locked,omitempty"`
}

func fieldTypeName(t form.FieldType) string {
	switch t {
	case form.FTText:
		return "text"
	case form.FTDate:
		return "date"
	case form.FTCheckBox:
		return "checkbox"
	case form.FTComboBox:
		return "combobox"
	case form.FTListBox:
		return "listbox"
	case form.FTRadioButtonGroup:
		return "radio"
	}
	return strings.ToLower(t.String())
}

func opFormFields(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	fields, err := api.FormFields(q.ctx, rs, q.conf())
	if err != nil {
		low := strings.ToLower(err.Error())
		if strings.Contains(low, "no form") || strings.Contains(low, "acroform") {
			return q.sendJSON([]fieldResult{})
		}
		return err
	}
	out := make([]fieldResult, 0, len(fields))
	for _, f := range fields {
		out = append(out, fieldResult{ID: f.ID, Name: f.Name, Type: fieldTypeName(f.Typ), Value: f.V, Options: f.Opts, Pages: f.Pages, Locked: f.Locked})
	}
	return q.sendJSON(out)
}

// opFormFill fills fields given as {"nameOrID": value}; value is a string,
// bool (checkboxes) or string array (list boxes).
func opFormFill(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	var values map[string]any
	if err := json.Unmarshal([]byte(q.r.FormValue("values")), &values); err != nil {
		return badRequest("values must be a JSON object: %v", err)
	}
	conf := q.conf()
	fg, err := api.ExportForm(q.ctx, rs, "document.pdf", conf)
	if err != nil {
		return err
	}
	lookup := func(id, name string) (any, bool) {
		if v, ok := values[name]; ok && name != "" {
			return v, true
		}
		v, ok := values[id]
		return v, ok
	}
	str := func(v any) string {
		switch t := v.(type) {
		case string:
			return t
		case bool:
			return strconv.FormatBool(t)
		case float64:
			return strconv.FormatFloat(t, 'f', -1, 64)
		case nil:
			return ""
		}
		return fmt.Sprint(v)
	}
	truthy := func(v any) bool {
		switch t := v.(type) {
		case bool:
			return t
		case string:
			switch strings.ToLower(t) {
			case "true", "on", "yes", "1", "checked", "x":
				return true
			}
		case float64:
			return t != 0
		}
		return false
	}
	list := func(v any) []string {
		switch t := v.(type) {
		case []any:
			out := make([]string, 0, len(t))
			for _, e := range t {
				out = append(out, str(e))
			}
			return out
		case string:
			if t == "" {
				return nil
			}
			return strings.Split(t, ",")
		}
		return []string{str(v)}
	}
	for i := range fg.Forms {
		f := &fg.Forms[i]
		for _, x := range f.TextFields {
			if v, ok := lookup(x.ID, x.Name); ok {
				x.Value = str(v)
			}
		}
		for _, x := range f.DateFields {
			if v, ok := lookup(x.ID, x.Name); ok {
				x.Value = str(v)
			}
		}
		for _, x := range f.CheckBoxes {
			if v, ok := lookup(x.ID, x.Name); ok {
				x.Value = truthy(v)
			}
		}
		for _, x := range f.RadioButtonGroups {
			if v, ok := lookup(x.ID, x.Name); ok {
				x.Value = str(v)
			}
		}
		for _, x := range f.ComboBoxes {
			if v, ok := lookup(x.ID, x.Name); ok {
				x.Value = str(v)
			}
		}
		for _, x := range f.ListBoxes {
			if v, ok := lookup(x.ID, x.Name); ok {
				x.Values = list(v)
			}
		}
	}
	js, err := json.Marshal(fg)
	if err != nil {
		return err
	}
	if _, err := rs.Seek(0, io.SeekStart); err != nil {
		return err
	}
	var out bytes.Buffer
	if err := api.FillForm(q.ctx, rs, bytes.NewReader(js), &out, newConf(q.r.FormValue("password"))); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}

func opAnnotate(q *request) error {
	rs, err := q.file()
	if err != nil {
		return err
	}
	var items []Item
	if err := json.Unmarshal([]byte(q.r.FormValue("items")), &items); err != nil {
		return badRequest("items must be a JSON array: %v", err)
	}
	var out bytes.Buffer
	if err := Annotate(q.ctx, rs, &out, items, q.conf()); err != nil {
		return err
	}
	return q.sendPDF(out.Bytes())
}
