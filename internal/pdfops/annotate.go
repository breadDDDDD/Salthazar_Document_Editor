package pdfops

import (
	"bytes"
	"context"
	"encoding/base64"
	"fmt"
	"io"
	"math/rand/v2"
	"strconv"
	"strings"

	"github.com/pdfcpu/pdfcpu/pkg/api"
	"github.com/pdfcpu/pdfcpu/pkg/font"
	pdffont "github.com/pdfcpu/pdfcpu/pkg/pdfcpu/font"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu/model"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu/types"
)

// Item is one thing to flatten onto a page.
//
// Coordinates are PDF points with the origin at the lower-left corner of the
// page's visible area (CropBox, falling back to MediaBox) of the UNROTATED
// page; a non-zero box origin is compensated for here. /Rotate is ignored.
//
// Text items: the box top is y+h; the first baseline is placed at
// y+h-0.8*fontSize and every further line ("\n") 1.2*fontSize lower.
// If h <= 0, y is the baseline of the first line. Text uses a standard PDF
// core font (Helvetica default, also Times-Roman, Courier and their
// Bold/Italic variants) with WinAnsi encoding; characters outside it become spaces.
//
// Image items: the PNG/JPEG data URL is drawn stretched to exactly x,y,w,h.
type Item struct {
	Page     int     `json:"page"`
	Type     string  `json:"type"`
	X        float64 `json:"x"`
	Y        float64 `json:"y"`
	W        float64 `json:"w"`
	H        float64 `json:"h"`
	Text     string  `json:"text"`
	FontSize float64 `json:"fontSize"`
	Color    string  `json:"color"`
	Font     string  `json:"font"`
	PNG      string  `json:"png"`
}

// Annotate reads a PDF from rs, draws items on top of their pages and writes the result to w.
func Annotate(c context.Context, rs io.ReadSeeker, w io.Writer, items []Item, conf *model.Configuration) error {
	ctx, err := api.ReadAndValidate(c, rs, conf)
	if err != nil {
		return err
	}
	byPage := map[int][]Item{}
	var order []int
	for i, it := range items {
		if it.Page < 1 || it.Page > ctx.PageCount {
			return badRequest("item %d: page %d out of range (document has %d pages)", i+1, it.Page, ctx.PageCount)
		}
		switch it.Type {
		case "text", "image":
		default:
			return badRequest("item %d: unknown type %q (want text or image)", i+1, it.Type)
		}
		if _, ok := byPage[it.Page]; !ok {
			order = append(order, it.Page)
		}
		byPage[it.Page] = append(byPage[it.Page], it)
	}
	for _, p := range order {
		if err := annotatePage(c, ctx, p, byPage[p]); err != nil {
			return err
		}
	}
	return api.WriteContext(c, ctx, w)
}

func annotatePage(c context.Context, ctx *model.Context, pageNr int, items []Item) error {
	xt := ctx.XRefTable
	pageDict, _, inh, err := xt.PageDict(c, pageNr, false)
	if err != nil {
		return err
	}
	box, err := visibleBox(xt, pageDict, inh)
	if err != nil {
		return err
	}
	ox, oy := box.LL.X, box.LL.Y

	fonts := types.Dict{}
	fontNames := map[string]string{} // core font name -> resource name
	xobjs := types.Dict{}
	var cs bytes.Buffer

	for i, it := range items {
		x, y := it.X+ox, it.Y+oy
		switch it.Type {
		case "text":
			if strings.TrimSpace(it.Text) == "" {
				continue
			}
			fn := it.Font
			if fn == "" || !font.IsCoreFont(fn) || fn == "Symbol" || fn == "ZapfDingbats" {
				fn = "Helvetica"
			}
			res, ok := fontNames[fn]
			if !ok {
				ref, err := pdffont.CoreFontDict(xt, fn)
				if err != nil {
					return err
				}
				res = "F" + strconv.Itoa(len(fontNames)+1)
				fontNames[fn] = res
				fonts[res] = *ref
			}
			size := it.FontSize
			if size <= 0 {
				size = 12
			}
			r, g, b := parseHex(it.Color)
			base := y + it.H - 0.8*size
			if it.H <= 0 {
				base = y
			}
			lines := strings.Split(strings.ReplaceAll(it.Text, "\r\n", "\n"), "\n")
			fmt.Fprintf(&cs, "BT /%s %s Tf %s %s %s rg %s TL %s %s Td\n", res, num(size), num(r), num(g), num(b), num(size*1.2), num(x), num(base))
			for li, line := range lines {
				if li > 0 {
					cs.WriteString("T* ")
				}
				fmt.Fprintf(&cs, "(%s) Tj\n", escapeString(model.DecodeUTF8ToByte(line)))
			}
			cs.WriteString("ET\n")
		case "image":
			if it.W <= 0 || it.H <= 0 {
				return badRequest("image item on page %d needs positive w and h", pageNr)
			}
			data, err := decodeDataURL(it.PNG)
			if err != nil {
				return badRequest("image item %d on page %d: %v", i+1, pageNr, err)
			}
			ref, _, _, err := model.CreateImageResource(xt, bytes.NewReader(data))
			if err != nil {
				return badRequest("image item %d on page %d: %v", i+1, pageNr, err)
			}
			name := "I" + strconv.Itoa(len(xobjs)+1)
			xobjs[name] = *ref
			fmt.Fprintf(&cs, "q %s 0 0 %s %s %s cm /%s Do Q\n", num(it.W), num(it.H), num(x), num(y), name)
		}
	}
	if cs.Len() == 0 {
		return nil
	}

	// Wrap everything in one Form XObject so our resource names cannot clash with the page's.
	formRes := types.Dict{}
	if len(fonts) > 0 {
		formRes["Font"] = fonts
	}
	if len(xobjs) > 0 {
		formRes["XObject"] = xobjs
	}
	mb := box
	if inhMB, err := mediaBox(xt, pageDict, inh); err == nil && inhMB != nil {
		mb = inhMB
	}
	formSD, err := xt.NewStreamDictForBuf(cs.Bytes())
	if err != nil {
		return err
	}
	formSD.InsertName("Type", "XObject")
	formSD.InsertName("Subtype", "Form")
	formSD.Insert("BBox", mb.Array())
	formSD.Insert("Resources", formRes)
	if err := formSD.Encode(); err != nil {
		return err
	}
	formRef, err := xt.IndRefForNewObject(*formSD)
	if err != nil {
		return err
	}

	// Register the form in the page's resources.
	formName := fmt.Sprintf("DSAnnot%d_%d", pageNr, rand.IntN(1_000_000))
	if err := addPageXObject(xt, pageDict, inh, formName, *formRef); err != nil {
		return err
	}

	// Isolate the existing content (q ... Q) so a leftover graphics state cannot shift our drawing.
	pre, err := newContentStream(xt, []byte("q\n"))
	if err != nil {
		return err
	}
	post, err := newContentStream(xt, []byte("\nQ q /"+formName+" Do Q\n"))
	if err != nil {
		return err
	}
	contents := types.Array{*pre}
	if obj, found := pageDict.Find("Contents"); found && obj != nil {
		switch o := obj.(type) {
		case types.IndirectRef:
			d, err := xt.Dereference(o)
			if err != nil {
				return err
			}
			if arr, ok := d.(types.Array); ok {
				contents = append(contents, arr...)
			} else {
				contents = append(contents, o)
			}
		case types.Array:
			contents = append(contents, o...)
		default:
			// A direct content stream is invalid PDF; wrap it as an object.
			ref, err := xt.IndRefForNewObject(o)
			if err != nil {
				return err
			}
			contents = append(contents, *ref)
		}
	}
	contents = append(contents, *post)
	pageDict["Contents"] = contents
	return nil
}

func newContentStream(xt *model.XRefTable, b []byte) (*types.IndirectRef, error) {
	sd, err := xt.NewStreamDictForBuf(b)
	if err != nil {
		return nil, err
	}
	if err := sd.Encode(); err != nil {
		return nil, err
	}
	return xt.IndRefForNewObject(*sd)
}

func addPageXObject(xt *model.XRefTable, pageDict types.Dict, inh *model.InheritedPageAttrs, name string, ref types.IndirectRef) error {
	var res types.Dict
	if obj, found := pageDict.Find("Resources"); found && obj != nil {
		d, err := xt.DereferenceDict(obj)
		if err != nil {
			return err
		}
		res = d
	}
	if res == nil {
		if inh != nil && inh.Resources != nil {
			res = inh.Resources.Clone().(types.Dict)
		} else {
			res = types.NewDict()
		}
		pageDict["Resources"] = res
	}
	var xo types.Dict
	if obj, found := res.Find("XObject"); found && obj != nil {
		d, err := xt.DereferenceDict(obj)
		if err != nil {
			return err
		}
		xo = d
	}
	if xo == nil {
		xo = types.NewDict()
		res["XObject"] = xo
	}
	xo[name] = ref
	return nil
}

func rectFrom(xt *model.XRefTable, obj types.Object) (*types.Rectangle, error) {
	a, err := xt.DereferenceArray(obj)
	if err != nil || len(a) != 4 {
		return nil, err
	}
	v := make([]float64, 4)
	for i := range a {
		f, err := xt.DereferenceNumber(a[i])
		if err != nil {
			return nil, err
		}
		v[i] = f
	}
	return types.NewRectangle(min(v[0], v[2]), min(v[1], v[3]), max(v[0], v[2]), max(v[1], v[3])), nil
}

func mediaBox(xt *model.XRefTable, d types.Dict, inh *model.InheritedPageAttrs) (*types.Rectangle, error) {
	if obj, ok := d.Find("MediaBox"); ok {
		if r, err := rectFrom(xt, obj); err == nil && r != nil {
			return r, nil
		}
	}
	if inh != nil && inh.MediaBox != nil {
		return inh.MediaBox, nil
	}
	return types.NewRectangle(0, 0, 612, 792), nil
}

// visibleBox returns the CropBox, or the MediaBox when none is set.
func visibleBox(xt *model.XRefTable, d types.Dict, inh *model.InheritedPageAttrs) (*types.Rectangle, error) {
	if obj, ok := d.Find("CropBox"); ok {
		if r, err := rectFrom(xt, obj); err == nil && r != nil {
			return r, nil
		}
	}
	if inh != nil && inh.CropBox != nil {
		return inh.CropBox, nil
	}
	return mediaBox(xt, d, inh)
}

func decodeDataURL(s string) ([]byte, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return nil, fmt.Errorf("missing png data URL")
	}
	if i := strings.Index(s, ","); strings.HasPrefix(s, "data:") && i >= 0 {
		if !strings.Contains(s[:i], ";base64") {
			return nil, fmt.Errorf("data URL must be base64 encoded")
		}
		s = s[i+1:]
	}
	b, err := base64.StdEncoding.DecodeString(s)
	if err != nil {
		b, err = base64.RawStdEncoding.DecodeString(strings.TrimRight(s, "="))
	}
	if err != nil {
		return nil, fmt.Errorf("invalid base64 image data: %v", err)
	}
	return b, nil
}

func parseHex(s string) (r, g, b float64) {
	s = strings.TrimPrefix(strings.TrimSpace(s), "#")
	if len(s) == 3 {
		s = string([]byte{s[0], s[0], s[1], s[1], s[2], s[2]})
	}
	if len(s) != 6 {
		return 0, 0, 0
	}
	v, err := strconv.ParseUint(s, 16, 32)
	if err != nil {
		return 0, 0, 0
	}
	return float64(v>>16&0xff) / 255, float64(v>>8&0xff) / 255, float64(v&0xff) / 255
}

func num(f float64) string { return strconv.FormatFloat(f, 'f', -1, 64) }

func escapeString(s string) string {
	var sb strings.Builder
	for i := 0; i < len(s); i++ {
		switch ch := s[i]; ch {
		case '(', ')', '\\':
			sb.WriteByte('\\')
			sb.WriteByte(ch)
		case '\r', '\n', '\t':
			sb.WriteByte(' ')
		default:
			sb.WriteByte(ch)
		}
	}
	return sb.String()
}
