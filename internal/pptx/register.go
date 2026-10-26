// Package pptx converts between .pptx (PresentationML) and the JSON slide
// model used by web/slides.js. Stdlib only.
package pptx

import (
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"strings"
)

const maxBody = 200 << 20

// Register adds this package's /api routes to mux.
func Register(mux *http.ServeMux) {
	mux.HandleFunc("POST /api/pptx/to-json", handleToJSON)
	mux.HandleFunc("POST /api/pptx/from-json", handleFromJSON)
}

func tooLarge(w http.ResponseWriter, err error) bool {
	var mbe *http.MaxBytesError
	if errors.As(err, &mbe) {
		http.Error(w, "request body too large (max 200 MB)", http.StatusRequestEntityTooLarge)
		return true
	}
	return false
}

func handleToJSON(w http.ResponseWriter, r *http.Request) {
	b, err := io.ReadAll(http.MaxBytesReader(w, r.Body, maxBody))
	if err != nil {
		if !tooLarge(w, err) {
			http.Error(w, "cannot read request: "+err.Error(), http.StatusBadRequest)
		}
		return
	}
	var m *Model
	var title string
	if len(b) == 0 {
		m, err = BlankModel()
		title = "Presentation"
	} else {
		m, title, err = ToModel(b)
	}
	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	enc.Encode(map[string]any{"title": title, "model": m})
}

func handleFromJSON(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, maxBody)
	var modelJSON []byte
	var original []byte
	ct := r.Header.Get("Content-Type")
	if strings.HasPrefix(ct, "multipart/") {
		if err := r.ParseMultipartForm(64 << 20); err != nil {
			if !tooLarge(w, err) {
				http.Error(w, "invalid multipart form: "+err.Error(), http.StatusBadRequest)
			}
			return
		}
		if r.MultipartForm != nil {
			defer r.MultipartForm.RemoveAll()
		}
		if v := r.FormValue("model"); v != "" {
			modelJSON = []byte(v)
		} else if f, _, err := r.FormFile("model"); err == nil {
			modelJSON, _ = io.ReadAll(f)
			f.Close()
		}
		if f, _, err := r.FormFile("original"); err == nil {
			original, err = io.ReadAll(f)
			f.Close()
			if err != nil {
				http.Error(w, "cannot read original: "+err.Error(), http.StatusBadRequest)
				return
			}
		}
	} else {
		b, err := io.ReadAll(r.Body)
		if err != nil {
			if !tooLarge(w, err) {
				http.Error(w, "cannot read request: "+err.Error(), http.StatusBadRequest)
			}
			return
		}
		modelJSON = b
	}
	if len(modelJSON) == 0 {
		http.Error(w, "missing \"model\" field (slide model JSON)", http.StatusBadRequest)
		return
	}
	var m Model
	if err := json.Unmarshal(modelJSON, &m); err != nil {
		// accept {"model": {...}} too
		var wrap struct {
			Model *Model `json:"model"`
		}
		if err2 := json.Unmarshal(modelJSON, &wrap); err2 != nil || wrap.Model == nil {
			http.Error(w, "invalid model JSON: "+err.Error(), http.StatusBadRequest)
			return
		}
		m = *wrap.Model
	} else if len(m.Slides) == 0 && m.Size.CX == 0 {
		var wrap struct {
			Model *Model `json:"model"`
		}
		if json.Unmarshal(modelJSON, &wrap) == nil && wrap.Model != nil {
			m = *wrap.Model
		}
	}
	out, err := FromModel(&m, original)
	if err != nil {
		http.Error(w, "cannot create .pptx: "+err.Error(), http.StatusInternalServerError)
		return
	}
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
	w.Write(out)
}

// BlankModel returns the model of a new 16:9 presentation with one empty
// Title Slide.
func BlankModel() (*Model, error) {
	title, sub := "ctrTitle", "subTitle"
	m := &Model{Size: Size{CX: 12192000, CY: 6858000}, Slides: []*Slide{{
		ID: "s1", Layout: "ppt/slideLayouts/slideLayout1.xml", Shapes: []*Shape{
			{ID: "sp2", Name: "Title 1", Kind: "text", Ph: &title, Paras: []*Para{{Runs: []*Run{}}}},
			{ID: "sp3", Name: "Subtitle 2", Kind: "text", Ph: &sub, PhIdx: "1", Paras: []*Para{{Runs: []*Run{}}}},
		}}}}
	b, err := FromModel(m, nil)
	if err != nil {
		return nil, err
	}
	out, _, err := ToModel(b)
	return out, err
}
