// Package docx converts between .docx (WordprocessingML) and simple semantic
// HTML suitable for a contenteditable editor. Stdlib only.
package docx

import (
	"encoding/json"
	"errors"
	"io"
	"net/http"
)

const maxBody = 100 << 20

// Register adds this package's /api routes to mux.
func Register(mux *http.ServeMux) {
	mux.HandleFunc("POST /api/docx/to-html", handleToHTML)
	mux.HandleFunc("POST /api/docx/from-html", handleFromHTML)
}

func readBody(w http.ResponseWriter, r *http.Request) ([]byte, bool) {
	b, err := io.ReadAll(http.MaxBytesReader(w, r.Body, maxBody))
	if err != nil {
		var mbe *http.MaxBytesError
		if errors.As(err, &mbe) {
			http.Error(w, "request body too large (max 100 MB)", http.StatusRequestEntityTooLarge)
		} else {
			http.Error(w, "cannot read request: "+err.Error(), http.StatusBadRequest)
		}
		return nil, false
	}
	return b, true
}

func handleToHTML(w http.ResponseWriter, r *http.Request) {
	b, ok := readBody(w, r)
	if !ok {
		return
	}
	if len(b) == 0 {
		http.Error(w, "empty request body: send the .docx file bytes", http.StatusBadRequest)
		return
	}
	h, title, err := ToHTML(b)
	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	enc.Encode(map[string]string{"html": h, "title": title})
}

func handleFromHTML(w http.ResponseWriter, r *http.Request) {
	b, ok := readBody(w, r)
	if !ok {
		return
	}
	var req struct {
		HTML string `json:"html"`
	}
	if err := json.Unmarshal(b, &req); err != nil {
		http.Error(w, "invalid JSON: expected {\"html\": \"...\"}", http.StatusBadRequest)
		return
	}
	out, err := FromHTML(req.HTML)
	if err != nil {
		http.Error(w, "cannot create .docx: "+err.Error(), http.StatusInternalServerError)
		return
	}
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
	w.Write(out)
}
