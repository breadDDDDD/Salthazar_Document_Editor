// Package doc converts legacy Word 97-2003 binary documents (.doc, MS-DOC)
// to the editor's HTML dialect (the same one internal/docx produces).
// Stdlib only.
package doc

import (
	"encoding/json"
	"errors"
	"io"
	"net/http"
)

const maxBody = 100 << 20

// Register adds this package's /api routes to mux.
func Register(mux *http.ServeMux) {
	mux.HandleFunc("POST /api/doc/to-html", handleToHTML)
}

func handleToHTML(w http.ResponseWriter, r *http.Request) {
	b, err := io.ReadAll(http.MaxBytesReader(w, r.Body, maxBody))
	if err != nil {
		var mbe *http.MaxBytesError
		if errors.As(err, &mbe) {
			http.Error(w, "request body too large (max 100 MB)", http.StatusRequestEntityTooLarge)
		} else {
			http.Error(w, "cannot read request: "+err.Error(), http.StatusBadRequest)
		}
		return
	}
	if len(b) == 0 {
		http.Error(w, "empty request body: send the .doc file bytes", http.StatusBadRequest)
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
