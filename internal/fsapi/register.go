// Package fsapi exposes the local file system to the UI: listing folders,
// reading files and writing them atomically.
package fsapi

import (
	"encoding/json"
	"errors"
	"io"
	"io/fs"
	"mime"
	"net/http"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"
)

// Entry is one item of a folder listing.
type Entry struct {
	Name    string `json:"name"`
	Path    string `json:"path"`
	IsDir   bool   `json:"isDir"`
	Size    int64  `json:"size"`
	ModTime string `json:"modTime"`
}

// Listing is the response of GET /api/fs/list.
type Listing struct {
	Path    string  `json:"path"`
	Parent  string  `json:"parent"`
	Entries []Entry `json:"entries"`
}

// Place is a well-known folder shown in the file picker sidebar.
type Place struct {
	Name string `json:"name"`
	Path string `json:"path"`
}

// Register adds this package's /api routes to mux.
func Register(mux *http.ServeMux) {
	mux.HandleFunc("GET /api/fs/home", handleHome)
	mux.HandleFunc("GET /api/fs/list", handleList)
	mux.HandleFunc("GET /api/fs/read", handleRead)
	mux.HandleFunc("POST /api/fs/write", handleWrite)
}

func writeJSON(w http.ResponseWriter, v any) {
	w.Header().Set("Content-Type", "application/json")
	_ = json.NewEncoder(w).Encode(v)
}

// fsError maps a file system error to an HTTP status and plain-text message.
func fsError(w http.ResponseWriter, err error) {
	switch {
	case errors.Is(err, fs.ErrNotExist):
		http.Error(w, "not found: "+err.Error(), http.StatusNotFound)
	case errors.Is(err, fs.ErrPermission):
		http.Error(w, "permission denied: "+err.Error(), http.StatusForbidden)
	default:
		http.Error(w, err.Error(), http.StatusInternalServerError)
	}
}

func handleHome(w http.ResponseWriter, r *http.Request) {
	home, err := os.UserHomeDir()
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	places := []Place{}
	for _, name := range []string{"Desktop", "Documents", "Downloads"} {
		p := filepath.Join(home, name)
		if st, err := os.Stat(p); err == nil && st.IsDir() {
			places = append(places, Place{Name: name, Path: p})
		}
	}
	writeJSON(w, map[string]any{"path": home, "places": places})
}

// absPath cleans p and makes it absolute.
func absPath(p string) (string, error) {
	if p == "" {
		return "", errors.New("missing path")
	}
	return filepath.Abs(filepath.FromSlash(p))
}

func parentOf(p string) string {
	d := filepath.Dir(p)
	if d == p {
		return ""
	}
	return d
}

func handleList(w http.ResponseWriter, r *http.Request) {
	p := r.URL.Query().Get("path")
	if p == "" {
		if runtime.GOOS == "windows" {
			writeJSON(w, Listing{Path: "", Parent: "", Entries: drives()})
			return
		}
		p = "/"
	}
	l, err := list(p)
	if err != nil {
		fsError(w, err)
		return
	}
	writeJSON(w, l)
}

// drives returns the existing Windows drive roots as folder entries.
func drives() []Entry {
	out := []Entry{}
	for c := 'C'; c <= 'Z'; c++ {
		root := string(c) + `:\`
		if st, err := os.Stat(root); err == nil && st.IsDir() {
			out = append(out, Entry{Name: string(c) + ":", Path: root, IsDir: true, ModTime: st.ModTime().UTC().Format(time.RFC3339)})
		}
	}
	return out
}

func list(p string) (*Listing, error) {
	abs, err := absPath(p)
	if err != nil {
		return nil, err
	}
	st, err := os.Stat(abs)
	if err != nil {
		return nil, err
	}
	if !st.IsDir() {
		return nil, &fs.PathError{Op: "list", Path: abs, Err: errors.New("not a folder")}
	}
	des, err := os.ReadDir(abs)
	if err != nil {
		return nil, err
	}
	entries := make([]Entry, 0, len(des))
	for _, de := range des {
		name := de.Name()
		if strings.HasPrefix(name, ".") {
			continue
		}
		full := filepath.Join(abs, name)
		info, err := os.Stat(full) // follows symlinks
		if err != nil {
			continue
		}
		e := Entry{Name: name, Path: full, IsDir: info.IsDir(), ModTime: info.ModTime().UTC().Format(time.RFC3339)}
		if !e.IsDir {
			e.Size = info.Size()
		}
		entries = append(entries, e)
	}
	sort.SliceStable(entries, func(i, j int) bool {
		if entries[i].IsDir != entries[j].IsDir {
			return entries[i].IsDir
		}
		a, b := strings.ToLower(entries[i].Name), strings.ToLower(entries[j].Name)
		if a != b {
			return a < b
		}
		return entries[i].Name < entries[j].Name
	})
	return &Listing{Path: abs, Parent: parentOf(abs), Entries: entries}, nil
}

func handleRead(w http.ResponseWriter, r *http.Request) {
	abs, err := absPath(r.URL.Query().Get("path"))
	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	f, err := os.Open(abs)
	if err != nil {
		fsError(w, err)
		return
	}
	defer f.Close()
	st, err := f.Stat()
	if err != nil {
		fsError(w, err)
		return
	}
	if st.IsDir() {
		http.Error(w, "is a folder: "+abs, http.StatusBadRequest)
		return
	}
	ct := mime.TypeByExtension(strings.ToLower(filepath.Ext(abs)))
	if ct == "" {
		ct = "application/octet-stream"
	}
	w.Header().Set("Content-Type", ct)
	http.ServeContent(w, r, "", st.ModTime(), f)
}

func handleWrite(w http.ResponseWriter, r *http.Request) {
	abs, err := absPath(r.URL.Query().Get("path"))
	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	if err := writeAtomic(abs, r.Body); err != nil {
		var bad badRequest
		if errors.As(err, &bad) {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		fsError(w, err)
		return
	}
	writeJSON(w, map[string]any{"ok": true, "path": abs})
}

type badRequest string

func (b badRequest) Error() string { return string(b) }

// writeAtomic writes data to a temp file next to dst and renames it over dst.
func writeAtomic(dst string, data io.Reader) error {
	dir := filepath.Dir(dst)
	st, err := os.Stat(dir)
	if err != nil {
		if errors.Is(err, fs.ErrNotExist) {
			return badRequest("folder does not exist: " + dir)
		}
		return err
	}
	if !st.IsDir() {
		return badRequest("not a folder: " + dir)
	}
	if st, err := os.Stat(dst); err == nil && st.IsDir() {
		return badRequest("a folder with that name exists: " + dst)
	}
	tmp, err := os.CreateTemp(dir, "."+filepath.Base(dst)+".*.tmp")
	if err != nil {
		return err
	}
	tmpName := tmp.Name()
	ok := false
	defer func() {
		if !ok {
			tmp.Close()
			os.Remove(tmpName)
		}
	}()
	if _, err := io.Copy(tmp, data); err != nil {
		return err
	}
	if err := tmp.Sync(); err != nil {
		return err
	}
	if err := tmp.Close(); err != nil {
		return err
	}
	if err := os.Rename(tmpName, dst); err != nil {
		// Windows cannot always replace an existing file by rename.
		if rmErr := os.Remove(dst); rmErr != nil && !errors.Is(rmErr, fs.ErrNotExist) {
			return err
		}
		if err := os.Rename(tmpName, dst); err != nil {
			return err
		}
	}
	ok = true
	return nil
}
