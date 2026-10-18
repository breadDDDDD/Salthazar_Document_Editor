package fsapi

import (
	"bytes"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"net/url"
	"os"
	"path/filepath"
	"testing"
)

func newServer(t *testing.T) *httptest.Server {
	t.Helper()
	mux := http.NewServeMux()
	Register(mux)
	srv := httptest.NewServer(mux)
	t.Cleanup(srv.Close)
	return srv
}

func q(p string) string { return url.QueryEscape(p) }

func TestList(t *testing.T) {
	dir := t.TempDir()
	for _, f := range []string{"b.txt", "A.pdf", "c.docx", ".hidden"} {
		if err := os.WriteFile(filepath.Join(dir, f), []byte("x"), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	for _, d := range []string{"zeta", "Alpha", ".git"} {
		if err := os.Mkdir(filepath.Join(dir, d), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	srv := newServer(t)
	res, err := http.Get(srv.URL + "/api/fs/list?path=" + q(dir))
	if err != nil {
		t.Fatal(err)
	}
	defer res.Body.Close()
	if res.StatusCode != 200 {
		t.Fatalf("status %d", res.StatusCode)
	}
	var l Listing
	if err := json.NewDecoder(res.Body).Decode(&l); err != nil {
		t.Fatal(err)
	}
	want := []string{"Alpha", "zeta", "A.pdf", "b.txt", "c.docx"}
	if len(l.Entries) != len(want) {
		t.Fatalf("got %d entries %+v, want %v", len(l.Entries), l.Entries, want)
	}
	for i, w := range want {
		if l.Entries[i].Name != w {
			t.Errorf("entry %d = %q, want %q", i, l.Entries[i].Name, w)
		}
	}
	if !l.Entries[0].IsDir || l.Entries[2].IsDir || l.Entries[2].Size != 1 {
		t.Errorf("bad entry fields: %+v", l.Entries)
	}
	if l.Parent != filepath.Dir(l.Path) {
		t.Errorf("parent = %q", l.Parent)
	}
}

func TestStatusCodes(t *testing.T) {
	dir := t.TempDir()
	missing := filepath.Join(dir, "nope")
	srv := newServer(t)
	cases := []struct {
		name, method, path string
		want               int
	}{
		{"list missing", "GET", "/api/fs/list?path=" + q(missing), 404},
		{"read missing", "GET", "/api/fs/read?path=" + q(missing), 404},
		{"read no path", "GET", "/api/fs/read", 400},
		{"read dir", "GET", "/api/fs/read?path=" + q(dir), 400},
		{"write missing parent", "POST", "/api/fs/write?path=" + q(filepath.Join(missing, "x.txt")), 400},
		{"write onto dir", "POST", "/api/fs/write?path=" + q(dir), 400},
		{"home", "GET", "/api/fs/home", 200},
		{"list drives or root", "GET", "/api/fs/list?path=", 200},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			req, _ := http.NewRequest(c.method, srv.URL+c.path, bytes.NewReader([]byte("data")))
			res, err := http.DefaultClient.Do(req)
			if err != nil {
				t.Fatal(err)
			}
			res.Body.Close()
			if res.StatusCode != c.want {
				t.Errorf("status %d, want %d", res.StatusCode, c.want)
			}
		})
	}
}

func TestWriteReadRoundtrip(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "note.txt")
	srv := newServer(t)
	for _, body := range []string{"first version", "second, overwriting"} {
		res, err := http.Post(srv.URL+"/api/fs/write?path="+q(p), "application/octet-stream", bytes.NewBufferString(body))
		if err != nil {
			t.Fatal(err)
		}
		var out struct {
			OK   bool   `json:"ok"`
			Path string `json:"path"`
		}
		json.NewDecoder(res.Body).Decode(&out)
		res.Body.Close()
		if !out.OK || out.Path != p {
			t.Fatalf("write response %+v (status %d)", out, res.StatusCode)
		}
		res, err = http.Get(srv.URL + "/api/fs/read?path=" + q(p))
		if err != nil {
			t.Fatal(err)
		}
		got, _ := io.ReadAll(res.Body)
		res.Body.Close()
		if string(got) != body {
			t.Errorf("read %q, want %q", got, body)
		}
		if ct := res.Header.Get("Content-Type"); ct == "" || ct == "application/octet-stream" {
			t.Errorf("content type %q for .txt", ct)
		}
	}
	des, _ := os.ReadDir(dir)
	if len(des) != 1 {
		t.Errorf("temp files left behind: %v", des)
	}
}
