package pdfops

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"image"
	"image/color"
	"image/png"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/pdfcpu/pdfcpu/pkg/api"
)

func testPNG(t *testing.T, w, h int, c color.Color) []byte {
	t.Helper()
	img := image.NewNRGBA(image.Rect(0, 0, w, h))
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			img.Set(x, y, c)
		}
	}
	var b bytes.Buffer
	if err := png.Encode(&b, img); err != nil {
		t.Fatal(err)
	}
	return b.Bytes()
}

// testPDF builds an n-page PDF from generated images.
func testPDF(t *testing.T, n int) []byte {
	t.Helper()
	srv := newServer()
	defer srv.Close()
	var files [][]byte
	for i := 0; i < n; i++ {
		files = append(files, testPNG(t, 40+i*10, 30, color.RGBA{uint8(60 * i), 100, 200, 255}))
	}
	b := post(t, srv, "images-to-pdf", files, map[string]string{"pageSize": "image"}, http.StatusOK)
	if got := pages(t, b); got != n {
		t.Fatalf("testPDF: got %d pages, want %d", got, n)
	}
	return b
}

func newServer() *httptest.Server {
	mux := http.NewServeMux()
	Register(mux)
	return httptest.NewServer(mux)
}

func post(t *testing.T, srv *httptest.Server, op string, files [][]byte, fields map[string]string, want int) []byte {
	t.Helper()
	var body bytes.Buffer
	mw := multipart.NewWriter(&body)
	for i, f := range files {
		fw, err := mw.CreateFormFile("file", "f"+string(rune('a'+i)))
		if err != nil {
			t.Fatal(err)
		}
		fw.Write(f)
	}
	for k, v := range fields {
		mw.WriteField(k, v)
	}
	mw.Close()
	resp, err := http.Post(srv.URL+"/api/pdf/"+op, mw.FormDataContentType(), &body)
	if err != nil {
		t.Fatal(err)
	}
	defer resp.Body.Close()
	var out bytes.Buffer
	out.ReadFrom(resp.Body)
	if resp.StatusCode != want {
		t.Fatalf("%s: status %d, want %d: %s", op, resp.StatusCode, want, out.String())
	}
	return out.Bytes()
}

func pages(t *testing.T, pdf []byte) int {
	t.Helper()
	n, err := api.PageCount(context.Background(), bytes.NewReader(pdf), newConf(""))
	if err != nil {
		t.Fatalf("page count: %v", err)
	}
	return n
}

func TestOperations(t *testing.T) {
	srv := newServer()
	defer srv.Close()
	pdf3 := testPDF(t, 3)
	pdf2 := testPDF(t, 2)

	if n := pages(t, post(t, srv, "merge", [][]byte{pdf3, pdf2}, nil, 200)); n != 5 {
		t.Errorf("merge: %d pages, want 5", n)
	}
	post(t, srv, "merge", [][]byte{pdf3}, nil, 400)

	if n := pages(t, post(t, srv, "extract", [][]byte{pdf3}, map[string]string{"pages": "1,3"}, 200)); n != 2 {
		t.Errorf("extract: %d pages, want 2", n)
	}
	post(t, srv, "extract", [][]byte{pdf3}, map[string]string{"pages": "9"}, 400)

	if n := pages(t, post(t, srv, "delete", [][]byte{pdf3}, map[string]string{"pages": "2"}, 200)); n != 2 {
		t.Errorf("delete: %d pages, want 2", n)
	}
	post(t, srv, "delete", [][]byte{pdf3}, map[string]string{"pages": "1-3"}, 400)

	if n := pages(t, post(t, srv, "rotate", [][]byte{pdf3}, map[string]string{"degrees": "90"}, 200)); n != 3 {
		t.Errorf("rotate: %d pages", n)
	}
	post(t, srv, "rotate", [][]byte{pdf3}, map[string]string{"degrees": "45"}, 400)

	reordered := post(t, srv, "reorder", [][]byte{pdf3}, map[string]string{"order": "3,1,2"}, 200)
	if n := pages(t, reordered); n != 3 {
		t.Errorf("reorder: %d pages", n)
	}
	// Page widths were 40,50,60 px -> page 1 of reordered must be the widest original.
	var info infoResult
	if err := json.Unmarshal(post(t, srv, "info", [][]byte{reordered}, nil, 200), &info); err != nil {
		t.Fatal(err)
	}
	if info.Pages != 3 || len(info.PageSizes) != 3 {
		t.Fatalf("info: %+v", info)
	}
	if !(info.PageSizes[0].W > info.PageSizes[2].W && info.PageSizes[2].W > info.PageSizes[1].W) {
		t.Errorf("reorder: unexpected page widths %+v", info.PageSizes)
	}

	if n := pages(t, post(t, srv, "watermark", [][]byte{pdf3}, map[string]string{"text": "DRAFT", "opacity": "0.4"}, 200)); n != 3 {
		t.Errorf("watermark: %d pages", n)
	}
	a4 := post(t, srv, "images-to-pdf", [][]byte{testPNG(t, 10, 10, color.Black)}, nil, 200)
	if err := json.Unmarshal(post(t, srv, "info", [][]byte{a4}, nil, 200), &info); err != nil || info.PageSizes[0].W != 595 {
		t.Errorf("images-to-pdf default A4: %+v %v", info, err)
	}
	if n := pages(t, post(t, srv, "compress", [][]byte{pdf3}, nil, 200)); n != 3 {
		t.Errorf("compress: %d pages", n)
	}

	enc := post(t, srv, "encrypt", [][]byte{pdf3}, map[string]string{"password": "s3cret"}, 200)
	if err := json.Unmarshal(post(t, srv, "info", [][]byte{enc}, nil, 200), &info); err != nil || !info.Encrypted {
		t.Errorf("info on encrypted: %+v %v", info, err)
	}
	post(t, srv, "decrypt", [][]byte{enc}, map[string]string{"password": "wrong"}, 400)
	dec := post(t, srv, "decrypt", [][]byte{enc}, map[string]string{"password": "s3cret"}, 200)
	if n := pages(t, dec); n != 3 {
		t.Errorf("decrypt: %d pages", n)
	}

	zb := post(t, srv, "split", [][]byte{pdf3}, nil, 200)
	zr, err := zip.NewReader(bytes.NewReader(zb), int64(len(zb)))
	if err != nil {
		t.Fatal(err)
	}
	if len(zr.File) != 3 || zr.File[0].Name != "page-001.pdf" {
		t.Errorf("split: %d entries", len(zr.File))
	}

	var fields []fieldResult
	if err := json.Unmarshal(post(t, srv, "form-fields", [][]byte{pdf3}, nil, 200), &fields); err != nil || len(fields) != 0 {
		t.Errorf("form-fields: %v %v", fields, err)
	}
}

func TestAnnotate(t *testing.T) {
	srv := newServer()
	defer srv.Close()
	pdf := testPDF(t, 2)
	sig := "data:image/png;base64," + base64.StdEncoding.EncodeToString(testPNG(t, 20, 10, color.NRGBA{255, 0, 0, 128}))
	items := []Item{
		{Page: 1, Type: "text", X: 50, Y: 700, W: 200, H: 40, Text: "Hello (world) \\ é\nsecond line", FontSize: 14, Color: "#336699", Font: "Times-Roman"},
		{Page: 2, Type: "text", X: 10, Y: 10, Text: "Courier text", Font: "Courier"},
		{Page: 2, Type: "image", X: 100, Y: 100, W: 150, H: 75, PNG: sig},
	}
	js, _ := json.Marshal(items)
	out := post(t, srv, "annotate", [][]byte{pdf}, map[string]string{"items": string(js)}, 200)
	if n := pages(t, out); n != 2 {
		t.Fatalf("annotate: %d pages", n)
	}
	if err := api.Validate(context.Background(), bytes.NewReader(out), newConf(""), nil); err != nil {
		t.Errorf("annotated PDF fails validation: %v", err)
	}
	// Annotating again (stacking) must still work.
	out2 := post(t, srv, "annotate", [][]byte{out}, map[string]string{"items": string(js)}, 200)
	if n := pages(t, out2); n != 2 {
		t.Fatalf("annotate twice: %d pages", n)
	}

	bad, _ := json.Marshal([]Item{{Page: 5, Type: "text", Text: "x"}})
	post(t, srv, "annotate", [][]byte{pdf}, map[string]string{"items": string(bad)}, 400)
	bad, _ = json.Marshal([]Item{{Page: 1, Type: "image", W: 10, H: 10, PNG: "data:image/png;base64,!!"}})
	post(t, srv, "annotate", [][]byte{pdf}, map[string]string{"items": string(bad)}, 400)
	post(t, srv, "annotate", [][]byte{pdf}, map[string]string{"items": "nope"}, 400)
}

func TestAnnotateContentStream(t *testing.T) {
	// The text must end up as real text operators inside the page's form XObject.
	srv := newServer()
	defer srv.Close()
	pdf := testPDF(t, 1)
	js, _ := json.Marshal([]Item{{Page: 1, Type: "text", X: 20, Y: 20, Text: "FindMe", FontSize: 10}})
	out := post(t, srv, "annotate", [][]byte{pdf}, map[string]string{"items": string(js)}, 200)
	ctx, err := api.ReadAndValidate(context.Background(), bytes.NewReader(out), newConf(""))
	if err != nil {
		t.Fatal(err)
	}
	xt := ctx.XRefTable
	pd, _, _, err := xt.PageDict(context.Background(), 1, false)
	if err != nil {
		t.Fatal(err)
	}
	res, _ := xt.DereferenceDict(pd["Resources"])
	xo, _ := xt.DereferenceDict(res["XObject"])
	found := false
	for name, ref := range xo {
		if !strings.HasPrefix(name, "DSAnnot") {
			continue
		}
		sd, _, err := xt.DereferenceStreamDict(ref)
		if err != nil || sd == nil {
			t.Fatalf("deref form: %v", err)
		}
		if err := sd.Decode(); err != nil {
			t.Fatal(err)
		}
		if strings.Contains(string(sd.Content), "(FindMe) Tj") {
			found = true
		}
	}
	if !found {
		t.Error("annotation text not found in form XObject")
	}
}
