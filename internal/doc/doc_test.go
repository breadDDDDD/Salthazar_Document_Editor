package doc

import (
	"bytes"
	"encoding/json"
	"math/rand"
	"net/http"
	"net/http/httptest"
	"os"
	"regexp"
	"strings"
	"testing"
)

const formPath = "../../Form Permohonan Ujian Skripsi S1 Fisika.doc"

func readFixture(t *testing.T, p string) []byte {
	t.Helper()
	b, err := os.ReadFile(p)
	if err != nil {
		t.Skipf("fixture not available: %v", err)
	}
	return b
}

// plain flattens HTML to text (entities decoded) for content assertions.
func plain(h string) string {
	return stripTags(strings.ReplaceAll(h, "</p>", "</p>\n"))
}

func checkClean(t *testing.T, h string) {
	t.Helper()
	for i, r := range h {
		if r < 0x20 && r != '\t' && r != '\n' {
			t.Fatalf("stray control char %#x at %d: %q", r, i, h[max(0, i-40):min(len(h), i+40)])
		}
	}
	for _, tag := range []string{"p", "table", "tr", "td", "ul", "ol", "li", "span", "strong", "em", "div", "a"} {
		o := len(regexp.MustCompile(`<`+tag+`[ >]`).FindAllString(h, -1))
		c := strings.Count(h, "</"+tag+">")
		if o != c {
			t.Errorf("unbalanced <%s>: %d open, %d close", tag, o, c)
		}
	}
}

func TestFormDoc(t *testing.T) {
	b := readFixture(t, formPath)
	h, _, err := ToHTML(b)
	if err != nil {
		t.Fatal(err)
	}
	checkClean(t, h)
	txt := plain(h)
	for _, want := range []string{
		"Program Studi : Fisika",
		"Permohonan Penjadwalan Ujian Skripsi/Tugas Akhir",
		"Yang bertanda tangan di bawah ini saya,",
		"Geervan Thanra",
		"21/477331/PA/20650",
		"PREDIKSI ENERGI CELAH PITA PADA PEROVSKIT HALIDA GANDA MENGGUNAKAN MACHINE LEARNING",
		"Moh. Adhib Ulil Absor, S.Si., M.Sc.,Ph.D.",
		"Mengajukan permohonan untuk menempuh ujian skripsi/tugas akhir",
		"Atas perhatian Bapak/Ibu saya ucapkan terima kasih.",
		"Yogyakarta, 20 Juni  2025",
		"Mengetahui/ menyetujui",
		"Moh. Adhib Ulil Absor, S.Si., M.Sc., Ph.D.,", // text box content
		"Telp. 085921554041",
	} {
		if !strings.Contains(txt, want) {
			t.Errorf("missing text %q", want)
		}
	}
	if !strings.HasPrefix(h, `<div class="docx-doc" data-page-w="11940" data-page-h="16860"`) {
		t.Errorf("missing/incorrect page wrapper: %.200s", h)
	}
	// Table: 3 rows of "n. Hari, tanggal : date , Jam : time".
	if n := strings.Count(h, "<table"); n != 1 {
		t.Errorf("want 1 table, got %d", n)
	}
	if n := strings.Count(h, "<tr>"); n != 3 {
		t.Errorf("want 3 table rows, got %d", n)
	}
	if n := strings.Count(txt, "Hari, tanggal"); n != 3 {
		t.Errorf("want 3 'Hari, tanggal' cells, got %d", n)
	}
	if n := strings.Count(h, "<td"); n != 18 {
		t.Errorf("want 18 cells, got %d", n)
	}
	// Formatting.
	for _, want := range []string{
		"<strong>Geervan Thanra</strong>",
		"<em>MACHINE LEARNING</em>",
		"font-family:&#39;Times New Roman&#39;",
		"font-size:18pt",
		"font-size:12pt",
		"text-align:center",
		"text-align:right",
		`<div class="docx-textbox"`,
	} {
		if !strings.Contains(h, want) {
			t.Errorf("missing formatting %q", want)
		}
	}
	// Pictures: the letterhead (PNG, inline) and the signature (JPEG, floating).
	if !strings.Contains(h, `<img src="data:image/png;base64,`) {
		t.Error("missing inline PNG letterhead")
	}
	if !strings.Contains(h, `<img src="data:image/jpeg;base64,`) {
		t.Error("missing floating JPEG picture")
	}
	if strings.Contains(h, "[picture") || strings.Contains(h, "picture]") {
		t.Error("unexpected picture placeholder")
	}
}

func TestRichDoc(t *testing.T) {
	// testdata/rich.doc was produced by Microsoft Word (Save As Word 97-2003).
	b := readFixture(t, "testdata/rich.doc")
	h, title, err := ToHTML(b)
	if err != nil {
		t.Fatal(err)
	}
	checkClean(t, h)
	if title != "Chapter One" {
		t.Errorf("title = %q", title)
	}
	for _, want := range []string{
		`<header class="docx-header">`, "Header text here", `<footer class="docx-footer">`, "Footer text",
		`<h1 `, ">Chapter One</h1>", `<h2 `, ">Section Two</h2>",
		"<strong>bold</strong>", "<em>italic</em>", "<u>under</u>", "<s>strike</s>",
		`<span style="color:#ff0000">red</span>`, "font-size:16pt;font-family:Arial", "<sub>2</sub>", "<sup>2</sup>",
		"background-color:#ffff00", `<a href="https://example.com/a?b=1">`, "text-align:center",
		`<ul data-numfmt="bullet"`, "bullet one", "bullet two",
		`<ol start="1" data-numfmt="decimal"`, "num one", "num two",
		`colspan="2"`, `rowspan="2"`, "border-top:0.5pt solid #000000",
		`<img src="data:image/png;base64,`, ` width="40" height="20"`,
		`<hr class="page-break">`, "Second page text",
		"margin-top:12pt", "margin-left:36pt", "line-height:1.725",
		`<div class="docx-footnotes">`, "<sup>1</sup> The footnote text.",
	} {
		if !strings.Contains(h, want) {
			t.Errorf("missing %q", want)
		}
	}
	if strings.Count(h, "<li ") != 4 {
		t.Errorf("want 4 list items, got %d", strings.Count(h, "<li "))
	}
	if strings.Contains(plain(h), "HYPERLINK") {
		t.Error("field instruction text leaked into output")
	}
}

func TestErrors(t *testing.T) {
	if _, _, err := ToHTML([]byte("hello world, not a doc at all")); err == nil {
		t.Error("expected error for plain text")
	}
	if _, _, err := ToHTML([]byte("{\\rtf1\\ansi hello}")); err == nil || !strings.Contains(err.Error(), "RTF") {
		t.Errorf("expected RTF error, got %v", err)
	}
	b := readFixture(t, formPath)

	// Rename the WordDocument stream to make it look like an Excel file.
	x := append([]byte(nil), b...)
	name := utf16le("WordDocument")
	i := bytes.Index(x, name)
	if i < 0 {
		t.Fatal("directory entry not found")
	}
	copy(x[i:], utf16le("Workbook\x00\x00\x00\x00"))
	if _, _, err := ToHTML(x); err == nil || !strings.Contains(err.Error(), "Excel") {
		t.Errorf("expected Excel error, got %v", err)
	}

	// Set fEncrypted in the FIB.
	cf, err := openCFB(b)
	if err != nil {
		t.Fatal(err)
	}
	var start uint32
	for _, e := range cf.rootChildren() {
		if e.name == "WordDocument" {
			start = e.start
		}
	}
	y := append([]byte(nil), b...)
	off := (int(start) + 1) * cf.secSize
	y[off+0x0B] |= 0x01
	if _, _, err := ToHTML(y); err != errEncrypted {
		t.Errorf("expected encryption error, got %v", err)
	}
}

func utf16le(s string) []byte {
	var out []byte
	for _, r := range s {
		out = append(out, byte(r), byte(r>>8))
	}
	return out
}

func TestCorruptNoPanic(t *testing.T) {
	b := readFixture(t, formPath)
	rnd := rand.New(rand.NewSource(1))
	for n := 0; n < 300; n++ {
		x := append([]byte(nil), b...)
		switch n % 3 {
		case 0:
			x = x[:rnd.Intn(len(x))]
		default:
			for k := 0; k < 50; k++ {
				x[512+rnd.Intn(len(x)-512)] = byte(rnd.Intn(256))
			}
		}
		ToHTML(x) // must not panic (errors are fine)
	}
}

func TestHandler(t *testing.T) {
	b := readFixture(t, formPath)
	mux := http.NewServeMux()
	Register(mux)
	rec := httptest.NewRecorder()
	mux.ServeHTTP(rec, httptest.NewRequest("POST", "/api/doc/to-html", bytes.NewReader(b)))
	if rec.Code != 200 {
		t.Fatalf("status %d: %s", rec.Code, rec.Body.String())
	}
	var res struct{ HTML, Title string }
	if err := json.Unmarshal(rec.Body.Bytes(), &res); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.HTML, "Geervan Thanra") {
		t.Error("handler output missing text")
	}
	rec = httptest.NewRecorder()
	mux.ServeHTTP(rec, httptest.NewRequest("POST", "/api/doc/to-html", strings.NewReader("junk")))
	if rec.Code != 400 {
		t.Errorf("junk: status %d", rec.Code)
	}
}

func TestDocxRenamedToDoc(t *testing.T) {
	b := readFixture(t, "../../USULAN PENELITIAN S1-geervan thanra.docx")
	h, _, err := ToHTML(b)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(plain(h), "PENDAHULUAN") {
		t.Error("docx fallback output missing text")
	}
}
