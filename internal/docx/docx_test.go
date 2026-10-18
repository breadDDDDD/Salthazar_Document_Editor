package docx

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"image"
	"image/color"
	"image/png"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

func tinyPNG(t *testing.T) string {
	img := image.NewRGBA(image.Rect(0, 0, 4, 2))
	img.Set(0, 0, color.RGBA{255, 0, 0, 255})
	var b bytes.Buffer
	if err := png.Encode(&b, img); err != nil {
		t.Fatal(err)
	}
	return "data:image/png;base64," + base64.StdEncoding.EncodeToString(b.Bytes())
}

func validateZip(t *testing.T, data []byte) map[string]string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("zip: %v", err)
	}
	parts := map[string]string{}
	for _, f := range zr.File {
		rc, _ := f.Open()
		b, _ := io.ReadAll(rc)
		rc.Close()
		parts[f.Name] = string(b)
		if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
			d := xml.NewDecoder(bytes.NewReader(b))
			for {
				_, err := d.Token()
				if err == io.EOF {
					break
				}
				if err != nil {
					t.Fatalf("part %s is not well-formed XML: %v", f.Name, err)
				}
			}
		}
	}
	for _, need := range []string{"[Content_Types].xml", "_rels/.rels", "word/document.xml",
		"word/_rels/document.xml.rels", "word/styles.xml", "word/numbering.xml", "docProps/core.xml"} {
		if _, ok := parts[need]; !ok {
			t.Errorf("missing part %s", need)
		}
	}
	return parts
}

func TestRoundTrip(t *testing.T) {
	src := `<h1>My Title</h1>
<p style="text-align:center">Hello <strong>bold</strong> and <em>italic</em> <u>under</u> <s>gone</s> x<sup>2</sup> H<sub>2</sub>O
<span style="color: rgb(255, 0, 0); background-color:#ffff00; font-size:18pt; font-family:'Times New Roman'">styled</span>
&amp; &lt;tag&gt; <a href="https://example.com/?a=1&amp;b=2">a link</a></p>
<h2>Sub</h2>
<ul><li>one</li><li>two<ul><li>nested</li></ul></li></ul>
<ol><li>first</li><li>second</li></ol>
<table><tr><th>A</th><th>B</th></tr><tr><td colspan="2">wide</td></tr></table>
<p>pic <img src="` + tinyPNG(t) + `" width="40" height="20" alt="red box"></p>
<hr class="page-break">
<p>tab<span style="white-space:pre">	</span>end<br>line2</p>
<blockquote>quoted</blockquote>
<p></p>`
	out, err := FromHTML(src)
	if err != nil {
		t.Fatal(err)
	}
	parts := validateZip(t, out)
	if _, ok := parts["word/media/image1.png"]; !ok {
		t.Error("image not embedded")
	}
	doc := parts["word/document.xml"]
	for _, want := range []string{`w:val="Heading1"`, `<w:b/>`, `<w:numId`, `<w:tbl>`, `<w:gridSpan w:val="2"/>`,
		`<w:drawing>`, `w:type="page"`, `<w:tab/>`, `w:fill="FFFF00"`, `<w:color w:val="FF0000"/>`, `<w:sz w:val="36"/>`,
		`w:jc w:val="center"`, `<w:hyperlink r:id=`, `<w:ind w:left="720"/>`, `<w:vertAlign w:val="superscript"/>`} {
		if !strings.Contains(doc, want) {
			t.Errorf("document.xml missing %s", want)
		}
	}
	if !strings.Contains(parts["word/_rels/document.xml.rels"], `https://example.com/?a=1&amp;b=2`) {
		t.Error("hyperlink rel missing")
	}

	h, title, err := ToHTML(out)
	if err != nil {
		t.Fatal(err)
	}
	if title != "My Title" {
		t.Errorf("title = %q", title)
	}
	for _, want := range []string{"<h1>My Title</h1>", "<h2>Sub</h2>", "<strong>bold</strong>", "<em>italic</em>",
		"<u>under</u>", "<s>gone</s>", "<sup>2</sup>", "<sub>2</sub>", "color:#ff0000", "background-color:#ffff00",
		"font-size:18pt", "Times New Roman", "&amp; &lt;tag&gt;", `<a href="https://example.com/?a=1&amp;b=2">`,
		"<ul><li>one</li><li>two<ul><li>nested</li></ul></li></ul>", "<ol><li>first</li><li>second</li></ol>",
		`colspan="2"`, "<td>", `<img src="data:image/png;base64,`, `width="40" height="20"`, `alt="red box"`,
		`<hr class="page-break">`, "\t", "end<br>line2", "text-align:center", "margin-left:36pt", "<p><br></p>"} {
		if !strings.Contains(h, want) {
			t.Errorf("html missing %q\n%s", want, h)
		}
	}
	// second round trip must still be valid and stable in structure
	out2, err := FromHTML(h)
	if err != nil {
		t.Fatal(err)
	}
	validateZip(t, out2)
	h2, _, err := ToHTML(out2)
	if err != nil {
		t.Fatal(err)
	}
	if h2 != h {
		t.Errorf("second round trip differs:\n%s\n---\n%s", h, h2)
	}
}

func TestTolerantHTML(t *testing.T) {
	src := `<!DOCTYPE html><html><head><title>T &amp; x</title><style>p{color:red}</style></head><body>
<p>unclosed <b>bold <i>both</b> after<p>second para<div>div text</div>
<ul><li>a<li>b</ul><table><tr><td>1<td>2<tr><td>3</table><font color="#00f" size="5" face="Arial">font</font>
<p>attr <a href=https://x.org>noquote</a> <img src="https://remote/img.png" alt="remote"> &nbsp; &#169; &unknown; <script>alert(1)</script> 1 < 2
<!-- comment --></body></html>`
	out, err := FromHTML(src)
	if err != nil {
		t.Fatal(err)
	}
	validateZip(t, out)
	h, title, err := ToHTML(out)
	if err != nil {
		t.Fatal(err)
	}
	if title != "T & x" {
		t.Errorf("title = %q", title)
	}
	for _, want := range []string{"second para", "div text", "<li>a</li><li>b</li>", "<td><p>1</p></td><td><p>2</p></td>",
		"color:#0000ff", "font-size:18pt", "Arial", `href="https://x.org"`, "[remote]", "©", "1 &lt; 2"} {
		if !strings.Contains(h, want) {
			t.Errorf("html missing %q\n%s", want, h)
		}
	}
	if strings.Contains(h, "alert") || strings.Contains(h, "color:red") {
		t.Error("script/style content leaked")
	}
}

func TestEmptyAndInvalid(t *testing.T) {
	out, err := FromHTML("")
	if err != nil {
		t.Fatal(err)
	}
	validateZip(t, out)
	if _, _, err := ToHTML(out); err != nil {
		t.Fatal(err)
	}
	if _, _, err := ToHTML([]byte("not a zip")); err == nil {
		t.Error("expected error for invalid docx")
	}
}

// TestWordStyleInput covers constructs Word writes but we do not emit:
// vMerge, highlight names, Title style, bullets via style numbering.
func TestWordStyleInput(t *testing.T) {
	doc := `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Doc Title</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:highlight w:val="yellow"/><w:b w:val="0"/></w:rPr><w:t xml:space="preserve">hi </w:t></w:r><w:r><w:t>there</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>r1</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>r2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sectPr/></w:body></w:document>`
	styles := `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>`
	var b bytes.Buffer
	zw := zip.NewWriter(&b)
	for name, c := range map[string]string{"word/document.xml": doc, "word/styles.xml": styles} {
		f, _ := zw.Create(name)
		f.Write([]byte(c))
	}
	zw.Close()
	h, title, err := ToHTML(b.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	if title != "Doc Title" {
		t.Errorf("title %q", title)
	}
	for _, want := range []string{"<h1>Doc Title</h1>", `<span style="background-color:#ffff00">hi </span>there`, `rowspan="2"`} {
		if !strings.Contains(h, want) {
			t.Errorf("html missing %q\n%s", want, h)
		}
	}
}

func TestHandlers(t *testing.T) {
	mux := http.NewServeMux()
	Register(mux)
	srv := httptest.NewServer(mux)
	defer srv.Close()
	resp, err := http.Post(srv.URL+"/api/docx/from-html", "application/json", strings.NewReader(`{"html":"<p>Hi</p>"}`))
	if err != nil {
		t.Fatal(err)
	}
	body, _ := io.ReadAll(resp.Body)
	resp.Body.Close()
	if resp.StatusCode != 200 || !strings.Contains(resp.Header.Get("Content-Type"), "wordprocessingml") {
		t.Fatalf("from-html: %d %s", resp.StatusCode, body)
	}
	resp, err = http.Post(srv.URL+"/api/docx/to-html", "application/octet-stream", bytes.NewReader(body))
	if err != nil {
		t.Fatal(err)
	}
	j, _ := io.ReadAll(resp.Body)
	resp.Body.Close()
	if resp.StatusCode != 200 || !strings.Contains(string(j), `<p>Hi</p>`) {
		t.Fatalf("to-html: %d %s", resp.StatusCode, j)
	}
	resp, _ = http.Post(srv.URL+"/api/docx/to-html", "application/octet-stream", strings.NewReader("junk"))
	if resp.StatusCode != 400 {
		t.Errorf("junk docx status %d", resp.StatusCode)
	}
	resp.Body.Close()
}
