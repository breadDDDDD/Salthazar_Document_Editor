package docx

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

const sampleDocx = "../../USULAN PENELITIAN S1-geervan thanra.docx"

func loadSample(t *testing.T) []byte {
	t.Helper()
	b, err := os.ReadFile(sampleDocx)
	if err != nil {
		t.Skipf("sample document not available: %v", err)
	}
	return b
}

// dumpHTML writes HTML to $DOCX_DUMP_DIR (if set) for manual inspection.
func dumpHTML(t *testing.T, name, h string) {
	if dir := os.Getenv("DOCX_DUMP_DIR"); dir != "" {
		os.WriteFile(filepath.Join(dir, name), []byte(h), 0o644)
	}
}

func TestSampleDocxFidelity(t *testing.T) {
	data := loadSample(t)
	h, title, err := ToHTML(data)
	if err != nil {
		t.Fatal(err)
	}
	dumpHTML(t, "sample.html", h)
	t.Logf("title=%q html=%d bytes", title, len(h))

	// Page setup from sectPr: A4 (11906 x 16838 twips), margins 2268/1701/1701/2268.
	for _, want := range []string{`class="docx-doc"`, `data-page-w="11906"`, `data-page-h="16838"`,
		`data-margin-top="2268"`, `data-margin-left="2268"`, `data-margin-right="1701"`,
		`data-font="Times New Roman"`, `data-font-size="12"`} {
		if !strings.Contains(h, want) {
			t.Errorf("missing page setup %s", want)
		}
	}
	// Normal style is Times New Roman 12pt: body paragraphs must say so explicitly.
	if n := strings.Count(h, "font-family:&#39;Times New Roman&#39;, serif;font-size:12pt"); n < 50 {
		t.Errorf("only %d paragraphs carry Times New Roman 12pt", n)
	}
	if strings.Contains(h, "Calibri") && strings.Count(h, "Calibri") > strings.Count(h, "Times New Roman") {
		t.Errorf("Calibri dominates output; style inheritance not resolved")
	}
	// Heading 1: centered, bold, 14pt (sz 28) Times New Roman.
	h1 := regexp.MustCompile(`<h1 style="([^"]*)">[^<]*BAB I PENDAHULUAN`).FindStringSubmatch(h)
	if h1 == nil {
		t.Fatalf("Heading 1 'BAB I PENDAHULUAN' not found")
	}
	for _, want := range []string{"text-align:center", "font-size:14pt", "font-weight:bold", "Times New Roman"} {
		if !strings.Contains(html2text(h1[1]), want) {
			t.Errorf("heading 1 style %q missing %s", h1[1], want)
		}
	}
	// Heading 2 is numbered through numbering.xml (ilvl 1) -> label like "1.1".
	if !regexp.MustCompile(`<h2 [^>]*>\d+\.\d+\x{00a0} Latar Belakang`).MatchString(h) {
		t.Errorf("numbered heading 2 'x.y Latar Belakang' not found")
	}
	// 1.5 line spacing (line=360 auto) and justified text from direct formatting.
	if !strings.Contains(h, "line-height:1.725") || !strings.Contains(h, "text-align:justify") {
		t.Errorf("line spacing / justification not preserved")
	}
	// First-line indent of 426 twips = 21.3pt.
	if !strings.Contains(h, "text-indent:21.3pt") {
		t.Errorf("first-line indent missing")
	}
	// Lists, tables (TableGrid borders), images, footer.
	for _, want := range []string{"<ol", "<li", "<table", "border-top:0.5pt solid #000000", "<img src=\"data:image/png;base64,",
		`<footer class="docx-footer">`} {
		if !strings.Contains(h, want) {
			t.Errorf("html missing %q", want)
		}
	}
	// The EMF cover picture must be converted (no placeholder).
	if strings.Contains(h, "[Picture") {
		t.Errorf("a picture could not be converted")
	}
}

func html2text(s string) string { return strings.ReplaceAll(s, "&#39;", "'") }

func TestSampleDocxRoundTrip(t *testing.T) {
	data := loadSample(t)
	h1, _, err := ToHTML(data)
	if err != nil {
		t.Fatal(err)
	}
	out, err := FromHTML(h1)
	if err != nil {
		t.Fatal(err)
	}
	parts := validateZip(t, out)
	doc := parts["word/document.xml"]
	for _, want := range []string{`<w:pgSz w:w="11906" w:h="16838"`, `w:left="2268"`, `w:ascii="Times New Roman"`,
		`<w:sz w:val="24"/>`, `w:line="360" w:lineRule="auto"`, `<w:jc w:val="both"/>`, `w:firstLine="426"`} {
		if !strings.Contains(doc, want) {
			t.Errorf("document.xml missing %s", want)
		}
	}
	if !strings.Contains(parts["word/styles.xml"], `w:ascii="Times New Roman"`) {
		t.Errorf("styles.xml default font not preserved")
	}
	if _, ok := parts["word/footer1.xml"]; !ok {
		t.Errorf("footer not written")
	}
	h2, _, err := ToHTML(out)
	if err != nil {
		t.Fatal(err)
	}
	dumpHTML(t, "sample-roundtrip.html", h2)
	// the effective formatting of every character must survive the round trip
	f1, f2 := charFormats(h1), charFormats(h2)
	total, diff := 0, 0
	for k, v := range f1 {
		total += v
		if d := v - f2[k]; d > 0 {
			diff += d
		}
	}
	if total == 0 || diff*100 > total { // allow 1% drift
		for k, v := range f1 {
			if v != f2[k] {
				t.Logf("%-60s %6d -> %6d", k, v, f2[k])
			}
		}
		t.Errorf("round trip changed the formatting of %d of %d characters", diff, total)
	}
}

// charFormats computes, for every text character, its effective font, size,
// weight, style and colour by cascading inline CSS and tags, and counts them.
func charFormats(h string) map[string]int {
	root, _ := parseHTML(h)
	out := map[string]int{}
	type st struct{ font, size, weight, style, color, align string }
	var walk func(n *hnode, s st)
	walk = func(n *hnode, s st) {
		if n.tag == "" {
			c := 0
			for _, r := range n.text {
				if r > ' ' && r != 0xa0 {
					c++
				}
			}
			if c > 0 {
				out[fmt.Sprintf("%s|%s|%s|%s|%s|%s", s.font, s.size, s.weight, s.style, s.color, s.align)] += c
			}
			return
		}
		switch n.tag {
		case "strong", "b":
			s.weight = "bold"
		case "em", "i":
			s.style = "italic"
		}
		css := parseCSS(n.attr("style"))
		if f := css["font-family"]; f != "" {
			s.font = strings.Trim(strings.TrimSpace(strings.Split(f, ",")[0]), `"'`)
		}
		if v := css["font-size"]; v != "" {
			s.size = v
		}
		if v := css["font-weight"]; v != "" {
			s.weight = v
		}
		if v := css["font-style"]; v != "" {
			s.style = v
		}
		if v := css["color"]; v != "" && v != "inherit" {
			s.color = v
		} else if v == "inherit" {
			s.color = ""
		}
		if v := css["text-align"]; v != "" {
			s.align = v
		}
		for _, k := range n.kids {
			walk(k, s)
		}
	}
	walk(root, st{weight: "normal", style: "normal"})
	return out
}
