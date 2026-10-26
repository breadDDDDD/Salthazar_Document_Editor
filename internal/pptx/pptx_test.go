package pptx

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const testFile = "../../Summary of research.pptx"

func loadTestFile(t *testing.T) []byte {
	t.Helper()
	b, err := os.ReadFile(testFile)
	if err != nil {
		t.Skipf("test file not available: %v", err)
	}
	return b
}

func mustModel(t *testing.T, b []byte) *Model {
	t.Helper()
	m, _, err := ToModel(b)
	if err != nil {
		t.Fatalf("ToModel: %v", err)
	}
	return m
}

// clone deep-copies a model through JSON, like the browser round trip.
func clone(t *testing.T, m *Model) *Model {
	t.Helper()
	b, err := json.Marshal(m)
	if err != nil {
		t.Fatal(err)
	}
	var out Model
	if err := json.Unmarshal(b, &out); err != nil {
		t.Fatal(err)
	}
	return &out
}

func slideTitle(s *Slide) string {
	for _, sh := range s.Shapes {
		if sh.Ph != nil && (*sh.Ph == "title" || *sh.Ph == "ctrTitle") {
			return strings.TrimSpace(shapeText(sh))
		}
	}
	return ""
}

func allXMLParse(t *testing.T, pkg []byte) {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("output is not a zip: %v", err)
	}
	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, _ := f.Open()
		d := xml.NewDecoder(rc)
		for {
			_, err := d.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				t.Errorf("%s: invalid XML: %v", f.Name, err)
				break
			}
		}
		rc.Close()
	}
}

func TestReadSummary(t *testing.T) {
	b := loadTestFile(t)
	m, title, err := ToModel(b)
	if err != nil {
		t.Fatal(err)
	}
	if len(m.Slides) != 24 {
		t.Fatalf("slides = %d, want 24", len(m.Slides))
	}
	if m.Size.CX != 9144000 || m.Size.CY != 6858000 {
		t.Errorf("size = %+v", m.Size)
	}
	if !strings.Contains(title, "Energy Bandgap") {
		t.Errorf("title = %q", title)
	}
	if got := slideTitle(m.Slides[0]); !strings.HasPrefix(got, "Energy Bandgap on Double Halide Perovskites") {
		t.Errorf("slide 1 title = %q", got)
	}
	if got := slideTitle(m.Slides[2]); got != "Dataset" {
		t.Errorf("slide 3 title = %q", got)
	}
	if m.Theme.MinorFont != "Calibri" || m.Theme.MajorFont != "Calibri Light" {
		t.Errorf("theme fonts = %q/%q", m.Theme.MajorFont, m.Theme.MinorFont)
	}
	if len(m.Layouts) != 4 {
		t.Errorf("layouts = %d", len(m.Layouts))
	}
	images, tables, opaque := 0, 0, 0
	for si, s := range m.Slides {
		for _, sh := range s.Shapes {
			if sh.Kind != "opaque" && (sh.W <= 0 || sh.H <= 0) {
				t.Errorf("slide %d shape %s (%s) has zero geometry", si+1, sh.ID, sh.Name)
			}
			switch sh.Kind {
			case "image":
				images++
				if !strings.HasPrefix(sh.Src, "media:") || !strings.HasPrefix(m.Media[strings.TrimPrefix(sh.Src, "media:")], "data:image/") {
					t.Errorf("slide %d image %s not decoded: %q", si+1, sh.ID, sh.Src)
				}
			case "table":
				tables++
				if len(sh.Rows) == 0 || len(sh.ColW) == 0 {
					t.Errorf("slide %d table empty", si+1)
				}
				for _, row := range sh.Rows {
					if len(row) != len(sh.ColW) {
						t.Errorf("slide %d table row has %d cells, grid %d", si+1, len(row), len(sh.ColW))
					}
				}
			case "opaque":
				opaque++
				if sh.XML == "" {
					t.Errorf("opaque without xml")
				}
			}
			for _, p := range sh.Paras {
				for _, r := range p.Runs {
					if !strings.HasPrefix(r.Color, "#") || len(r.Color) != 7 {
						t.Errorf("slide %d run color %q not resolved", si+1, r.Color)
					}
					if r.Font == "" || strings.HasPrefix(r.Font, "+") {
						t.Errorf("slide %d run font %q not resolved", si+1, r.Font)
					}
					if r.Size <= 0 {
						t.Errorf("slide %d run size %v", si+1, r.Size)
					}
				}
			}
		}
	}
	if images < 15 || tables != 8 || opaque < 2 {
		t.Errorf("images=%d tables=%d opaque=%d", images, tables, opaque)
	}
	// Title placeholder on the custom layout inherits accent1 lumMod 50% and Segoe UI Light? (slide sets +mn-lt only on body)
	s3 := m.Slides[2]
	for _, sh := range s3.Shapes {
		if sh.Ph != nil && *sh.Ph == "title" {
			r := sh.Paras[0].Runs[0]
			if r.Size != 36 || r.Font != "Segoe UI Light" || r.Color != "#1F4E79" {
				t.Errorf("slide 3 title run = %+v", r)
			}
		}
	}
	// layout decor carries the full-slide background picture
	if len(s3.Decor) == 0 || s3.Decor[0].Kind != "image" {
		t.Errorf("slide 3 decor = %d shapes", len(s3.Decor))
	}
	// notes
	notes := 0
	for _, s := range m.Slides {
		if s.Notes != "" {
			notes++
		}
	}
	if notes == 0 {
		t.Errorf("no notes found")
	}
}

func TestRoundTrip(t *testing.T) {
	b := loadTestFile(t)
	m := mustModel(t, b)
	out, err := FromModel(clone(t, m), b)
	if err != nil {
		t.Fatal(err)
	}
	allXMLParse(t, out)
	writeOut(t, "roundtrip.pptx", out)
	m2 := mustModel(t, out)
	compareModels(t, m, m2)
}

// TestRoundTripRegenerated forces every shape to be regenerated (no verbatim reuse).
func TestRoundTripRegenerated(t *testing.T) {
	b := loadTestFile(t)
	m := mustModel(t, b)
	c := clone(t, m)
	for _, s := range c.Slides {
		for _, sh := range s.Shapes {
			if sh.Kind != "opaque" {
				sh.Name += " " // makes the JSON differ from the original
			}
		}
	}
	out, err := FromModel(c, b)
	if err != nil {
		t.Fatal(err)
	}
	allXMLParse(t, out)
	writeOut(t, "regenerated.pptx", out)
	m2 := mustModel(t, out)
	compareModels(t, c, m2)
}

func compareModels(t *testing.T, a, b *Model) {
	t.Helper()
	if len(a.Slides) != len(b.Slides) {
		t.Fatalf("slide count %d vs %d", len(a.Slides), len(b.Slides))
	}
	for i := range a.Slides {
		sa, sb := a.Slides[i], b.Slides[i]
		if len(sa.Shapes) != len(sb.Shapes) {
			t.Errorf("slide %d: %d vs %d shapes", i+1, len(sa.Shapes), len(sb.Shapes))
			continue
		}
		if sa.Notes != sb.Notes {
			t.Errorf("slide %d notes differ", i+1)
		}
		if !jsonEq(sa.Bg, sb.Bg) {
			t.Errorf("slide %d bg %v vs %v", i+1, sa.Bg, sb.Bg)
		}
		for j := range sa.Shapes {
			x, y := sa.Shapes[j], sb.Shapes[j]
			if x.Kind == "opaque" {
				if y.Kind != "opaque" || x.X != y.X || x.W != y.W {
					t.Errorf("slide %d opaque shape %d differs", i+1, j)
				}
				continue
			}
			x2, y2 := *x, *y
			x2.ID, y2.ID = "", ""
			if !jsonEq(x2, y2) {
				ja, _ := json.Marshal(x2)
				jb, _ := json.Marshal(y2)
				t.Errorf("slide %d shape %d (%s) differs:\n%s\n%s", i+1, j, x.Name, trunc(diffHint(string(ja), string(jb))), "")
				_ = jb
			}
		}
	}
}

func diffHint(a, b string) string {
	i := 0
	for i < len(a) && i < len(b) && a[i] == b[i] {
		i++
	}
	s := i - 120
	if s < 0 {
		s = 0
	}
	ea, eb := i+120, i+120
	if ea > len(a) {
		ea = len(a)
	}
	if eb > len(b) {
		eb = len(b)
	}
	return "A: ..." + a[s:ea] + "\nB: ..." + b[s:eb]
}

func trunc(s string) string { return s }

func writeOut(t *testing.T, name string, data []byte) {
	dir := os.Getenv("PPTX_TEST_OUT")
	if dir == "" {
		dir = t.TempDir()
	}
	os.MkdirAll(dir, 0o755)
	if err := os.WriteFile(filepath.Join(dir, name), data, 0o644); err != nil {
		t.Fatal(err)
	}
}

func TestEditOperations(t *testing.T) {
	b := loadTestFile(t)
	m := mustModel(t, b)
	c := clone(t, m)
	// edit title text of slide 3
	for _, sh := range c.Slides[2].Shapes {
		if sh.Ph != nil && *sh.Ph == "title" {
			sh.Paras[0].Runs[0].Text = "Dataset (edited)"
			sh.Paras[0].Runs[0].B = true
		}
	}
	// delete slide 2, duplicate slide 1 at the end, move slide 24 to the front
	s1 := c.Slides[0]
	dup := clone(t, &Model{Slides: []*Slide{s1}}).Slides[0]
	dup.ID = "dup"
	last := c.Slides[len(c.Slides)-1]
	rest := append([]*Slide{}, c.Slides[2:len(c.Slides)-1]...)
	c.Slides = append(append([]*Slide{last, s1}, rest...), dup)
	// add a new slide with a text box, a shape and an image
	png := "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
	title := "title"
	fill := "#FF0000"
	c.Slides = append(c.Slides, &Slide{ID: "new1", Layout: c.Layouts[3].ID, Notes: "New notes\nline 2", Shapes: []*Shape{
		{ID: "n1", Kind: "text", Ph: &title, Paras: []*Para{{Align: "l", Runs: []*Run{{Text: "Brand new slide", Size: 36}}}}},
		{ID: "n2", Kind: "text", X: 914400, Y: 2000000, W: 4000000, H: 800000, Paras: []*Para{{Align: "ctr", Runs: []*Run{{Text: "Hello\nworld", Size: 20, Color: "#123456", Font: "Arial", Link: strp("https://example.com")}}}}},
		{ID: "n3", Kind: "shape", Geom: "ellipse", X: 5000000, Y: 2000000, W: 1000000, H: 1000000, Fill: &fill, Line: &Line{Color: "#000000", Width: 12700, Dash: "dash"}},
		{ID: "n4", Kind: "image", X: 100000, Y: 100000, W: 500000, H: 500000, Src: png},
	}})
	out, err := FromModel(c, b)
	if err != nil {
		t.Fatal(err)
	}
	allXMLParse(t, out)
	writeOut(t, "edited.pptx", out)
	m2 := mustModel(t, out)
	if len(m2.Slides) != 25 {
		t.Fatalf("slides = %d, want 25", len(m2.Slides))
	}
	if got, want := slideTitle(m2.Slides[0]), slideTitle(m.Slides[23]); got != want {
		t.Errorf("first slide title %q, want %q", got, want)
	}
	if got, want := slideTitle(m2.Slides[1]), slideTitle(m.Slides[0]); got != want {
		t.Errorf("second slide title %q want %q", got, want)
	}
	if got := slideTitle(m2.Slides[2]); got != "Dataset (edited)" {
		t.Errorf("edited title = %q", got)
	}
	for _, sh := range m2.Slides[2].Shapes {
		if sh.Ph != nil && *sh.Ph == "title" && !sh.Paras[0].Runs[0].B {
			t.Errorf("bold lost")
		}
	}
	if got, want := slideTitle(m2.Slides[23]), slideTitle(m.Slides[0]); got != want {
		t.Errorf("duplicate title %q want %q", got, want)
	}
	ns := m2.Slides[24]
	if slideTitle(ns) != "Brand new slide" || ns.Notes != "New notes\nline 2" {
		t.Errorf("new slide: title %q notes %q", slideTitle(ns), ns.Notes)
	}
	var kinds []string
	for _, sh := range ns.Shapes {
		kinds = append(kinds, sh.Kind)
	}
	if strings.Join(kinds, ",") != "text,text,shape,image" {
		t.Errorf("new slide kinds = %v", kinds)
	}
	tb := ns.Shapes[1]
	if len(tb.Paras[0].Runs) != 3 || tb.Paras[0].Runs[1].Text != "\n" || tb.Paras[0].Runs[0].Color != "#123456" || tb.Paras[0].Runs[0].Font != "Arial" || tb.Paras[0].Runs[0].Link == nil {
		js, _ := json.Marshal(tb.Paras)
		t.Errorf("text box runs = %s", js)
	}
	el := ns.Shapes[2]
	if el.Geom != "ellipse" || el.Fill == nil || *el.Fill != "#FF0000" || el.Line == nil || el.Line.Dash != "dash" {
		t.Errorf("ellipse = %+v", el)
	}
	// removed slide 2 must not be referenced; titles of all slides unique set
	for _, s := range m2.Slides {
		if slideTitle(s) == slideTitle(m.Slides[1]) && slideTitle(m.Slides[1]) != "" {
			t.Errorf("deleted slide still present: %q", slideTitle(s))
		}
	}
}

func TestBlank(t *testing.T) {
	m, err := BlankModel()
	if err != nil {
		t.Fatal(err)
	}
	if len(m.Slides) != 1 || len(m.Layouts) != 3 || m.Size.CX != 12192000 {
		t.Fatalf("blank model: %d slides %d layouts", len(m.Slides), len(m.Layouts))
	}
	sh := m.Slides[0].Shapes
	if len(sh) != 2 || sh[0].W == 0 || sh[1].W == 0 {
		t.Fatalf("blank placeholders: %+v", sh)
	}
	if len(m.Layouts[1].Placeholders) != 2 {
		t.Errorf("layout placeholders = %d", len(m.Layouts[1].Placeholders))
	}
	m.Slides[0].Shapes[0].Paras[0].Runs = []*Run{{Text: "Hello", Size: 60}}
	m.Slides = append(m.Slides, &Slide{ID: "x", Layout: m.Layouts[1].ID, Notes: "n", Shapes: []*Shape{}})
	out, err := FromModel(m, nil)
	if err != nil {
		t.Fatal(err)
	}
	allXMLParse(t, out)
	writeOut(t, "blank.pptx", out)
	m2 := mustModel(t, out)
	if slideTitle(m2.Slides[0]) != "Hello" || m2.Slides[1].Notes != "n" {
		t.Errorf("blank round trip: %q %q", slideTitle(m2.Slides[0]), m2.Slides[1].Notes)
	}
}
