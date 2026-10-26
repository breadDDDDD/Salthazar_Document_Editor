package pptx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"path"
	"sort"
	"strings"
)

const xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"

// Relationship type suffixes.
const (
	rtSlide       = relDoc + "/slide"
	rtSlideLayout = relDoc + "/slideLayout"
	rtSlideMaster = relDoc + "/slideMaster"
	rtTheme       = relDoc + "/theme"
	rtImage       = relDoc + "/image"
	rtHyperlink   = relDoc + "/hyperlink"
	rtNotesSlide  = relDoc + "/notesSlide"
	rtNotesMaster = relDoc + "/notesMaster"
	rtTableStyles = relDoc + "/tableStyles"
	rtOfficeDoc   = relDoc + "/officeDocument"
)

const (
	ctSlide       = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ctNotesSlide  = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
	ctNotesMaster = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
	ctTheme       = "application/vnd.openxmlformats-officedocument.theme+xml"
)

// Rel is one package relationship. Target is an absolute part name (no
// leading slash) unless External.
type Rel struct {
	ID       string `json:"-"`
	Type     string `json:"type"`
	Target   string `json:"target"`
	External bool   `json:"external,omitempty"`
}

// Pkg is an opened OPC package.
type Pkg struct {
	Files map[string][]byte
	Order []string
	dom   map[string]*Node
	rels  map[string]map[string]Rel
}

func newPkg() *Pkg {
	return &Pkg{Files: map[string][]byte{}, dom: map[string]*Node{}, rels: map[string]map[string]Rel{}}
}

func openPkg(b []byte) (*Pkg, error) {
	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return nil, fmt.Errorf("not a valid .pptx (zip) file: %v", err)
	}
	p := newPkg()
	for _, f := range zr.File {
		if f.FileInfo().IsDir() {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(io.LimitReader(rc, 512<<20))
		rc.Close()
		if err != nil {
			return nil, fmt.Errorf("cannot read %s: %v", f.Name, err)
		}
		name := strings.TrimPrefix(f.Name, "/")
		if _, dup := p.Files[name]; !dup {
			p.Order = append(p.Order, name)
		}
		p.Files[name] = data
	}
	return p, nil
}

// findFile does a case-insensitive lookup (some writers vary case).
func (p *Pkg) findFile(name string) (string, bool) {
	if _, ok := p.Files[name]; ok {
		return name, true
	}
	ln := strings.ToLower(name)
	for k := range p.Files {
		if strings.ToLower(k) == ln {
			return k, true
		}
	}
	return name, false
}

// XML returns the parsed DOM of a part (cached), or nil.
func (p *Pkg) XML(name string) *Node {
	if name == "" {
		return nil
	}
	name, ok := p.findFile(name)
	if !ok {
		return nil
	}
	if n, ok := p.dom[name]; ok {
		return n
	}
	n, err := parseXML(p.Files[name])
	if err != nil {
		n = nil
	}
	p.dom[name] = n
	return n
}

func relsPartName(part string) string {
	dir, base := path.Split(part)
	return dir + "_rels/" + base + ".rels"
}

// Rels returns relationships of a part keyed by id.
func (p *Pkg) Rels(part string) map[string]Rel {
	if r, ok := p.rels[part]; ok {
		return r
	}
	out := map[string]Rel{}
	root := p.XML(relsPartName(part))
	if root != nil {
		base := path.Dir(part)
		if part == "" {
			base = ""
		}
		for _, e := range root.Elems() {
			if e.Local != "Relationship" {
				continue
			}
			r := Rel{ID: e.A("Id"), Type: e.A("Type"), Target: e.A("Target")}
			if strings.EqualFold(e.A("TargetMode"), "External") {
				r.External = true
			} else {
				r.Target = resolveTarget(base, r.Target)
			}
			out[r.ID] = r
		}
	}
	p.rels[part] = out
	return out
}

func resolveTarget(base, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(path.Clean(target), "/")
	}
	if base == "" || base == "." {
		return path.Clean(target)
	}
	return strings.TrimPrefix(path.Clean(base+"/"+target), "/")
}

// relTarget makes a relative target from part to an absolute part name.
func relTarget(fromPart, to string) string {
	fromDir := strings.Split(path.Dir(fromPart), "/")
	if path.Dir(fromPart) == "." {
		fromDir = nil
	}
	toParts := strings.Split(to, "/")
	i := 0
	for i < len(fromDir) && i < len(toParts)-1 && fromDir[i] == toParts[i] {
		i++
	}
	var sb []string
	for j := i; j < len(fromDir); j++ {
		sb = append(sb, "..")
	}
	sb = append(sb, toParts[i:]...)
	return strings.Join(sb, "/")
}

// relByType returns the first relationship of the given type, sorted by id for determinism.
func relByType(rels map[string]Rel, typ string) (Rel, bool) {
	ids := sortedRelIDs(rels)
	for _, id := range ids {
		if sameRelKind(rels[id].Type, typ) {
			return rels[id], true
		}
	}
	return Rel{}, false
}

// sameRelKind compares relationship types ignoring the strict/transitional base URI.
func sameRelKind(a, b string) bool {
	return a[strings.LastIndex(a, "/"):] == b[strings.LastIndex(b, "/"):]
}

func sortedRelIDs(rels map[string]Rel) []string {
	ids := make([]string, 0, len(rels))
	for id := range rels {
		ids = append(ids, id)
	}
	sort.Slice(ids, func(i, j int) bool {
		a, b := ids[i], ids[j]
		if len(a) != len(b) {
			return len(a) < len(b)
		}
		return a < b
	})
	return ids
}

// writeRels serializes a relationships part.
func writeRels(rels []Rel, fromPart string) []byte {
	var sb strings.Builder
	sb.WriteString(xmlDecl)
	sb.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for _, r := range rels {
		t := r.Target
		mode := ""
		if r.External {
			mode = ` TargetMode="External"`
		} else {
			t = relTarget(fromPart, t)
		}
		fmt.Fprintf(&sb, `<Relationship Id="%s" Type="%s" Target="%s"%s/>`, esc(r.ID), esc(r.Type), esc(t), mode)
	}
	sb.WriteString(`</Relationships>`)
	return []byte(sb.String())
}

// mainPart returns the presentation part name.
func (p *Pkg) mainPart() string {
	rels := p.Rels("")
	for _, id := range sortedRelIDs(rels) {
		if sameRelKind(rels[id].Type, rtOfficeDoc) {
			return rels[id].Target
		}
	}
	return "ppt/presentation.xml"
}

func (p *Pkg) zip() ([]byte, error) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	names := []string{"[Content_Types].xml"}
	seen := map[string]bool{"[Content_Types].xml": true}
	for _, n := range p.Order {
		if _, ok := p.Files[n]; ok && !seen[n] {
			names = append(names, n)
			seen[n] = true
		}
	}
	var rest []string
	for n := range p.Files {
		if !seen[n] {
			rest = append(rest, n)
		}
	}
	sort.Strings(rest)
	names = append(names, rest...)
	for _, n := range names {
		data, ok := p.Files[n]
		if !ok {
			continue
		}
		method := zip.Deflate
		w, err := zw.CreateHeader(&zip.FileHeader{Name: n, Method: method})
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(data); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func mimeForExt(ext string) string {
	switch strings.ToLower(strings.TrimPrefix(ext, ".")) {
	case "png":
		return "image/png"
	case "jpg", "jpeg", "jpe", "jfif":
		return "image/jpeg"
	case "gif":
		return "image/gif"
	case "bmp":
		return "image/bmp"
	case "svg":
		return "image/svg+xml"
	case "webp":
		return "image/webp"
	case "tif", "tiff":
		return "image/tiff"
	case "emf":
		return "image/x-emf"
	case "wmf":
		return "image/x-wmf"
	}
	return ""
}

// browserImage reports whether a browser can display the format.
func browserImage(ext string) bool {
	switch strings.ToLower(strings.TrimPrefix(ext, ".")) {
	case "png", "jpg", "jpeg", "jpe", "jfif", "gif", "bmp", "svg", "webp":
		return true
	}
	return false
}
