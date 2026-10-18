package docx

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
	"html"
	"io"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// xnode is a minimal generic XML element tree (namespace prefixes dropped).
type xnode struct {
	name  string
	attrs []xml.Attr
	kids  []*xnode
	text  string
}

func (n *xnode) attr(local string) string {
	if n == nil {
		return ""
	}
	for _, a := range n.attrs {
		if a.Name.Local == local {
			return a.Value
		}
	}
	return ""
}

func (n *xnode) child(name string) *xnode {
	if n == nil {
		return nil
	}
	for _, k := range n.kids {
		if k.name == name {
			return k
		}
	}
	return nil
}

// val returns the w:val attribute of child name, and whether the child exists.
func (n *xnode) val(name string) (string, bool) {
	c := n.child(name)
	if c == nil {
		return "", false
	}
	return c.attr("val"), true
}

// onOff reports whether a toggle property like <w:b/> is set and enabled.
func (n *xnode) onOff(name string) bool {
	v, ok := n.val(name)
	if !ok {
		return false
	}
	switch v {
	case "0", "false", "off", "none":
		return false
	}
	return true
}

func parseXML(data []byte) (*xnode, error) {
	d := xml.NewDecoder(bytes.NewReader(data))
	root := &xnode{}
	stack := []*xnode{root}
	for {
		tok, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		top := stack[len(stack)-1]
		switch t := tok.(type) {
		case xml.StartElement:
			n := &xnode{name: t.Name.Local, attrs: t.Attr}
			top.kids = append(top.kids, n)
			stack = append(stack, n)
		case xml.EndElement:
			if len(stack) > 1 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			top.text += string(t)
		}
	}
	return root, nil
}

type rel struct {
	typ, target string
	external    bool
}

type reader struct {
	files    map[string]*zip.File
	rels     map[string]rel
	headings map[string]int    // styleId -> heading level
	styleNum map[string][2]int // styleId -> numId, ilvl
	numFmt   map[int]map[int]string
}

func (r *reader) read(name string) ([]byte, error) {
	f, ok := r.files[name]
	if !ok {
		return nil, fmt.Errorf("missing part %s", name)
	}
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(io.LimitReader(rc, 512<<20))
}

func (r *reader) readXML(name string) *xnode {
	b, err := r.read(name)
	if err != nil {
		return nil
	}
	n, err := parseXML(b)
	if err != nil {
		return nil
	}
	return n
}

func walkAll(n *xnode, name string, fn func(*xnode)) {
	if n == nil {
		return
	}
	for _, k := range n.kids {
		if k.name == name {
			fn(k)
		} else {
			walkAll(k, name, fn)
		}
	}
}

var headingRe = regexp.MustCompile(`(?i)^heading\s*([1-6])$`)

// ToHTML converts a .docx file to semantic HTML. It returns the HTML body
// fragment and the document title (from core properties or the first heading).
func ToHTML(docx []byte) (string, string, error) {
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return "", "", errors.New("not a valid .docx file (zip): " + err.Error())
	}
	r := &reader{files: map[string]*zip.File{}, rels: map[string]rel{},
		headings: map[string]int{}, styleNum: map[string][2]int{}, numFmt: map[int]map[int]string{}}
	for _, f := range zr.File {
		r.files[strings.TrimPrefix(f.Name, "/")] = f
	}
	docPath := "word/document.xml"
	if root := r.readXML("_rels/.rels"); root != nil {
		walkAll(root, "Relationship", func(n *xnode) {
			if strings.HasSuffix(n.attr("Type"), "/officeDocument") {
				docPath = strings.TrimPrefix(n.attr("Target"), "/")
			}
		})
	}
	docData, err := r.read(docPath)
	if err != nil {
		return "", "", errors.New("not a Word document: " + err.Error())
	}
	doc, err := parseXML(docData)
	if err != nil {
		return "", "", errors.New("invalid document.xml: " + err.Error())
	}
	docDir := path.Dir(docPath)
	relsPath := path.Join(docDir, "_rels", path.Base(docPath)+".rels")
	if root := r.readXML(relsPath); root != nil {
		walkAll(root, "Relationship", func(n *xnode) {
			t := n.attr("Target")
			ext := n.attr("TargetMode") == "External"
			if !ext {
				if strings.HasPrefix(t, "/") {
					t = strings.TrimPrefix(t, "/")
				} else {
					t = path.Join(docDir, t)
				}
			}
			r.rels[n.attr("Id")] = rel{typ: n.attr("Type"), target: t, external: ext}
		})
	}
	stylesPath, numPath := path.Join(docDir, "styles.xml"), path.Join(docDir, "numbering.xml")
	for _, rl := range r.rels {
		if strings.HasSuffix(rl.typ, "/styles") {
			stylesPath = rl.target
		} else if strings.HasSuffix(rl.typ, "/numbering") {
			numPath = rl.target
		}
	}
	r.loadStyles(r.readXML(stylesPath))
	r.loadNumbering(r.readXML(numPath))

	title := ""
	if core := r.readXML("docProps/core.xml"); core != nil {
		walkAll(core, "title", func(n *xnode) {
			if title == "" {
				title = strings.TrimSpace(n.text)
			}
		})
	}

	var body *xnode
	walkAll(doc, "body", func(n *xnode) {
		if body == nil {
			body = n
		}
	})
	if body == nil {
		return "", title, nil
	}
	bc := &blockConv{r: r}
	bc.blocks(body.kids)
	bc.closeLists()
	if title == "" {
		title = bc.firstHeading
	}
	return bc.out.String(), title, nil
}

func (r *reader) loadStyles(root *xnode) {
	walkAll(root, "style", func(s *xnode) {
		if s.attr("type") != "paragraph" {
			return
		}
		id := s.attr("styleId")
		name, _ := s.val("name")
		lvl := 0
		if m := headingRe.FindStringSubmatch(name); m != nil {
			lvl, _ = strconv.Atoi(m[1])
		} else if m := headingRe.FindStringSubmatch(id); m != nil {
			lvl, _ = strconv.Atoi(m[1])
		} else if strings.EqualFold(name, "title") || strings.EqualFold(id, "title") {
			lvl = 1
		} else if ppr := s.child("pPr"); ppr != nil {
			if v, ok := ppr.val("outlineLvl"); ok {
				if n, err := strconv.Atoi(v); err == nil && n < 6 {
					lvl = n + 1
				}
			}
		}
		if lvl > 0 {
			r.headings[id] = lvl
		}
		if np := s.child("pPr").child("numPr"); np != nil {
			id2, _ := np.val("numId")
			il, _ := np.val("ilvl")
			n, _ := strconv.Atoi(id2)
			l, _ := strconv.Atoi(il)
			if n > 0 {
				r.styleNum[id] = [2]int{n, l}
			}
		}
	})
}

func (r *reader) loadNumbering(root *xnode) {
	if root == nil {
		return
	}
	abs := map[string]map[int]string{}
	walkAll(root, "abstractNum", func(a *xnode) {
		m := map[int]string{}
		for _, l := range a.kids {
			if l.name == "lvl" {
				il, _ := strconv.Atoi(l.attr("ilvl"))
				f, _ := l.val("numFmt")
				m[il] = f
			}
		}
		abs[a.attr("abstractNumId")] = m
	})
	walkAll(root, "num", func(n *xnode) {
		id, _ := strconv.Atoi(n.attr("numId"))
		a, _ := n.val("abstractNumId")
		r.numFmt[id] = abs[a]
	})
}

func (r *reader) isBullet(numID, ilvl int) bool {
	m, ok := r.numFmt[numID]
	if !ok {
		return true
	}
	f, ok := m[ilvl]
	if !ok {
		return true
	}
	return f == "bullet" || f == "none"
}

// blockConv converts a sequence of block-level elements (body or table cell).
type blockConv struct {
	r            *reader
	out          strings.Builder
	lists        []string // open list tags, one open <li> per level
	firstHeading string
}

func (b *blockConv) closeLists() {
	for len(b.lists) > 0 {
		b.popList()
	}
}

func (b *blockConv) popList() {
	t := b.lists[len(b.lists)-1]
	b.lists = b.lists[:len(b.lists)-1]
	b.out.WriteString("</li></" + t + ">")
}

func (b *blockConv) blocks(kids []*xnode) {
	for _, k := range kids {
		switch k.name {
		case "p":
			b.paragraph(k)
		case "tbl":
			b.closeLists()
			b.table(k)
		case "sectPr", "sdtPr", "sdtEndPr", "bookmarkStart", "bookmarkEnd", "del", "moveFrom":
		case "AlternateContent":
			if c := k.child("Choice"); c != nil {
				b.blocks(c.kids)
			} else if f := k.child("Fallback"); f != nil {
				b.blocks(f.kids)
			}
		default:
			b.blocks(k.kids)
		}
	}
}

func (b *blockConv) paragraph(p *xnode) {
	ppr := p.child("pPr")
	style, _ := ppr.val("pStyle")
	level := b.r.headings[style]
	numID, ilvl := 0, 0
	if sn, ok := b.r.styleNum[style]; ok {
		numID, ilvl = sn[0], sn[1]
	}
	if np := ppr.child("numPr"); np != nil {
		if v, ok := np.val("numId"); ok {
			numID, _ = strconv.Atoi(v)
		}
		if v, ok := np.val("ilvl"); ok {
			ilvl, _ = strconv.Atoi(v)
		}
	}
	if ilvl < 0 || ilvl > 8 {
		ilvl = 0
	}
	var css []string
	if jc, ok := ppr.val("jc"); ok {
		switch jc {
		case "center":
			css = append(css, "text-align:center")
		case "right", "end":
			css = append(css, "text-align:right")
		case "both", "distribute":
			css = append(css, "text-align:justify")
		}
	}
	if numID == 0 {
		if ind := ppr.child("ind"); ind != nil {
			l := ind.attr("left")
			if l == "" {
				l = ind.attr("start")
			}
			if n, err := strconv.Atoi(l); err == nil && n > 0 {
				css = append(css, fmt.Sprintf("margin-left:%gpt", float64(n)/20))
			}
		}
	}
	styleAttr := ""
	if len(css) > 0 {
		styleAttr = ` style="` + strings.Join(css, ";") + `"`
	}

	ic := &inlineConv{r: b.r}
	ic.walk(p.kids, runFmt{})
	content := ic.html()
	parts := strings.Split(content, pageBreakMark)

	if numID > 0 {
		tag := "ol"
		if b.r.isBullet(numID, ilvl) {
			tag = "ul"
		}
		for len(b.lists) > ilvl+1 {
			b.popList()
		}
		if len(b.lists) == ilvl+1 && b.lists[ilvl] != tag {
			b.popList()
		}
		if len(b.lists) == ilvl+1 {
			b.out.WriteString("</li>")
		} else {
			for len(b.lists) < ilvl+1 {
				b.out.WriteString("<" + tag + ">")
				b.lists = append(b.lists, tag)
				if len(b.lists) < ilvl+1 {
					b.out.WriteString("<li>")
				}
			}
		}
		b.out.WriteString("<li" + styleAttr + ">" + strings.Join(parts, ""))
		return
	}
	b.closeLists()
	tag := "p"
	if level > 0 {
		tag = "h" + strconv.Itoa(level)
		if b.firstHeading == "" {
			b.firstHeading = strings.TrimSpace(html.UnescapeString(stripTags(content)))
		}
	}
	for i, part := range parts {
		if i > 0 {
			b.out.WriteString(`<hr class="page-break">`)
		}
		if strings.TrimSpace(part) == "" {
			if len(parts) > 1 {
				continue
			}
			part = "<br>"
		}
		b.out.WriteString("<" + tag + styleAttr + ">" + part + "</" + tag + ">")
	}
}

var tagRe = regexp.MustCompile(`<[^>]*>`)

func stripTags(s string) string { return tagRe.ReplaceAllString(s, "") }

type cellInfo struct {
	col, span, vm int // vm: 0 none, 1 restart, 2 continue
	html          string
}

func (b *blockConv) table(t *xnode) {
	var rows [][]cellInfo
	var collectRows func(kids []*xnode)
	collectRows = func(kids []*xnode) {
		for _, tr := range kids {
			switch tr.name {
			case "tr":
				var cells []cellInfo
				col := 0
				var collectCells func(kids []*xnode)
				collectCells = func(kids []*xnode) {
					for _, tc := range kids {
						switch tc.name {
						case "tc":
							tcpr := tc.child("tcPr")
							span := 1
							if v, ok := tcpr.val("gridSpan"); ok {
								if n, err := strconv.Atoi(v); err == nil && n > 1 {
									span = n
								}
							}
							vm := 0
							if v, ok := tcpr.val("vMerge"); ok {
								if v == "restart" {
									vm = 1
								} else {
									vm = 2
								}
							}
							cb := &blockConv{r: b.r}
							cb.blocks(tc.kids)
							cb.closeLists()
							cells = append(cells, cellInfo{col: col, span: span, vm: vm, html: cb.out.String()})
							col += span
						case "trPr", "tblPrEx":
						default:
							collectCells(tc.kids)
						}
					}
				}
				collectCells(tr.kids)
				rows = append(rows, cells)
			case "tblPr", "tblGrid":
			default:
				collectRows(tr.kids)
			}
		}
	}
	collectRows(t.kids)
	b.out.WriteString(`<table border="1" style="border-collapse:collapse"><tbody>`)
	for ri, row := range rows {
		b.out.WriteString("<tr>")
		for _, c := range row {
			if c.vm == 2 {
				continue
			}
			attrs := ""
			if c.span > 1 {
				attrs += fmt.Sprintf(` colspan="%d"`, c.span)
			}
			if c.vm == 1 {
				rs := 1
				for _, next := range rows[ri+1:] {
					found := false
					for _, nc := range next {
						if nc.col == c.col && nc.vm == 2 {
							found = true
						}
					}
					if !found {
						break
					}
					rs++
				}
				if rs > 1 {
					attrs += fmt.Sprintf(` rowspan="%d"`, rs)
				}
			}
			h := c.html
			if h == "" {
				h = "<p><br></p>"
			}
			b.out.WriteString("<td" + attrs + ">" + h + "</td>")
		}
		b.out.WriteString("</tr>")
	}
	b.out.WriteString("</tbody></table>")
}

const pageBreakMark = "\x00PB\x00"

type runFmt struct {
	b, i, u, s, sup, sub bool
	color, bg, font      string
	sz                   int // half-points
}

func (f runFmt) css() string {
	var css []string
	if f.color != "" {
		css = append(css, "color:#"+f.color)
	}
	if f.bg != "" {
		css = append(css, "background-color:"+f.bg)
	}
	if f.sz > 0 {
		css = append(css, fmt.Sprintf("font-size:%gpt", float64(f.sz)/2))
	}
	if f.font != "" {
		css = append(css, "font-family:"+cssFontName(f.font))
	}
	return strings.Join(css, ";")
}

func cssFontName(f string) string {
	f = strings.Map(func(r rune) rune {
		if r == '"' || r == '\'' || r == ';' || r == '<' || r == '>' || r == '&' {
			return -1
		}
		return r
	}, f)
	if strings.ContainsRune(f, ' ') {
		return "'" + f + "'"
	}
	return f
}

type segment struct {
	f       runFmt
	content string
	raw     bool // content is final HTML (images, links, page-break)
}

type inlineConv struct {
	r    *reader
	segs []segment
}

func (ic *inlineConv) add(f runFmt, content string) {
	if n := len(ic.segs); n > 0 && !ic.segs[n-1].raw && ic.segs[n-1].f == f {
		ic.segs[n-1].content += content
		return
	}
	ic.segs = append(ic.segs, segment{f: f, content: content})
}

func (ic *inlineConv) addRaw(h string) {
	ic.segs = append(ic.segs, segment{content: h, raw: true})
}

func (ic *inlineConv) html() string {
	var sb strings.Builder
	for _, s := range ic.segs {
		if s.raw {
			sb.WriteString(s.content)
			continue
		}
		open, close := "", ""
		wrap := func(o, c string) { open += o; close = c + close }
		if css := s.f.css(); css != "" {
			wrap(`<span style="`+html.EscapeString(css)+`">`, "</span>")
		}
		if s.f.b {
			wrap("<strong>", "</strong>")
		}
		if s.f.i {
			wrap("<em>", "</em>")
		}
		if s.f.u {
			wrap("<u>", "</u>")
		}
		if s.f.s {
			wrap("<s>", "</s>")
		}
		if s.f.sup {
			wrap("<sup>", "</sup>")
		} else if s.f.sub {
			wrap("<sub>", "</sub>")
		}
		sb.WriteString(open + s.content + close)
	}
	return sb.String()
}

var highlightCSS = map[string]string{
	"yellow": "#ffff00", "green": "#00ff00", "cyan": "#00ffff", "magenta": "#ff00ff",
	"blue": "#0000ff", "red": "#ff0000", "darkBlue": "#000080", "darkCyan": "#008080",
	"darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
	"darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#c0c0c0", "black": "#000000",
	"white": "#ffffff",
}

var hexRe = regexp.MustCompile(`^[0-9A-Fa-f]{6}$`)

func readRunFmt(rpr *xnode, f runFmt) runFmt {
	if rpr == nil {
		return f
	}
	if _, ok := rpr.val("b"); ok {
		f.b = rpr.onOff("b")
	}
	if _, ok := rpr.val("i"); ok {
		f.i = rpr.onOff("i")
	}
	if v, ok := rpr.val("u"); ok {
		f.u = v != "none" && v != "0" && v != "false"
	}
	if rpr.onOff("strike") || rpr.onOff("dstrike") {
		f.s = true
	}
	if v, ok := rpr.val("vertAlign"); ok {
		f.sup, f.sub = v == "superscript", v == "subscript"
	}
	if v, ok := rpr.val("color"); ok && hexRe.MatchString(v) && v != "000000" {
		f.color = strings.ToLower(v)
	}
	if v, ok := rpr.val("highlight"); ok {
		if c, ok := highlightCSS[v]; ok {
			f.bg = c
		}
	}
	if shd := rpr.child("shd"); shd != nil {
		if fill := shd.attr("fill"); hexRe.MatchString(fill) {
			f.bg = "#" + strings.ToLower(fill)
		}
	}
	if v, ok := rpr.val("sz"); ok {
		if n, err := strconv.Atoi(v); err == nil && n > 0 {
			f.sz = n
		}
	}
	if rf := rpr.child("rFonts"); rf != nil {
		if a := rf.attr("ascii"); a != "" {
			f.font = a
		} else if a := rf.attr("hAnsi"); a != "" {
			f.font = a
		}
	}
	return f
}

func (ic *inlineConv) walk(kids []*xnode, base runFmt) {
	for _, k := range kids {
		switch k.name {
		case "r":
			ic.run(k, base)
		case "hyperlink":
			href := ""
			if id := k.attr("id"); id != "" {
				if rl, ok := ic.r.rels[id]; ok && rl.external {
					href = rl.target
				}
			}
			if a := k.attr("anchor"); a != "" && href == "" {
				href = "#" + a
			}
			sub := &inlineConv{r: ic.r}
			sub.walk(k.kids, base)
			inner := strings.ReplaceAll(sub.html(), pageBreakMark, "")
			if href == "" || !safeHref(href) {
				ic.addRaw(inner)
			} else {
				ic.addRaw(`<a href="` + html.EscapeString(href) + `">` + inner + `</a>`)
			}
		case "pPr", "rPr", "del", "moveFrom", "instrText", "delText", "fldChar", "bookmarkStart", "bookmarkEnd", "proofErr", "sdtPr", "sdtEndPr":
		case "AlternateContent":
			if c := k.child("Choice"); c != nil {
				ic.walk(c.kids, base)
			} else if f := k.child("Fallback"); f != nil {
				ic.walk(f.kids, base)
			}
		default:
			ic.walk(k.kids, base)
		}
	}
}

func safeHref(h string) bool {
	l := strings.ToLower(strings.TrimSpace(h))
	return !strings.HasPrefix(l, "javascript:") && !strings.HasPrefix(l, "vbscript:") && !strings.HasPrefix(l, "data:")
}

func (ic *inlineConv) run(r *xnode, base runFmt) {
	f := readRunFmt(r.child("rPr"), base)
	for _, k := range r.kids {
		switch k.name {
		case "t":
			if k.text != "" {
				ic.add(f, html.EscapeString(k.text))
			}
		case "tab", "ptab":
			ic.add(f, `<span style="white-space:pre">`+"\t</span>")
		case "br", "cr":
			if k.attr("type") == "page" {
				ic.addRaw(pageBreakMark)
			} else {
				ic.add(f, "<br>")
			}
		case "noBreakHyphen":
			ic.add(f, "‑")
		case "drawing", "pict", "object":
			ic.image(k)
		case "AlternateContent":
			if c := k.child("Choice"); c != nil {
				for _, cc := range c.kids {
					if cc.name == "drawing" {
						ic.image(cc)
					}
				}
			}
		}
	}
}

func (ic *inlineConv) image(d *xnode) {
	var embed string
	var cx, cy int64
	alt := ""
	var find func(n *xnode)
	find = func(n *xnode) {
		for _, k := range n.kids {
			switch k.name {
			case "blip":
				if embed == "" {
					embed = k.attr("embed")
				}
			case "imagedata":
				if embed == "" {
					embed = k.attr("id")
				}
			case "extent":
				if cx == 0 {
					cx, _ = strconv.ParseInt(k.attr("cx"), 10, 64)
					cy, _ = strconv.ParseInt(k.attr("cy"), 10, 64)
				}
			case "docPr":
				alt = k.attr("descr")
			case "txbxContent":
				continue
			}
			find(k)
		}
	}
	find(d)
	rl, ok := ic.r.rels[embed]
	if embed == "" || !ok || rl.external {
		return
	}
	data, err := ic.r.read(rl.target)
	if err != nil {
		return
	}
	mime := ""
	switch strings.ToLower(path.Ext(rl.target)) {
	case ".png":
		mime = "image/png"
	case ".jpg", ".jpeg":
		mime = "image/jpeg"
	case ".gif":
		mime = "image/gif"
	case ".bmp":
		mime = "image/bmp"
	case ".svg":
		mime = "image/svg+xml"
	case ".webp":
		mime = "image/webp"
	default:
		return
	}
	attrs := ""
	if cx > 0 && cy > 0 {
		attrs = fmt.Sprintf(` width="%d" height="%d"`, (cx+4762)/9525, (cy+4762)/9525)
	}
	ic.addRaw(`<img src="data:` + mime + `;base64,` + base64.StdEncoding.EncodeToString(data) + `"` + attrs + ` alt="` + html.EscapeString(alt) + `">`)
}
