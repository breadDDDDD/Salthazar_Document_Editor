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
	"math"
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

func parseXML(data []byte) (*xnode, error) {
	d := xml.NewDecoder(bytes.NewReader(data))
	d.Strict = false
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

type note struct {
	kind, id string
	num      int
}

type reader struct {
	files    map[string]*zip.File
	theme    themeInfo
	styles   map[string]*styleDef
	defR     rpr
	defP     ppr
	defPara  string
	defTable string
	nums     map[int]*numDef
	counters map[string]*numCounter
	numUsed  map[int]bool
	notes    []note
	noteBody map[string]map[string]*xnode // kind -> id -> footnote element
	noteRels map[string]map[string]rel
	tabStop  int // twips
	imgCache map[string]string
	fld      []int // complex field stack: 0 = instruction part, 1 = result part
}

func (r *reader) read(name string) ([]byte, error) {
	f, ok := r.files[name]
	if !ok {
		// part names are case-insensitive in OPC
		for k, v := range r.files {
			if strings.EqualFold(k, name) {
				f, ok = v, true
				break
			}
		}
		if !ok {
			return nil, fmt.Errorf("missing part %s", name)
		}
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

// readRels loads the relationships of a part, resolving internal targets to
// package paths.
func (r *reader) readRels(part string) map[string]rel {
	out := map[string]rel{}
	dir := path.Dir(part)
	root := r.readXML(path.Join(dir, "_rels", path.Base(part)+".rels"))
	walkAll(root, "Relationship", func(n *xnode) {
		t := n.attr("Target")
		ext := n.attr("TargetMode") == "External"
		if !ext {
			if strings.HasPrefix(t, "/") {
				t = strings.TrimPrefix(t, "/")
			} else {
				t = path.Join(dir, t)
			}
		}
		out[n.attr("Id")] = rel{typ: n.attr("Type"), target: t, external: ext}
	})
	return out
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

var (
	headingRe      = regexp.MustCompile(`(?i)^heading\s*([1-9])$`)
	numPlaceholder = regexp.MustCompile(`%[1-9]`)
)

// ToHTML converts a .docx file to HTML in the editor dialect (see board
// decision "DOCX HTML DIALECT v2"): a div.docx-doc wrapper carrying page
// setup, optional header/footer, and blocks with explicit inline formatting.
// It returns the HTML and the document title.
func ToHTML(docx []byte) (string, string, error) {
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return "", "", errors.New("not a valid .docx file (zip): " + err.Error())
	}
	r := &reader{files: map[string]*zip.File{}, counters: map[string]*numCounter{}, numUsed: map[int]bool{},
		noteBody: map[string]map[string]*xnode{}, noteRels: map[string]map[string]rel{}, imgCache: map[string]string{}, tabStop: 720}
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
	rels := r.readRels(docPath)
	parts := map[string]string{
		"styles": path.Join(docDir, "styles.xml"), "numbering": path.Join(docDir, "numbering.xml"),
		"theme": path.Join(docDir, "theme/theme1.xml"), "settings": path.Join(docDir, "settings.xml"),
		"footnotes": path.Join(docDir, "footnotes.xml"), "endnotes": path.Join(docDir, "endnotes.xml"),
	}
	for _, rl := range rels {
		if i := strings.LastIndexByte(rl.typ, '/'); i >= 0 && !rl.external {
			if _, ok := parts[rl.typ[i+1:]]; ok {
				parts[rl.typ[i+1:]] = rl.target
			}
		}
	}
	r.loadTheme(r.readXML(parts["theme"]))
	r.loadStyles(r.readXML(parts["styles"]))
	r.loadNumbering(r.readXML(parts["numbering"]))
	if s := r.readXML(parts["settings"]); s != nil {
		walkAll(s, "defaultTabStop", func(n *xnode) {
			if v, err := strconv.Atoi(n.attr("val")); err == nil && v > 0 {
				r.tabStop = v
			}
		})
	}
	for _, kind := range []string{"footnote", "endnote"} {
		p := parts[kind+"s"]
		m := map[string]*xnode{}
		walkAll(r.readXML(p), kind, func(n *xnode) { m[n.attr("id")] = n })
		r.noteBody[kind] = m
		r.noteRels[kind] = r.readRels(p)
	}

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

	// sections: page setup from the final section; header/footer references
	// are inherited from earlier sections when a section has none.
	var sects []*xnode
	walkAll(body, "sectPr", func(n *xnode) { sects = append(sects, n) })
	var hdrRef, ftrRef string
	for _, s := range sects {
		for _, k := range s.kids {
			if (k.name == "headerReference" || k.name == "footerReference") && (k.attr("type") == "default" || k.attr("type") == "") {
				if k.name == "headerReference" {
					hdrRef = k.attr("id")
				} else {
					ftrRef = k.attr("id")
				}
			}
		}
	}
	var sect *xnode
	if len(sects) > 0 {
		sect = sects[len(sects)-1]
	}

	bc := &blockConv{r: r, rels: rels}
	bc.blocks(body.kids)
	bc.closeLists()
	bodyHTML := bc.out.String()
	if title == "" {
		title = bc.firstHeading
	}

	part := func(id, tag, class string) string {
		rl, ok := rels[id]
		if id == "" || !ok || rl.external {
			return ""
		}
		root := r.readXML(rl.target)
		if root == nil || len(root.kids) == 0 {
			return ""
		}
		saved := r.fld
		r.fld = nil
		pc := &blockConv{r: r, rels: r.readRels(rl.target)}
		pc.blocks(root.kids[0].kids)
		pc.closeLists()
		r.fld = saved
		h := pc.out.String()
		if strings.TrimSpace(stripTags(h)) == "" && !strings.Contains(h, "<img") {
			return ""
		}
		return "<" + tag + ` class="` + class + `">` + h + "</" + tag + ">"
	}
	header := part(hdrRef, "header", "docx-header")
	footer := part(ftrRef, "footer", "docx-footer")
	notes := r.notesHTML()

	var sb strings.Builder
	sb.WriteString(r.wrapperOpen(sect))
	sb.WriteString(header)
	sb.WriteString(footer)
	sb.WriteString(bodyHTML)
	sb.WriteString(notes)
	sb.WriteString("</div>")
	return sb.String(), title, nil
}

func (r *reader) wrapperOpen(sect *xnode) string {
	pw, ph := 12240, 15840
	mt, mr, mb, ml, hd, fd := 1440, 1440, 1440, 1440, 720, 720
	orient := "portrait"
	if sz := sect.child("pgSz"); sz != nil {
		if v, err := strconv.Atoi(sz.attr("w")); err == nil && v > 0 {
			pw = v
		}
		if v, err := strconv.Atoi(sz.attr("h")); err == nil && v > 0 {
			ph = v
		}
		if sz.attr("orient") == "landscape" || pw > ph {
			orient = "landscape"
		}
	}
	if m := sect.child("pgMar"); m != nil {
		for _, p := range []struct {
			dst *int
			a   string
		}{{&mt, "top"}, {&mr, "right"}, {&mb, "bottom"}, {&ml, "left"}, {&hd, "header"}, {&fd, "footer"}} {
			if v, err := strconv.Atoi(m.attr(p.a)); err == nil {
				if v < 0 {
					v = -v
				}
				*p.dst = v
			}
		}
		if g, err := strconv.Atoi(m.attr("gutter")); err == nil && g > 0 {
			ml += g
		}
	}
	base := r.defR
	if st := r.style(r.defPara); st != nil {
		base = base.merge(st.effR)
	}
	font := base.font()
	if font == "" {
		font = "Times New Roman"
	}
	sz := 10.0
	if base.sz > 0 {
		sz = float64(base.sz) / 2
	}
	return fmt.Sprintf(`<div class="docx-doc" data-page-w="%d" data-page-h="%d" data-margin-top="%d" data-margin-right="%d" data-margin-bottom="%d" data-margin-left="%d" data-header-dist="%d" data-footer-dist="%d" data-orient="%s" data-font="%s" data-font-size="%g" data-tab="%g">`,
		pw, ph, mt, mr, mb, ml, hd, fd, orient, html.EscapeString(font), sz, float64(r.tabStop)/20)
}

func (r *reader) notesHTML() string {
	if len(r.notes) == 0 {
		return ""
	}
	var sb strings.Builder
	sb.WriteString(`<div class="docx-footnotes" style="border-top:0.75pt solid #000000;margin-top:12pt;padding-top:4pt">`)
	for i := 0; i < len(r.notes); i++ { // notes may reference further notes
		n := r.notes[i]
		body := r.noteBody[n.kind][n.id]
		if body == nil {
			continue
		}
		nc := &blockConv{r: r, rels: r.noteRels[n.kind], noteNum: n.num}
		saved := r.fld
		r.fld = nil
		nc.blocks(body.kids)
		nc.closeLists()
		r.fld = saved
		h := nc.out.String()
		if !nc.noteMarked {
			h = fmt.Sprintf("<p><sup>%d</sup></p>", n.num) + h
		}
		sb.WriteString(h)
	}
	sb.WriteString("</div>")
	return sb.String()
}

// ---- block conversion ----

type tblCtx struct {
	p ppr
	r rpr
}

type listFrame struct {
	tag, key string
	left     int // twips, text indent of this level
	last     int // counter value of the last item
}

// blockConv converts a sequence of block-level elements (body, cell, note...).
type blockConv struct {
	r            *reader
	rels         map[string]rel
	out          strings.Builder
	lists        []listFrame
	firstHeading string
	tbl          *tblCtx
	noteNum      int
	noteMarked   bool
	prevStyle    string
	prevCtx      bool
}

func (b *blockConv) closeLists() {
	for len(b.lists) > 0 {
		b.popList()
	}
}

func (b *blockConv) popList() {
	t := b.lists[len(b.lists)-1]
	b.lists = b.lists[:len(b.lists)-1]
	b.out.WriteString("</li></" + t.tag + ">")
}

func (b *blockConv) blocks(kids []*xnode) {
	for i, k := range kids {
		switch k.name {
		case "p":
			var next *xnode
			if i+1 < len(kids) && kids[i+1].name == "p" {
				next = kids[i+1]
			}
			b.paragraph(k, next)
		case "tbl":
			b.closeLists()
			b.table(k)
			b.prevStyle = ""
		case "sectPr", "sdtPr", "sdtEndPr", "bookmarkStart", "bookmarkEnd", "del", "moveFrom", "tblPr", "tblGrid",
			"proofErr", "permStart", "permEnd", "commentRangeStart", "commentRangeEnd":
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

func (b *blockConv) paraStyleID(p *xnode) string {
	if s, ok := p.child("pPr").val("pStyle"); ok && s != "" {
		return s
	}
	return b.r.defPara
}

func fmtNum(v float64) string {
	return strconv.FormatFloat(math.Round(v*100)/100, 'f', -1, 64)
}

func fmtPt(prop string, v float64) string { return prop + ":" + fmtNum(v) + "pt" }

var serifRe = regexp.MustCompile(`(?i)times|cambria|georgia|garamond|book|roman|palatino|serif|minion|baskerville|bodoni|century|constantia`)
var monoRe = regexp.MustCompile(`(?i)courier|consolas|mono|lucida console`)

func fontStack(f rpr) string {
	names := []string{}
	add := func(n string) {
		n = strings.TrimSpace(n)
		if n == "" {
			return
		}
		for _, x := range names {
			if strings.EqualFold(x, n) {
				return
			}
		}
		names = append(names, n)
	}
	add(f.font())
	if len(names) == 0 {
		add("Times New Roman")
	}
	add(f.eastAsia)
	parts := make([]string, 0, len(names)+1)
	for _, n := range names {
		parts = append(parts, cssFontName(n))
	}
	switch {
	case monoRe.MatchString(names[0]):
		parts = append(parts, "monospace")
	case serifRe.MatchString(names[0]):
		parts = append(parts, "serif")
	default:
		parts = append(parts, "sans-serif")
	}
	return strings.Join(parts, ", ")
}

func cssFontName(f string) string {
	f = strings.Map(func(r rune) rune {
		if r == '"' || r == '\'' || r == ';' || r == '<' || r == '>' || r == '&' || r == ',' {
			return -1
		}
		return r
	}, f)
	if strings.ContainsAny(f, " 0123456789") {
		return "'" + f + "'"
	}
	return f
}

func isDark(c string) bool { return c == "000000" }

// runBaseCSS returns the CSS for the paragraph-level run formatting.
func runBaseCSS(base rpr, heading bool) []string {
	css := []string{"font-family:" + fontStack(base)}
	sz := 10.0
	if base.sz > 0 {
		sz = float64(base.sz) / 2
	}
	css = append(css, fmtPt("font-size", sz))
	if hexRe.MatchString(base.color) && !isDark(base.color) {
		css = append(css, "color:#"+base.color)
	} else if heading {
		css = append(css, "color:inherit")
	}
	if base.b.is() {
		css = append(css, "font-weight:bold")
	} else if heading {
		css = append(css, "font-weight:normal")
	}
	if base.i.is() {
		css = append(css, "font-style:italic")
	} else if heading {
		css = append(css, "font-style:normal")
	}
	if base.caps.is() {
		css = append(css, "text-transform:uppercase")
	}
	if base.smallCaps.is() {
		css = append(css, "font-variant:small-caps")
	}
	return css
}

func lineHeightCSS(p ppr, base rpr) string {
	if !p.line.ok || p.line.v <= 0 {
		return "line-height:1.15"
	}
	switch p.lineRule {
	case "exact":
		return fmtPt("line-height", float64(p.line.v)/20)
	case "atLeast":
		sz := 10.0
		if base.sz > 0 {
			sz = float64(base.sz) / 2
		}
		return fmtPt("line-height", math.Max(float64(p.line.v)/20, sz*1.15))
	}
	return "line-height:" + strconv.FormatFloat(math.Round(float64(p.line.v)/240*1.15*10000)/10000, 'f', -1, 64)
}

func spacingPt(v optInt, auto tri) float64 {
	if auto.is() {
		return 14
	}
	if v.ok {
		return float64(v.v) / 20
	}
	return 0
}

func (b *blockConv) paragraph(p, next *xnode) {
	r := b.r
	ppNode := p.child("pPr")
	styleID := b.paraStyleID(p)
	st := r.style(styleID)
	effP, baseR := r.defP, r.defR
	if b.tbl != nil {
		effP, baseR = effP.merge(b.tbl.p), baseR.merge(b.tbl.r)
	}
	if st != nil {
		effP, baseR = effP.merge(st.effP), baseR.merge(st.effR)
	}
	direct := r.parsePPr(ppNode)
	numP := effP.merge(direct)
	numID, ilvl := 0, 0
	if numP.numID.ok {
		numID = numP.numID.v
	}
	if numP.ilvl.ok {
		ilvl = numP.ilvl.v
	}
	if ilvl < 0 || ilvl > 8 {
		ilvl = 0
	}
	var lvl *numLevel
	if numID > 0 {
		if lvl = r.numLevel(numID, ilvl); lvl != nil {
			effP = effP.merge(lvl.p)
		}
	}
	effP = effP.merge(direct)
	level, isTitle := r.headingLevel(styleID)
	if level > 6 {
		level = 0
	}

	ic := &inlineConv{b: b, r: r, rels: b.rels, floatTab: effP.rightTab.is()}
	ic.walk(p.kids, baseR)
	content := ic.html(baseR, level > 0)
	empty := strings.TrimSpace(strings.ReplaceAll(content, pageBreakMark, "")) == ""
	if empty {
		// an empty paragraph's height comes from its paragraph mark
		baseR = baseR.merge(r.parseRPr(ppNode.child("rPr")))
	}

	label := ""
	if lvl != nil {
		label, lvl, _ = r.nextLabelFull(numID, ilvl)
	}
	listItem := lvl != nil && level == 0 && !(lvl.fmt == "none" && strings.TrimSpace(label) == "")

	// spacing (with contextual spacing between paragraphs of the same style)
	before, after := spacingPt(effP.before, effP.beforeAuto), spacingPt(effP.after, effP.afterAuto)
	if effP.contextual.is() {
		if b.prevStyle == styleID && b.prevCtx {
			before = 0
		}
		if next != nil && b.paraStyleID(next) == styleID {
			after = 0
		}
	}
	b.prevStyle, b.prevCtx = styleID, effP.contextual.is()

	css := []string{fmtPt("margin-top", before), fmtPt("margin-bottom", after)}
	if !listItem {
		if effP.left.ok && effP.left.v != 0 {
			css = append(css, fmtPt("margin-left", float64(effP.left.v)/20))
		}
		if effP.first.ok && effP.first.v != 0 {
			css = append(css, fmtPt("text-indent", float64(effP.first.v)/20))
		} else if effP.hanging.ok && effP.hanging.v != 0 {
			css = append(css, fmtPt("text-indent", -float64(effP.hanging.v)/20))
		}
	}
	if effP.right.ok && effP.right.v > 0 {
		css = append(css, fmtPt("margin-right", float64(effP.right.v)/20))
	}
	switch effP.jc {
	case "center":
		css = append(css, "text-align:center")
	case "right", "end":
		css = append(css, "text-align:right")
	case "both", "distribute", "lowKashida", "mediumKashida", "highKashida", "thaiDistribute":
		css = append(css, "text-align:justify")
	default:
		if level > 0 {
			css = append(css, "text-align:left")
		}
	}
	css = append(css, lineHeightCSS(effP, baseR))
	if hexRe.MatchString(effP.shd) {
		css = append(css, "background-color:#"+effP.shd)
	}
	if effP.borderTop != "" && effP.borderTop != "none" {
		css = append(css, "border-top:"+effP.borderTop, "padding-top:1pt")
	}
	if effP.borderBottom != "" && effP.borderBottom != "none" {
		css = append(css, "border-bottom:"+effP.borderBottom, "padding-bottom:1pt")
	}
	css = append(css, runBaseCSS(baseR, level > 0)...)
	styleAttr := ` style="` + html.EscapeString(strings.Join(css, ";")) + `"`

	parts := strings.Split(content, pageBreakMark)

	if effP.pageBreakBefore.is() && !listItem {
		b.closeLists()
		b.out.WriteString(`<hr class="page-break">`)
	}

	if listItem {
		b.listItem(lvl, numID, ilvl, label, effP, styleAttr, strings.Join(parts, ""))
	} else {
		b.closeLists()
		tag := "p"
		cls := ""
		if level > 0 {
			tag = "h" + strconv.Itoa(level)
			if isTitle {
				cls = ` class="title"`
			}
			if b.firstHeading == "" {
				b.firstHeading = strings.TrimSpace(html.UnescapeString(stripTags(content)))
			}
			if label != "" && lvl != nil && len(parts) > 0 {
				parts[0] = html.EscapeString(label) + "  " + parts[0]
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
			b.out.WriteString("<" + tag + cls + styleAttr + ">" + part + "</" + tag + ">")
		}
	}
	for _, a := range ic.after {
		b.closeLists()
		b.out.WriteString(a)
	}
	if sp := ppNode.child("sectPr"); sp != nil {
		if t, _ := sp.val("type"); t != "continuous" && next != nil {
			b.closeLists()
			b.out.WriteString(`<hr class="page-break">`)
		}
	}
}

var simpleOl = map[string]string{"decimal": "decimal", "lowerLetter": "lower-alpha", "upperLetter": "upper-alpha",
	"lowerRoman": "lower-roman", "upperRoman": "upper-roman", "decimalZero": "decimal-leading-zero"}

func (b *blockConv) listItem(lvl *numLevel, numID, ilvl int, label string, effP ppr, styleAttr, content string) {
	tag := "ol"
	listStyle, custom := "", false
	if lvl.fmt == "bullet" {
		tag = "ul"
		switch label {
		case "•", "●":
			listStyle = "disc"
		case "o", "◦", "○":
			listStyle = "circle"
		case "▪", "■", "§":
			listStyle = "square"
		default:
			custom = true
		}
	} else if ks, ok := simpleOl[lvl.fmt]; ok && lvl.text == "%"+strconv.Itoa(ilvl+1)+"." {
		listStyle = ks
	} else {
		custom = true
	}
	key := fmt.Sprintf("%s|%s|%s|%d|%v", tag, lvl.fmt, lvl.text, numID, custom)
	val := 0
	if c := b.r.counters[b.r.nums[numID].abs]; c != nil {
		val = c.vals[ilvl]
	}
	left := 0
	if effP.left.ok {
		left = effP.left.v
	}

	for len(b.lists) > ilvl+1 {
		b.popList()
	}
	if len(b.lists) == ilvl+1 && b.lists[ilvl].key != key {
		b.popList()
	}
	liAttr := ""
	if len(b.lists) == ilvl+1 {
		b.out.WriteString("</li>")
		if tag == "ol" && !custom && val != b.lists[ilvl].last+1 {
			liAttr = fmt.Sprintf(` value="%d"`, val)
		}
		b.lists[ilvl].last = val
	} else {
		for len(b.lists) < ilvl+1 {
			parentLeft := 0
			if n := len(b.lists); n > 0 {
				parentLeft = b.lists[n-1].left
			}
			if len(b.lists) < ilvl {
				// missing intermediate level: an unmarked wrapper item
				b.out.WriteString(`<ul style="list-style-type:none;margin-top:0;margin-bottom:0;padding-left:0"><li style="list-style-type:none">`)
				b.lists = append(b.lists, listFrame{tag: "ul", key: "wrapper", left: parentLeft})
				continue
			}
			pad := float64(left-parentLeft) / 20
			if pad < 0 {
				pad = 0
			}
			ls := listStyle
			if custom {
				ls = "none"
			}
			lt := lvl.text
			if lvl.fmt == "bullet" {
				lt = label
			}
			attrs := fmt.Sprintf(` data-numfmt="%s" data-lvltext="%s"`, html.EscapeString(lvl.fmt), html.EscapeString(lt))
			if tag == "ol" && val != 1 {
				attrs = fmt.Sprintf(` start="%d"`, val) + attrs
			}
			b.out.WriteString("<" + tag + attrs + ` style="list-style-type:` + ls + ";margin-top:0;margin-bottom:0;" + fmtPt("padding-left", pad) + `">`)
			b.lists = append(b.lists, listFrame{tag: tag, key: key, left: left, last: val})
		}
	}
	if custom {
		lbl := strings.NewReplacer(`\`, `\\`, `'`, `\'`).Replace(label) + " "
		styleAttr = strings.Replace(styleAttr, ` style="`, ` style="`+html.EscapeString("list-style-type:'"+lbl+"';"), 1)
	}
	if strings.TrimSpace(content) == "" {
		content = "<br>"
	}
	b.out.WriteString("<li" + liAttr + styleAttr + ">" + content)
}

// nextLabelFull wraps nextLabel and keeps the returned level non-nil when possible.
func (r *reader) nextLabelFull(numID, ilvl int) (string, *numLevel, int) {
	l, lv, v := r.nextLabel(numID, ilvl)
	if lv == nil {
		lv = r.numLevel(numID, ilvl)
	}
	return l, lv, v
}

var tagRe = regexp.MustCompile(`<[^>]*>`)

func stripTags(s string) string { return tagRe.ReplaceAllString(s, "") }

// ---- tables ----

type cellInfo struct {
	col, span, vm int // vm: 0 none, 1 restart, 2 continue
	tc            *xnode
}

type tblStyleInfo struct {
	borders map[string]*xnode // top,left,bottom,right,insideH,insideV
	mar     map[string]int    // top,left,bottom,right (twips)
	p       ppr
	r       rpr
	cond    map[string]*xnode // tblStylePr type -> node
	jc      string
	ind     optInt
}

func (r *reader) collectTblPr(info *tblStyleInfo, tp *xnode) {
	if tp == nil {
		return
	}
	if bd := tp.child("tblBorders"); bd != nil {
		for _, k := range bd.kids {
			name := map[string]string{"start": "left", "end": "right"}[k.name]
			if name == "" {
				name = k.name
			}
			info.borders[name] = k
		}
	}
	if m := tp.child("tblCellMar"); m != nil {
		for _, k := range m.kids {
			name := map[string]string{"start": "left", "end": "right"}[k.name]
			if name == "" {
				name = k.name
			}
			if v, err := strconv.Atoi(k.attr("w")); err == nil {
				info.mar[name] = v
			}
		}
	}
	if v, ok := tp.val("jc"); ok {
		info.jc = v
	}
	if ind := tp.child("tblInd"); ind != nil {
		if x, ok := atoiOpt(ind.attr("w")); ok && ind.attr("type") != "pct" {
			info.ind = x
		}
	}
}

func (b *blockConv) table(t *xnode) {
	r := b.r
	tblPr := t.child("tblPr")
	info := &tblStyleInfo{borders: map[string]*xnode{}, mar: map[string]int{"left": 108, "right": 108}, cond: map[string]*xnode{}}
	sid, _ := tblPr.val("tblStyle")
	if sid == "" {
		sid = r.defTable
	}
	if st := r.style(sid); st != nil {
		for i := len(st.chain) - 1; i >= 0; i-- {
			s := st.chain[i]
			r.collectTblPr(info, s.tblPr)
			for _, k := range s.node.kids {
				if k.name == "tblStylePr" {
					info.cond[k.attr("type")] = k
				}
			}
		}
		info.p, info.r = st.effP, st.effR
	}
	r.collectTblPr(info, tblPr)
	look := tblPr.child("tblLook")
	lookFlag := func(attr string, bit int64) bool {
		if look == nil {
			return attr == "firstRow" || attr == "firstColumn"
		}
		if v := look.attr(attr); v != "" {
			return v == "1" || v == "true"
		}
		x, _ := strconv.ParseInt(look.attr("val"), 16, 64)
		return x&bit != 0
	}
	firstRow, lastRow, firstCol := lookFlag("firstRow", 0x20), lookFlag("lastRow", 0x40), lookFlag("firstColumn", 0x80)
	noHBand := lookFlag("noHBand", 0x200)

	var grid []int
	for _, g := range kidsOf(t.child("tblGrid")) {
		if g.name == "gridCol" {
			v, _ := strconv.Atoi(g.attr("w"))
			grid = append(grid, v)
		}
	}

	var rows [][]cellInfo
	var rowNodes []*xnode
	var collectRows func(kids []*xnode)
	collectRows = func(kids []*xnode) {
		for _, tr := range kids {
			switch tr.name {
			case "tr":
				var cells []cellInfo
				col := 0
				if v, ok := tr.child("trPr").val("gridBefore"); ok {
					col, _ = strconv.Atoi(v)
				}
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
							cells = append(cells, cellInfo{col: col, span: span, vm: vm, tc: tc})
							col += span
						case "trPr", "tblPrEx":
						default:
							collectCells(tc.kids)
						}
					}
				}
				collectCells(tr.kids)
				rows = append(rows, cells)
				rowNodes = append(rowNodes, tr)
			case "tblPr", "tblGrid":
			default:
				collectRows(tr.kids)
			}
		}
	}
	collectRows(t.kids)
	ncols := len(grid)
	for _, row := range rows {
		if n := len(row); n > 0 && row[n-1].col+row[n-1].span > ncols {
			ncols = row[n-1].col + row[n-1].span
		}
	}

	tcss := []string{"border-collapse:collapse", "margin-top:0", "margin-bottom:0"}
	total := 0
	for _, g := range grid {
		total += g
	}
	if tw := tblPr.child("tblW"); tw != nil && tw.attr("type") == "pct" {
		v, _ := strconv.ParseFloat(strings.TrimSuffix(tw.attr("w"), "%"), 64)
		if !strings.HasSuffix(tw.attr("w"), "%") {
			v /= 50
		}
		if v > 0 {
			tcss = append(tcss, "width:"+fmtNum(v)+"%")
		}
	} else if total > 0 {
		tcss = append(tcss, fmtPt("width", float64(total)/20))
	}
	switch info.jc {
	case "center":
		tcss = append(tcss, "margin-left:auto", "margin-right:auto")
	case "right", "end":
		tcss = append(tcss, "margin-left:auto", "margin-right:0")
	default:
		if info.ind.ok && info.ind.v != 0 {
			tcss = append(tcss, fmtPt("margin-left", float64(info.ind.v)/20))
		}
	}
	b.out.WriteString(`<table style="` + html.EscapeString(strings.Join(tcss, ";")) + `"><tbody>`)

	condFor := func(ri, ci int) []*xnode {
		var out []*xnode
		add := func(k string) {
			if c := info.cond[k]; c != nil {
				out = append(out, c)
			}
		}
		dataRow := ri
		if firstRow {
			dataRow--
		}
		if !noHBand && dataRow >= 0 && !(lastRow && ri == len(rows)-1) {
			if dataRow%2 == 0 {
				add("band1Horz")
			} else {
				add("band2Horz")
			}
		}
		if firstCol && ci == 0 {
			add("firstCol")
		}
		if firstRow && ri == 0 {
			add("firstRow")
		}
		if lastRow && ri == len(rows)-1 && ri > 0 {
			add("lastRow")
		}
		return out
	}

	for ri, row := range rows {
		trcss := ""
		if h := rowNodes[ri].child("trPr").child("trHeight"); h != nil {
			if v, err := strconv.Atoi(h.attr("val")); err == nil && v > 0 {
				trcss = ` style="` + fmtPt("height", float64(v)/20) + `"`
			}
		}
		b.out.WriteString("<tr" + trcss + ">")
		for _, c := range row {
			if c.vm == 2 {
				continue
			}
			attrs := ""
			if c.span > 1 {
				attrs += fmt.Sprintf(` colspan="%d"`, c.span)
			}
			rs := 1
			if c.vm == 1 {
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
			tcpr := c.tc.child("tcPr")
			ctx := &tblCtx{p: info.p, r: info.r}
			shd := ""
			conds := condFor(ri, c.col)
			for _, cn := range conds {
				ctx.p = ctx.p.merge(r.parsePPr(cn.child("pPr")))
				ctx.r = ctx.r.merge(r.parseRPr(cn.child("rPr")))
				if s := cn.child("tcPr").child("shd"); s != nil {
					if v := r.colorOf(s, "fill"); hexRe.MatchString(v) {
						shd = v
					}
				}
			}
			if s := tcpr.child("shd"); s != nil {
				if v := r.colorOf(s, "fill"); hexRe.MatchString(v) {
					shd = v
				} else if v == "auto" {
					shd = ""
				}
			}
			var css []string
			w := 0
			for g := c.col; g < c.col+c.span && g < len(grid); g++ {
				w += grid[g]
			}
			if w == 0 {
				if tw := tcpr.child("tcW"); tw != nil && (tw.attr("type") == "dxa" || tw.attr("type") == "") {
					w, _ = strconv.Atoi(tw.attr("w"))
				}
			}
			if w > 0 {
				css = append(css, fmtPt("width", float64(w)/20))
			}
			// borders: cell overrides > conditional > table (outer edges vs inside)
			lastRowIdx := ri + rs - 1
			sides := map[string]string{
				"top":    pick(ri == 0, "top", "insideH"),
				"bottom": pick(lastRowIdx == len(rows)-1, "bottom", "insideH"),
				"left":   pick(c.col == 0, "left", "insideV"),
				"right":  pick(c.col+c.span >= ncols, "right", "insideV"),
			}
			cellB := map[string]*xnode{}
			for _, cn := range conds {
				if bd := cn.child("tcPr").child("tcBorders"); bd != nil {
					for _, k := range bd.kids {
						cellB[k.name] = k
					}
				}
			}
			if bd := tcpr.child("tcBorders"); bd != nil {
				for _, k := range bd.kids {
					cellB[k.name] = k
				}
			}
			for _, side := range []string{"top", "right", "bottom", "left"} {
				bn := cellB[side]
				if bn == nil {
					bn = cellB[map[string]string{"left": "start", "right": "end"}[side]]
				}
				if bn == nil {
					bn = info.borders[sides[side]]
				}
				v := "none"
				if bn != nil {
					v = r.borderCSS(bn)
				}
				css = append(css, "border-"+side+":"+v)
			}
			mar := map[string]int{}
			for k, v := range info.mar {
				mar[k] = v
			}
			if m := tcpr.child("tcMar"); m != nil {
				for _, k := range m.kids {
					name := map[string]string{"start": "left", "end": "right"}[k.name]
					if name == "" {
						name = k.name
					}
					if v, err := strconv.Atoi(k.attr("w")); err == nil {
						mar[name] = v
					}
				}
			}
			css = append(css, "padding:"+fmtNum(float64(mar["top"])/20)+"pt "+fmtNum(float64(mar["right"])/20)+"pt "+
				fmtNum(float64(mar["bottom"])/20)+"pt "+fmtNum(float64(mar["left"])/20)+"pt")
			if shd != "" {
				css = append(css, "background-color:#"+shd)
			}
			switch va, _ := tcpr.val("vAlign"); va {
			case "center":
				css = append(css, "vertical-align:middle")
			case "bottom":
				css = append(css, "vertical-align:bottom")
			default:
				css = append(css, "vertical-align:top")
			}
			cb := &blockConv{r: r, rels: b.rels, tbl: ctx, noteNum: b.noteNum}
			cb.blocks(c.tc.kids)
			cb.closeLists()
			h := cb.out.String()
			if h == "" {
				h = "<p><br></p>"
			}
			b.out.WriteString("<td" + attrs + ` style="` + html.EscapeString(strings.Join(css, ";")) + `">` + h + "</td>")
		}
		b.out.WriteString("</tr>")
	}
	b.out.WriteString("</tbody></table>")
}

func pick(c bool, a, b string) string {
	if c {
		return a
	}
	return b
}

// ---- inline conversion ----

const pageBreakMark = "\x00PB\x00"
const tabHTML = `<span style="white-space:pre">` + "\t</span>"

type segment struct {
	f       rpr
	content string
	raw     bool // content is final HTML (images, links, page-break)
}

type inlineConv struct {
	b        *blockConv
	r        *reader
	rels     map[string]rel
	segs     []segment
	after    []string // block HTML to emit after the paragraph (text boxes)
	floatTab bool     // right-align the text after the last tab (right tab stop)
}

func (ic *inlineConv) add(f rpr, content string) {
	if n := len(ic.segs); n > 0 && !ic.segs[n-1].raw && ic.segs[n-1].f == f {
		ic.segs[n-1].content += content
		return
	}
	ic.segs = append(ic.segs, segment{f: f, content: content})
}

func (ic *inlineConv) addRaw(h string) {
	ic.segs = append(ic.segs, segment{content: h, raw: true})
}

// html renders the segments; base is the paragraph-level formatting already
// expressed by the block's CSS, so only differences are emitted.
func (ic *inlineConv) html(base rpr, heading bool) string {
	segs := ic.segs
	floatAt := -1
	if ic.floatTab {
		for i := len(segs) - 1; i >= 0; i-- {
			if strings.Contains(segs[i].content, tabHTML) {
				if !segs[i].raw {
					floatAt = i
				}
				break
			}
		}
		if floatAt >= 0 {
			// split the segment so the float starts exactly at the last tab
			s := segs[floatAt]
			j := strings.LastIndex(s.content, tabHTML)
			if strings.Contains(s.content[j:], pageBreakMark) {
				floatAt = -1
			} else {
				segs = append(append(append([]segment{}, segs[:floatAt]...),
					segment{f: s.f, content: s.content[:j]}, segment{f: s.f, content: s.content[j:]}), segs[floatAt+1:]...)
				floatAt++
			}
		}
	}
	var sb strings.Builder
	for si, s := range segs {
		if si == floatAt {
			sb.WriteString(`<span style="float:right">`)
		}
		if s.raw {
			sb.WriteString(s.content)
			continue
		}
		if s.content == "" {
			continue
		}
		f := s.f
		open, close := "", ""
		wrap := func(o, c string) { open += o; close = c + close }
		var css []string
		if f.font() != "" && !strings.EqualFold(f.font(), base.font()) {
			css = append(css, "font-family:"+fontStack(f))
		}
		if f.sz > 0 && f.sz != base.sz {
			css = append(css, fmtPt("font-size", float64(f.sz)/2))
		}
		if f.color != base.color {
			if f.color == "auto" && hexRe.MatchString(base.color) && !isDark(base.color) {
				css = append(css, "color:#000000")
			} else if hexRe.MatchString(f.color) {
				if !isDark(f.color) || hexRe.MatchString(base.color) && !isDark(base.color) {
					css = append(css, "color:#"+f.color)
				}
			}
		}
		if hexRe.MatchString(f.bg) {
			css = append(css, "background-color:#"+f.bg)
		}
		if f.b.is() != base.b.is() && !f.b.is() {
			css = append(css, "font-weight:normal")
		}
		if f.i.is() != base.i.is() && !f.i.is() {
			css = append(css, "font-style:normal")
		}
		if f.caps.is() && !base.caps.is() {
			css = append(css, "text-transform:uppercase")
		}
		if f.smallCaps.is() && !base.smallCaps.is() {
			css = append(css, "font-variant:small-caps")
		}
		if f.spacing.ok && f.spacing.v != 0 && f.spacing != base.spacing {
			css = append(css, fmtPt("letter-spacing", float64(f.spacing.v)/20))
		}
		if len(css) > 0 {
			wrap(`<span style="`+html.EscapeString(strings.Join(css, ";"))+`">`, "</span>")
		}
		if f.b.is() && !base.b.is() {
			wrap("<strong>", "</strong>")
		}
		if f.i.is() && !base.i.is() {
			wrap("<em>", "</em>")
		}
		if f.u != "" && f.u != "none" {
			wrap("<u>", "</u>")
		}
		if f.strike.is() {
			wrap("<s>", "</s>")
		}
		switch f.vert {
		case "superscript":
			wrap("<sup>", "</sup>")
		case "subscript":
			wrap("<sub>", "</sub>")
		}
		sb.WriteString(open + s.content + close)
	}
	if floatAt >= 0 {
		sb.WriteString("</span>")
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

func (ic *inlineConv) inField() bool {
	n := len(ic.r.fld)
	return n > 0 && ic.r.fld[n-1] == 0
}

func (ic *inlineConv) walk(kids []*xnode, base rpr) {
	for _, k := range kids {
		switch k.name {
		case "r":
			ic.run(k, base)
		case "hyperlink":
			href := ""
			if id := k.attr("id"); id != "" {
				if rl, ok := ic.rels[id]; ok && rl.external {
					href = rl.target
				}
			}
			if a := k.attr("anchor"); a != "" && href == "" {
				href = "#" + a
			}
			sub := &inlineConv{b: ic.b, r: ic.r, rels: ic.rels, floatTab: ic.floatTab}
			sub.walk(k.kids, base)
			ic.after = append(ic.after, sub.after...)
			inner := strings.ReplaceAll(sub.html(base, false), pageBreakMark, "")
			if href == "" || !safeHref(href) || ic.inField() {
				ic.addRaw(inner)
			} else {
				ic.addRaw(`<a href="` + html.EscapeString(href) + `">` + inner + `</a>`)
			}
		case "oMath", "oMathPara":
			if !ic.inField() {
				ic.math(k, base)
			}
		case "pPr", "rPr", "del", "moveFrom", "instrText", "delText", "bookmarkStart", "bookmarkEnd", "proofErr",
			"sdtPr", "sdtEndPr", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd":
		case "fldSimple":
			ic.walk(k.kids, base)
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

// runProps returns the effective formatting of a run: paragraph base, then
// the character style chain, then direct formatting.
func (r *reader) runProps(rn *xnode, base rpr) rpr {
	var rp *xnode
	for _, k := range rn.kids {
		if k.name == "rPr" && (rp == nil || k.child("rFonts") != nil || k.child("b") != nil || k.child("sz") != nil || k.child("rStyle") != nil) {
			rp = k // prefer w:rPr over m:rPr inside math runs
		}
	}
	f := base
	if id, ok := rp.val("rStyle"); ok {
		if st := r.style(id); st != nil {
			f = f.merge(st.effR)
		}
	}
	return f.merge(r.parseRPr(rp))
}

var symbolGreek = map[rune]rune{
	'A': 'Α', 'B': 'Β', 'C': 'Χ', 'D': 'Δ', 'E': 'Ε', 'F': 'Φ', 'G': 'Γ', 'H': 'Η', 'I': 'Ι', 'J': 'ϑ', 'K': 'Κ',
	'L': 'Λ', 'M': 'Μ', 'N': 'Ν', 'O': 'Ο', 'P': 'Π', 'Q': 'Θ', 'R': 'Ρ', 'S': 'Σ', 'T': 'Τ', 'U': 'Υ', 'V': 'ς',
	'W': 'Ω', 'X': 'Ξ', 'Y': 'Ψ', 'Z': 'Ζ', 'a': 'α', 'b': 'β', 'c': 'χ', 'd': 'δ', 'e': 'ε', 'f': 'φ', 'g': 'γ',
	'h': 'η', 'i': 'ι', 'j': 'ϕ', 'k': 'κ', 'l': 'λ', 'm': 'μ', 'n': 'ν', 'o': 'ο', 'p': 'π', 'q': 'θ', 'r': 'ρ',
	's': 'σ', 't': 'τ', 'u': 'υ', 'v': 'ϖ', 'w': 'ω', 'x': 'ξ', 'y': 'ψ', 'z': 'ζ',
	0xB1: '±', 0xB4: '×', 0xB8: '÷', 0xA3: '≤', 0xB3: '≥', 0xB9: '≠', 0xBB: '≈', 0xA5: '∞', 0xB6: '∂', 0xD6: '√',
	0xE5: '∑', 0xF2: '∫', 0xAE: '→', 0xAC: '←', 0xAD: '↑', 0xAF: '↓', 0xB0: '°', 0xB7: '•', 0xD7: '⋅', 0x22: '∀',
	0x24: '∃', 0x27: '∋', 0xCE: '∈', 0xCF: '∉', 0xC7: '∩', 0xC8: '∪', 0xDE: '⇒', 0xDB: '⇔', 0xA2: '′', 0xB2: '″',
	0xBC: '…', 0xC6: '∅', 0xD1: '∇', 0xD5: '∏', 0xD8: '¬', 0xD9: '∧', 0xDA: '∨', 0xBA: '≡', 0x40: '≅', 0x5E: '⊥',
	0xA4: '⁄', 0xA6: 'ƒ', 0xC5: '⊕', 0xC4: '⊗', 0xCC: '⊂', 0xC9: '⊃', 0xCD: '⊆', 0xCA: '⊇', 0x60: '‾', 0x7E: '∼',
}

// symbolText maps text set in the legacy Symbol font to Unicode.
func symbolText(s, font string) string {
	if !strings.EqualFold(font, "Symbol") {
		return s
	}
	return strings.Map(func(c rune) rune {
		if c >= 0xF000 && c <= 0xF0FF {
			c -= 0xF000
		}
		if m, ok := symbolGreek[c]; ok {
			return m
		}
		return c
	}, s)
}

func (ic *inlineConv) run(rn *xnode, base rpr) {
	f := ic.r.runProps(rn, base)
	hidden := f.vanish.is()
	for _, k := range rn.kids {
		if k.name == "fldChar" {
			switch k.attr("fldCharType") {
			case "begin":
				ic.r.fld = append(ic.r.fld, 0)
			case "separate":
				if n := len(ic.r.fld); n > 0 {
					ic.r.fld[n-1] = 1
				}
			case "end":
				if n := len(ic.r.fld); n > 0 {
					ic.r.fld = ic.r.fld[:n-1]
				}
			}
			continue
		}
		if hidden || ic.inField() {
			continue
		}
		switch k.name {
		case "t":
			if k.text != "" {
				ic.add(f, html.EscapeString(symbolText(k.text, f.font())))
			}
		case "tab", "ptab":
			ic.add(f, tabHTML)
		case "br", "cr":
			if k.attr("type") == "page" {
				ic.addRaw(pageBreakMark)
			} else {
				ic.add(f, "<br>")
			}
		case "noBreakHyphen":
			ic.add(f, "‑")
		case "sym":
			code, err := strconv.ParseUint(k.attr("char"), 16, 32)
			if err == nil {
				font := k.attr("font")
				c := rune(code)
				var s string
				if strings.Contains(strings.ToLower(font), "wingdings") {
					s = bulletChar(string(c), font)
				} else {
					s = symbolText(string(c), font)
				}
				ic.add(f, html.EscapeString(s))
			}
		case "footnoteReference", "endnoteReference":
			kind := strings.TrimSuffix(k.name, "Reference")
			n := len(ic.r.notes) + 1
			ic.r.notes = append(ic.r.notes, note{kind: kind, id: k.attr("id"), num: n})
			ic.addRaw(fmt.Sprintf("<sup>%d</sup>", n))
		case "footnoteRef", "endnoteRef":
			if ic.b != nil && ic.b.noteNum > 0 {
				ic.b.noteMarked = true
				ic.addRaw(fmt.Sprintf("<sup>%d</sup>", ic.b.noteNum))
			}
		case "drawing", "pict", "object":
			ic.image(k)
		case "AlternateContent":
			if c := k.child("Choice"); c != nil {
				for _, cc := range c.kids {
					if cc.name == "drawing" || cc.name == "pict" {
						ic.image(cc)
					}
				}
			} else if fb := k.child("Fallback"); fb != nil {
				for _, cc := range fb.kids {
					if cc.name == "drawing" || cc.name == "pict" {
						ic.image(cc)
					}
				}
			}
		case "oMath", "oMathPara":
			ic.math(k, f)
		}
	}
}

// imageURL returns a data: URL for an embedded image part ("" if unsupported).
func (ic *inlineConv) imageURL(rl rel) string {
	if u, ok := ic.r.imgCache[rl.target]; ok {
		return u
	}
	u := ""
	if data, err := ic.r.read(rl.target); err == nil {
		if mime, out, ok := WebImage(data, path.Ext(rl.target)); ok {
			u = "data:" + mime + ";base64," + base64.StdEncoding.EncodeToString(out)
		}
	}
	ic.r.imgCache[rl.target] = u
	return u
}

var vmlLenRe = regexp.MustCompile(`(?i)(width|height)\s*:\s*([\d.]+)\s*(pt|px|in|cm|mm)?`)

func (ic *inlineConv) image(d *xnode) {
	var embed string
	var cx, cy int64
	var vw, vh float64 // VML size in px
	alt := ""
	var anchor *xnode
	var boxes []*xnode
	var find func(n *xnode)
	find = func(n *xnode) {
		for _, k := range n.kids {
			switch k.name {
			case "anchor":
				if anchor == nil {
					anchor = k
				}
			case "blip":
				if embed == "" {
					embed = k.attr("embed")
					if embed == "" {
						embed = k.attr("link")
					}
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
				if alt == "" {
					alt = k.attr("descr")
				}
			case "shape", "rect", "roundrect", "oval":
				if vw == 0 {
					for _, m := range vmlLenRe.FindAllStringSubmatch(k.attr("style"), -1) {
						v, _ := strconv.ParseFloat(m[2], 64)
						switch strings.ToLower(m[3]) {
						case "", "px":
						case "pt":
							v *= 4.0 / 3
						case "in":
							v *= 96
						case "cm":
							v *= 96 / 2.54
						case "mm":
							v *= 96 / 25.4
						}
						if strings.EqualFold(m[1], "width") {
							vw = v
						} else {
							vh = v
						}
					}
				}
			case "txbxContent":
				boxes = append(boxes, k)
				continue
			case "AlternateContent":
				if c := k.child("Choice"); c != nil {
					find(c)
				}
				continue
			}
			find(k)
		}
	}
	find(d)
	if anchor == nil && d.name == "drawing" {
		anchor = d.child("anchor")
	}
	for _, tb := range boxes {
		bc := &blockConv{r: ic.r, rels: ic.rels, noteNum: ic.b.noteNum}
		bc.blocks(tb.kids)
		bc.closeLists()
		if h := bc.out.String(); strings.TrimSpace(stripTags(h)) != "" || strings.Contains(h, "<img") {
			ic.after = append(ic.after, `<div class="docx-textbox" style="border:0.75pt solid #000000;padding:4pt;margin:4pt 0">`+h+`</div>`)
		}
	}
	if embed == "" {
		return
	}
	wpx, hpx := float64(cx)/9525, float64(cy)/9525
	if wpx <= 0 || hpx <= 0 {
		wpx, hpx = vw, vh
	}
	attrs := ""
	if wpx > 0 && hpx > 0 {
		attrs = fmt.Sprintf(` width="%d" height="%d"`, int(math.Round(wpx)), int(math.Round(hpx)))
	}
	style := ""
	if anchor != nil {
		align := ""
		if ph := anchor.child("positionH"); ph != nil {
			if a := ph.child("align"); a != nil {
				align = strings.TrimSpace(a.text)
			}
		}
		wrapSq := anchor.child("wrapSquare") != nil || anchor.child("wrapTight") != nil || anchor.child("wrapThrough") != nil
		switch {
		case wrapSq && (align == "right" || align == "outside"):
			style = "float:right;margin:0 0 4pt 8pt"
		case wrapSq && (align == "left" || align == "inside"):
			style = "float:left;margin:0 8pt 4pt 0"
		case anchor.child("wrapTopAndBottom") != nil || align == "center":
			style = "display:block;margin:4pt auto"
		}
	}
	rl, ok := ic.rels[embed]
	src := ""
	if ok && !rl.external {
		src = ic.imageURL(rl)
	}
	if src == "" {
		// keep a visible, correctly sized placeholder for pictures we cannot show
		label := alt
		if label == "" {
			label = "Picture"
		}
		ps := "display:inline-block;border:1px dashed #808080;color:#606060;font-size:9pt;text-align:center;vertical-align:middle"
		if wpx > 0 && hpx > 0 {
			ps += fmt.Sprintf(";width:%dpx;height:%dpx;line-height:%dpx", int(wpx), int(hpx), int(hpx))
		}
		ic.addRaw(`<span style="` + ps + `">[` + html.EscapeString(label) + `]</span>`)
		return
	}
	if style != "" {
		attrs += ` style="` + style + `"`
	}
	ic.addRaw(`<img src="` + src + `"` + attrs + ` alt="` + html.EscapeString(alt) + `">`)
}

func kidsOf(n *xnode) []*xnode {
	if n == nil {
		return nil
	}
	return n.kids
}
