package docx

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"fmt"
	"image"
	_ "image/gif"  // register decoder for DecodeConfig
	_ "image/jpeg" // register decoder for DecodeConfig
	_ "image/png"  // register decoder for DecodeConfig
	"math"
	"net/url"
	"regexp"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

const (
	nsW   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	nsR   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	relHL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	relIM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	relHD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	relFT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

	nsDecl = ` xmlns:w="` + nsW + `" xmlns:r="` + nsR + `"` +
		` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`
)

type pProps struct {
	style         string
	jc            string
	numID, ilvl   int
	listLeft      int // twips: text start of the enclosing list level
	indent        int // left indent, twips
	right         int // right indent, twips
	first         int // first-line indent (>0) or hanging (<0), twips
	firstSet      bool
	before, after optInt // spacing, twips
	line          int    // spacing line value
	lineRule      string // auto | exact | atLeast
	shd           string // RRGGBB
	pre           bool
	border        bool   // bottom border (horizontal rule)
	bdrTop        string // border XML attributes
	bdrBottom     string
	markRPr       string // paragraph mark run properties (empty paragraphs)
}

type rProps struct {
	b, i, u, s, sup, sub, link, pre bool
	noB, noI, caps, smallCaps       bool
	color, bg, font                 string // colors as RRGGBB ("auto" allowed for color)
	sz                              int    // half-points
	spacing                         int    // character spacing, twips
}

type para struct {
	pp        pProps
	runs      []string
	lastSpace bool
	brRun     []bool
}

type outRel struct {
	id, typ, target string
	external        bool
}

// lvlDef is one level of a generated abstract numbering definition.
type lvlDef struct {
	fmt, text, font string
	left, hanging   int
	start           int
}

type numInst struct {
	id     int
	levels [9]lvlDef
}

type pageSetup struct {
	w, h, top, right, bottom, left, header, footer int
	landscape                                      bool
	font                                           string
	sz                                             int  // half-points
	tab                                            int  // twips
	v2                                             bool // came from a docx-doc wrapper
}

type hfPart struct {
	name, kind, xml string
	rels            []outRel
}

type writer struct {
	out       *strings.Builder
	cur       *para
	blocks    int
	lastTable bool
	rels      []outRel
	media     [][2]string // name, data
	nums      []*numInst
	bulletNum int
	drawID    int
	firstHead string
	page      pageSetup
	header    *hnode
	footer    *hnode
	parts     []hfPart
}

func defaultPage() pageSetup {
	return pageSetup{w: 12240, h: 15840, top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720,
		font: "Calibri", sz: 22, tab: 720}
}

// FromHTML converts an HTML fragment or document to a .docx file.
func FromHTML(src string) ([]byte, error) {
	root, title := parseHTML(src)
	w := &writer{out: &strings.Builder{}, page: defaultPage()}
	w.container(root.kids, pProps{}, rProps{})
	w.flush()
	if w.blocks == 0 || w.lastTable {
		w.out.WriteString("<w:p/>")
	}
	if title == "" {
		title = w.firstHead
	}
	if w.header != nil {
		w.renderPart("header", w.header)
	}
	if w.footer != nil {
		w.renderPart("footer", w.footer)
	}
	return w.pack(title)
}

// renderPart converts a header/footer element into its own part.
func (w *writer) renderPart(kind string, n *hnode) {
	saveOut, saveRels, saveBlocks, saveLast, saveCur := w.out, w.rels, w.blocks, w.lastTable, w.cur
	w.out, w.rels, w.blocks, w.lastTable, w.cur = &strings.Builder{}, nil, 0, false, nil
	w.container(n.kids, pProps{}, rProps{})
	w.flush()
	if w.blocks == 0 || w.lastTable {
		w.out.WriteString("<w:p/>")
	}
	tag := "w:hdr"
	if kind == "footer" {
		tag = "w:ftr"
	}
	name := fmt.Sprintf("%s%d.xml", kind, len(w.parts)+1)
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n<" + tag + nsDecl + ">" + w.out.String() + "</" + tag + ">"
	w.parts = append(w.parts, hfPart{name: name, kind: kind, xml: xml, rels: w.rels})
	w.out, w.rels, w.blocks, w.lastTable, w.cur = saveOut, saveRels, saveBlocks, saveLast, saveCur
}

func (w *writer) addRel(typ, target string, external bool) string {
	id := fmt.Sprintf("rId%d", 10+len(w.rels))
	w.rels = append(w.rels, outRel{id, typ, target, external})
	return id
}

func (w *writer) ensure(pp pProps) *para {
	if w.cur == nil {
		w.cur = &para{pp: pp, lastSpace: true}
	}
	return w.cur
}

func (w *writer) addRun(xml string, isBr bool) {
	p := w.cur
	p.runs = append(p.runs, xml)
	p.brRun = append(p.brRun, isBr)
}

// flush writes the current paragraph if it has any runs.
func (w *writer) flush() {
	p := w.cur
	w.cur = nil
	if p == nil || len(p.runs) == 0 {
		return
	}
	if n := len(p.runs); p.brRun[n-1] && n > 1 {
		p.runs = p.runs[:n-1]
	}
	w.writePara(p.pp, strings.Join(p.runs, ""))
}

func (w *writer) writePara(pp pProps, runs string) {
	w.out.WriteString("<w:p>")
	var ppr strings.Builder
	if pp.style != "" {
		ppr.WriteString(`<w:pStyle w:val="` + pp.style + `"/>`)
	}
	if pp.numID > 0 {
		fmt.Fprintf(&ppr, `<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>`, pp.ilvl, pp.numID)
	}
	if pp.border || pp.bdrTop != "" || pp.bdrBottom != "" {
		ppr.WriteString("<w:pBdr>")
		if pp.bdrTop != "" {
			ppr.WriteString(`<w:top ` + pp.bdrTop + `/>`)
		}
		if pp.bdrBottom != "" {
			ppr.WriteString(`<w:bottom ` + pp.bdrBottom + `/>`)
		} else if pp.border {
			ppr.WriteString(`<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>`)
		}
		ppr.WriteString("</w:pBdr>")
	}
	if pp.shd != "" {
		ppr.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="` + pp.shd + `"/>`)
	}
	var sp []string
	if pp.before.ok {
		sp = append(sp, fmt.Sprintf(`w:before="%d"`, pp.before.v))
	}
	if pp.after.ok {
		sp = append(sp, fmt.Sprintf(`w:after="%d"`, pp.after.v))
	} else if pp.pre {
		sp = append(sp, `w:after="0"`)
	}
	if pp.line > 0 {
		sp = append(sp, fmt.Sprintf(`w:line="%d" w:lineRule="%s"`, pp.line, pp.lineRule))
	}
	if len(sp) > 0 {
		ppr.WriteString("<w:spacing " + strings.Join(sp, " ") + "/>")
	}
	if pp.numID == 0 && (pp.indent != 0 || pp.right > 0 || pp.firstSet) {
		var ind []string
		if pp.indent != 0 {
			ind = append(ind, fmt.Sprintf(`w:left="%d"`, pp.indent))
		}
		if pp.right > 0 {
			ind = append(ind, fmt.Sprintf(`w:right="%d"`, pp.right))
		}
		if pp.firstSet {
			if pp.first >= 0 {
				ind = append(ind, fmt.Sprintf(`w:firstLine="%d"`, pp.first))
			} else {
				ind = append(ind, fmt.Sprintf(`w:hanging="%d"`, -pp.first))
			}
		}
		ppr.WriteString("<w:ind " + strings.Join(ind, " ") + "/>")
	}
	if pp.jc != "" {
		ppr.WriteString(`<w:jc w:val="` + pp.jc + `"/>`)
	}
	ppr.WriteString(pp.markRPr)
	if ppr.Len() > 0 {
		w.out.WriteString("<w:pPr>" + ppr.String() + "</w:pPr>")
	}
	w.out.WriteString(runs)
	w.out.WriteString("</w:p>")
	w.blocks++
	w.lastTable = false
}

var blockTags = map[string]bool{
	"p": true, "div": true, "h1": true, "h2": true, "h3": true, "h4": true, "h5": true, "h6": true,
	"ul": true, "ol": true, "li": true, "table": true, "blockquote": true, "pre": true, "hr": true,
	"section": true, "article": true, "header": true, "footer": true, "nav": true, "aside": true,
	"main": true, "figure": true, "figcaption": true, "address": true, "dl": true, "dt": true,
	"dd": true, "body": true, "html": true, "center": true, "caption": true, "thead": true,
	"tbody": true, "tfoot": true, "tr": true, "td": true, "th": true, "form": true, "fieldset": true,
	"legend": true, "details": true, "summary": true,
}

var skipTags = map[string]bool{
	"head": true, "script": true, "style": true, "meta": true, "link": true, "noscript": true,
	"template": true, "iframe": true, "object": true, "svg": true, "math": true, "canvas": true,
	"input": true, "button": true, "select": true, "textarea": true, "video": true, "audio": true,
	"colgroup": true, "col": true,
}

func (w *writer) container(kids []*hnode, pp pProps, rp rProps) {
	for _, k := range kids {
		if k.tag != "" && blockTags[k.tag] {
			w.flush()
			w.block(k, pp, rp)
		} else {
			w.inline(k, pp, rp)
		}
	}
}

func alignOf(n *hnode, css map[string]string) string {
	a := css["text-align"]
	if a == "" {
		a = n.attr("align")
	}
	switch strings.ToLower(a) {
	case "center":
		return "center"
	case "right", "end":
		return "right"
	case "justify":
		return "both"
	case "left", "start":
		return "left"
	}
	return ""
}

// blockLike handles an element that maps to one or more paragraphs; if it
// produces nothing, an empty paragraph is written so blank lines survive.
func (w *writer) blockLike(n *hnode, pp pProps, rp rProps) {
	w.flush()
	before := w.blocks
	w.container(n.kids, pp, rp)
	w.flush()
	if w.blocks == before {
		// keep the paragraph mark formatting so blank lines keep their height
		pp.markRPr = runPr(rp)
		w.writePara(pp, "")
	}
}

func hasClass(n *hnode, c string) bool {
	return strings.Contains(" "+n.attr("class")+" ", " "+c+" ")
}

func attrInt(n *hnode, a string, def int) int {
	if v, err := strconv.Atoi(strings.TrimSpace(n.attr(a))); err == nil && v >= 0 {
		return v
	}
	return def
}

// wrapper reads the page setup carried by a div.docx-doc wrapper.
func (w *writer) wrapper(n *hnode) {
	p := &w.page
	// a wrapper that names the document font came from a converted .docx whose
	// paragraphs carry explicit formatting
	p.v2 = strings.TrimSpace(n.attr("data-font")) != ""
	p.w, p.h = attrInt(n, "data-page-w", p.w), attrInt(n, "data-page-h", p.h)
	p.top, p.right = attrInt(n, "data-margin-top", p.top), attrInt(n, "data-margin-right", p.right)
	p.bottom, p.left = attrInt(n, "data-margin-bottom", p.bottom), attrInt(n, "data-margin-left", p.left)
	p.header, p.footer = attrInt(n, "data-header-dist", p.header), attrInt(n, "data-footer-dist", p.footer)
	p.landscape = n.attr("data-orient") == "landscape"
	if f := strings.TrimSpace(n.attr("data-font")); f != "" {
		p.font = f
	}
	if f, err := strconv.ParseFloat(n.attr("data-font-size"), 64); err == nil && f > 0 {
		p.sz = clampSz(f * 2)
	}
	if f, err := strconv.ParseFloat(n.attr("data-tab"), 64); err == nil && f > 0 {
		p.tab = int(math.Round(f * 20))
	}
}

// blockCSS applies paragraph-level CSS to pp.
func blockCSS(css map[string]string, pp pProps) pProps {
	if v, ok := css["margin-top"]; ok {
		if t, ok := cssTwips(v); ok {
			pp.before = optInt{t, true}
		}
	}
	if v, ok := css["margin-bottom"]; ok {
		if t, ok := cssTwips(v); ok {
			pp.after = optInt{t, true}
		}
	}
	if v, ok := css["margin-right"]; ok {
		if t, ok := cssTwips(v); ok && t > 0 {
			pp.right = t
		}
	}
	if v, ok := css["text-indent"]; ok {
		if t, ok := cssTwips(v); ok {
			pp.first, pp.firstSet = t, t != 0
		}
	}
	if v := strings.ToLower(strings.TrimSpace(css["line-height"])); v != "" && v != "normal" {
		if f, err := strconv.ParseFloat(v, 64); err == nil && f > 0 {
			pp.line, pp.lineRule = int(math.Round(f/1.15*240)), "auto"
		} else if strings.HasSuffix(v, "%") {
			if f, err := strconv.ParseFloat(strings.TrimSuffix(v, "%"), 64); err == nil && f > 0 {
				pp.line, pp.lineRule = int(math.Round(f/100/1.15*240)), "auto"
			}
		} else if t, ok := cssTwips(v); ok && t > 0 {
			pp.line, pp.lineRule = t, "exact"
		}
	}
	if c := parseColor(css["background-color"]); c != "" {
		pp.shd = c
	}
	if b := borderXML(css["border-top"]); b != "" && b != "none" {
		pp.bdrTop = b
	}
	if b := borderXML(css["border-bottom"]); b != "" && b != "none" {
		pp.bdrBottom = b
	}
	return pp
}

func (w *writer) block(n *hnode, pp pProps, rp rProps) {
	if skipTags[n.tag] {
		return
	}
	if n.tag == "div" && hasClass(n, "docx-doc") {
		w.wrapper(n)
		w.container(n.kids, pp, rp)
		w.flush()
		return
	}
	if n.tag == "header" && hasClass(n, "docx-header") {
		w.header = n
		return
	}
	if n.tag == "footer" && hasClass(n, "docx-footer") {
		w.footer = n
		return
	}
	css := parseCSS(n.attr("style"))
	if a := alignOf(n, css); a != "" {
		if a == "left" {
			pp.jc = ""
		} else {
			pp.jc = a
		}
	}
	if n.tag != "li" && n.tag != "ul" && n.tag != "ol" && n.tag != "table" {
		if ml, ok := cssTwips(css["margin-left"]); ok && ml != 0 && pp.numID == 0 {
			pp.indent += ml
		} else if pl := cssLength(css["padding-left"]); pl > 0 && pp.numID == 0 {
			pp.indent += pl
		}
	}
	isPara := false
	switch n.tag {
	case "h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "dt", "dd", "figcaption", "address", "center", "caption", "summary", "legend", "pre", "div":
		isPara = true
	}
	if isPara {
		// paragraph-level properties apply to this element only, not inherited by nested blocks
		pp.before, pp.after, pp.line, pp.lineRule, pp.shd, pp.bdrTop, pp.bdrBottom = optInt{}, optInt{}, 0, "", "", "", ""
		pp.right, pp.first, pp.firstSet = 0, 0, false
		pp = blockCSS(css, pp)
		// block background is paragraph shading, not run highlight
		delete(css, "background-color")
		delete(css, "background")
	}
	heading := len(n.tag) == 2 && n.tag[0] == 'h' && n.tag[1] >= '1' && n.tag[1] <= '6'
	rp = applyInlineStyle(n, css, rp)
	if heading && strings.TrimSpace(strings.ToLower(css["color"])) == "inherit" {
		rp.color = "auto"
	}
	switch n.tag {
	case "h1", "h2", "h3", "h4", "h5", "h6":
		pp.style = "Heading" + n.tag[1:]
		if n.tag == "h1" && hasClass(n, "title") {
			pp.style = "Title"
		}
		pp.numID, pp.ilvl = 0, 0
		if w.firstHead == "" {
			w.firstHead = strings.TrimSpace(textOf(n))
		}
		w.blockLike(n, pp, rp)
	case "p", "div", "dt", "dd", "figcaption", "address", "center", "caption", "summary", "legend":
		if n.tag == "center" {
			pp.jc = "center"
		}
		if n.tag == "dd" {
			pp.indent += 720
		}
		if n.tag == "div" && hasClass(n, "docx-textbox") {
			pp.bdrTop, pp.bdrBottom = "", ""
		}
		if n.tag == "p" || hasDirectContent(n) {
			w.blockLike(n, pp, rp)
		} else {
			w.container(n.kids, pp, rp)
		}
	case "blockquote":
		pp.indent += 720
		w.container(n.kids, pp, rp)
		w.flush()
	case "pre":
		pp.pre = true
		rp.pre = true
		rp.font = "Courier New"
		w.blockLike(n, pp, rp)
	case "hr":
		if hasClass(n, "page-break") || strings.Contains(strings.ToLower(n.attr("style")), "page-break") {
			w.writePara(pProps{}, `<w:r><w:br w:type="page"/></w:r>`)
		} else {
			w.writePara(pProps{border: true}, "")
		}
	case "ul", "ol":
		w.list(n, pp, rp)
	case "li":
		pp.style = "ListParagraph"
		if pp.numID == 0 {
			pp.numID, pp.ilvl = w.bullet(), 0
		}
		w.blockLike(n, pp, rp)
	case "table":
		w.table(n, pp, rp)
	case "tr", "td", "th", "thead", "tbody", "tfoot":
		// stray table parts outside a table: treat as containers
		w.container(n.kids, pp, rp)
		w.flush()
	default:
		w.container(n.kids, pp, rp)
		w.flush()
	}
}

func hasDirectContent(n *hnode) bool {
	for _, k := range n.kids {
		if k.tag == "" {
			if strings.TrimSpace(k.text) != "" {
				return true
			}
		} else if !blockTags[k.tag] {
			return true
		}
	}
	for _, k := range n.kids {
		if k.tag != "" && blockTags[k.tag] {
			return false
		}
	}
	return true // empty div = blank line
}

func textOf(n *hnode) string {
	if n.tag == "" {
		return n.text
	}
	var sb strings.Builder
	for _, k := range n.kids {
		sb.WriteString(textOf(k))
	}
	return sb.String()
}

// ---- lists ----

func defaultLevels(bullet bool) [9]lvlDef {
	var lv [9]lvlDef
	bullets := []string{"•", "o", "▪"}
	formats := []string{"decimal", "lowerLetter", "lowerRoman"}
	for l := 0; l < 9; l++ {
		if bullet {
			lv[l] = lvlDef{fmt: "bullet", text: bullets[l%3], left: 720 * (l + 1), hanging: 360, start: 1}
			if bullets[l%3] == "o" {
				lv[l].font = "Courier New"
			}
		} else {
			lv[l] = lvlDef{fmt: formats[l%3], text: fmt.Sprintf("%%%d.", l+1), left: 720 * (l + 1), hanging: 360, start: 1}
		}
	}
	return lv
}

func (w *writer) newNum(levels [9]lvlDef) int {
	id := len(w.nums) + 1
	w.nums = append(w.nums, &numInst{id: id, levels: levels})
	return id
}

func (w *writer) bullet() int {
	if w.bulletNum == 0 {
		w.bulletNum = w.newNum(defaultLevels(true))
	}
	return w.bulletNum
}

var cssListFmt = map[string]string{
	"decimal": "decimal", "lower-alpha": "lowerLetter", "lower-latin": "lowerLetter", "upper-alpha": "upperLetter",
	"upper-latin": "upperLetter", "lower-roman": "lowerRoman", "upper-roman": "upperRoman",
	"decimal-leading-zero": "decimalZero", "disc": "bullet", "circle": "bullet", "square": "bullet",
}

var knownNumFmt = map[string]bool{"decimal": true, "lowerLetter": true, "upperLetter": true, "lowerRoman": true,
	"upperRoman": true, "decimalZero": true, "bullet": true, "none": true, "ordinal": true}

func (w *writer) list(n *hnode, pp pProps, rp rProps) {
	level := 0
	parentLeft := pp.indent
	if pp.numID > 0 {
		level = pp.ilvl + 1
		parentLeft = pp.listLeft
	}
	if level > 8 {
		level = 8
	}
	css := parseCSS(n.attr("style"))
	ls := strings.ToLower(strings.TrimSpace(css["list-style-type"]))
	numFmt := n.attr("data-numfmt")
	if !knownNumFmt[numFmt] {
		numFmt = cssListFmt[ls]
	}
	bullet := n.tag == "ul"
	if numFmt == "" {
		if bullet {
			numFmt = "bullet"
		} else {
			numFmt = "decimal"
		}
	}
	bullet = numFmt == "bullet"
	text := n.attr("data-lvltext")
	if text == "" || (bullet && utf8.RuneCountInString(text) > 2) || (!bullet && !strings.Contains(text, "%")) {
		switch {
		case ls == "circle":
			text = "o"
		case ls == "square":
			text = "▪"
		case bullet:
			text = defaultLevels(true)[level].text
		default:
			text = fmt.Sprintf("%%%d.", level+1)
		}
	}
	pad := 720
	if v, ok := cssTwips(css["padding-left"]); ok {
		pad = v
	}
	ml, _ := cssTwips(css["margin-left"])
	left := parentLeft + ml + pad
	start := 1
	if v, err := strconv.Atoi(strings.TrimSpace(n.attr("start"))); err == nil {
		start = v
	}
	levels := defaultLevels(bullet)
	hang := 360
	if left < hang {
		hang = left
	}
	lv := lvlDef{fmt: numFmt, text: text, left: left, hanging: hang, start: start}
	if text == "o" {
		lv.font = "Courier New"
	}
	levels[level] = lv
	numID := w.newNum(levels)
	lp := pProps{style: "ListParagraph", numID: numID, ilvl: level, jc: pp.jc, listLeft: left}
	for _, k := range n.kids {
		switch {
		case k.tag == "li":
			w.flush()
			w.block(k, lp, rp)
		case k.tag == "ul" || k.tag == "ol":
			w.flush()
			w.list(k, lp, rp)
		case k.tag == "" && strings.TrimSpace(k.text) == "":
		case k.tag != "" && blockTags[k.tag]:
			w.flush()
			w.block(k, lp, rp)
		default:
			w.inline(k, lp, rp)
		}
	}
	w.flush()
}

// ---- tables ----

var borderRe = regexp.MustCompile(`(?i)^\s*([\d.]+)(pt|px)?\s+(solid|double|dashed|dotted|inset|outset|groove|ridge)\s*(.*)$`)

// borderXML converts a CSS border shorthand into w: border attributes ("none" for none).
func borderXML(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return ""
	}
	l := strings.ToLower(s)
	if l == "none" || l == "0" || strings.HasPrefix(l, "0 ") || strings.HasPrefix(l, "0px") || strings.HasPrefix(l, "hidden") {
		return "none"
	}
	m := borderRe.FindStringSubmatch(s)
	if m == nil {
		return `w:val="single" w:sz="4" w:space="0" w:color="auto"`
	}
	wd, _ := strconv.ParseFloat(m[1], 64)
	if strings.ToLower(m[2]) == "px" {
		wd *= 0.75
	}
	val := map[string]string{"solid": "single", "double": "double", "dashed": "dashed", "dotted": "dotted",
		"inset": "inset", "outset": "outset", "groove": "threeDEngrave", "ridge": "threeDEmboss"}[strings.ToLower(m[3])]
	sz := int(math.Round(wd * 8))
	if val == "double" {
		sz = int(math.Round(wd * 8 / 3))
	}
	if sz < 2 {
		sz = 2
	}
	if sz > 96 {
		sz = 96
	}
	c := parseColor(m[4])
	if c == "" {
		c = "auto"
	}
	return fmt.Sprintf(`w:val="%s" w:sz="%d" w:space="0" w:color="%s"`, val, sz, c)
}

func (w *writer) table(n *hnode, pp pProps, rp rProps) {
	var rows []*hnode
	var caption *hnode
	var collect func(kids []*hnode)
	collect = func(kids []*hnode) {
		for _, k := range kids {
			switch k.tag {
			case "tr":
				rows = append(rows, k)
			case "thead", "tbody", "tfoot":
				collect(k.kids)
			case "caption":
				caption = k
			}
		}
	}
	collect(n.kids)
	if caption != nil {
		w.blockLike(caption, pProps{jc: "center"}, rProps{b: true})
	}
	type cell struct {
		n             *hnode
		col, span, rs int
		cont          bool // vMerge continuation
	}
	var grid [][]cell
	occupied := map[[2]int]*cell{} // (row, col) -> the cell spanning down into it
	ncols := 0
	styled := false
	for _, r := range rows {
		ri := len(grid)
		var cells []cell
		col := 0
		place := func() {
			for {
				o := occupied[[2]int{ri, col}]
				if o == nil {
					return
				}
				cells = append(cells, cell{n: o.n, col: col, span: o.span, cont: true})
				col += o.span
			}
		}
		for _, c := range r.kids {
			if c.tag != "td" && c.tag != "th" {
				continue
			}
			place()
			span := attrInt(c, "colspan", 1)
			if span < 1 {
				span = 1
			}
			if span > 63 {
				span = 63
			}
			rs := attrInt(c, "rowspan", 1)
			if rs < 1 {
				rs = 1
			}
			if strings.Contains(c.attr("style"), "border") {
				styled = true
			}
			cl := cell{n: c, col: col, span: span, rs: rs}
			cells = append(cells, cl)
			for k := 1; k < rs && k < 1000; k++ {
				cp := cl
				occupied[[2]int{ri + k, col}] = &cp
			}
			col += span
		}
		place()
		if len(cells) == 0 {
			continue
		}
		if col > ncols {
			ncols = col
		}
		grid = append(grid, cells)
	}
	if len(grid) == 0 {
		return
	}
	// column widths from cell CSS (single-span cells), remainder shared equally
	tcss := parseCSS(n.attr("style"))
	totalW := 9360
	if tw, ok := cssTwips(tcss["width"]); ok && tw > 0 {
		totalW = tw
	}
	colW := make([]int, ncols)
	for _, cells := range grid {
		for _, c := range cells {
			if c.cont || c.span != 1 || colW[c.col] > 0 {
				continue
			}
			if v, ok := cssTwips(parseCSS(c.n.attr("style"))["width"]); ok && v > 0 {
				colW[c.col] = v
			}
		}
	}
	known, unknown := 0, 0
	for _, v := range colW {
		if v > 0 {
			known += v
		} else {
			unknown++
		}
	}
	if unknown > 0 {
		rest := totalW - known
		if rest < unknown*300 {
			rest = unknown * 1000
		}
		for i := range colW {
			if colW[i] == 0 {
				colW[i] = rest / unknown
			}
		}
	}
	sum := 0
	for _, v := range colW {
		sum += v
	}

	w.out.WriteString(`<w:tbl><w:tblPr>`)
	if strings.HasSuffix(strings.TrimSpace(tcss["width"]), "%") {
		f, _ := strconv.ParseFloat(strings.TrimSuffix(strings.TrimSpace(tcss["width"]), "%"), 64)
		fmt.Fprintf(w.out, `<w:tblW w:w="%d" w:type="pct"/>`, int(math.Round(f*50)))
	} else if styled {
		fmt.Fprintf(w.out, `<w:tblW w:w="%d" w:type="dxa"/>`, sum)
	} else {
		w.out.WriteString(`<w:tblW w:w="0" w:type="auto"/>`)
	}
	ml := strings.TrimSpace(tcss["margin-left"])
	if ml == "auto" && strings.TrimSpace(tcss["margin-right"]) == "auto" {
		w.out.WriteString(`<w:jc w:val="center"/>`)
	} else if ml == "auto" {
		w.out.WriteString(`<w:jc w:val="right"/>`)
	} else if v, ok := cssTwips(ml); ok && v != 0 {
		fmt.Fprintf(w.out, `<w:tblInd w:w="%d" w:type="dxa"/>`, v)
	} else if pp.indent > 0 {
		fmt.Fprintf(w.out, `<w:tblInd w:w="%d" w:type="dxa"/>`, pp.indent)
	}
	w.out.WriteString(`<w:tblBorders>`)
	bval := `w:val="single" w:sz="4" w:space="0" w:color="auto"`
	if styled {
		bval = `w:val="nil"`
	}
	for _, side := range []string{"top", "left", "bottom", "right", "insideH", "insideV"} {
		w.out.WriteString(`<w:` + side + ` ` + bval + `/>`)
	}
	w.out.WriteString(`</w:tblBorders>`)
	if styled {
		w.out.WriteString(`<w:tblLayout w:type="fixed"/>`)
	}
	w.out.WriteString(`<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>`)
	for _, v := range colW {
		fmt.Fprintf(w.out, `<w:gridCol w:w="%d"/>`, v)
	}
	w.out.WriteString(`</w:tblGrid>`)
	for ri, cells := range grid {
		w.out.WriteString("<w:tr>")
		if css := parseCSS(rows[min(ri, len(rows)-1)].attr("style")); css["height"] != "" {
			if v, ok := cssTwips(css["height"]); ok && v > 0 {
				fmt.Fprintf(w.out, `<w:trPr><w:trHeight w:val="%d"/></w:trPr>`, v)
			}
		}
		used := 0
		for _, c := range cells {
			cw := 0
			for g := c.col; g < c.col+c.span && g < len(colW); g++ {
				cw += colW[g]
			}
			vm := ""
			if c.cont {
				vm = "cont"
			} else if c.rs > 1 {
				vm = "restart"
			}
			w.cell(c.n, c.span, cw, vm, rp)
			used = c.col + c.span
		}
		if used < ncols {
			cw := 0
			for g := used; g < ncols; g++ {
				cw += colW[g]
			}
			w.cell(nil, ncols-used, cw, "", rp)
		}
		w.out.WriteString("</w:tr>")
	}
	w.out.WriteString("</w:tbl>")
	w.blocks++
	w.lastTable = true
}

func (w *writer) cell(n *hnode, span, width int, vm string, rp rProps) {
	fmt.Fprintf(w.out, `<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>`, width)
	if span > 1 {
		fmt.Fprintf(w.out, `<w:gridSpan w:val="%d"/>`, span)
	}
	switch vm {
	case "restart":
		w.out.WriteString(`<w:vMerge w:val="restart"/>`)
	case "cont":
		w.out.WriteString(`<w:vMerge/>`)
	}
	var css map[string]string
	if n != nil {
		css = parseCSS(n.attr("style"))
		var bd strings.Builder
		for _, side := range []string{"top", "left", "bottom", "right"} {
			v := css["border-"+side]
			if v == "" {
				v = css["border"]
			}
			if b := borderXML(v); b != "" {
				if b == "none" {
					b = `w:val="nil"`
				}
				bd.WriteString(`<w:` + side + ` ` + b + `/>`)
			}
		}
		if bd.Len() > 0 {
			w.out.WriteString("<w:tcBorders>" + bd.String() + "</w:tcBorders>")
		}
		if c := parseColor(css["background-color"]); c != "" {
			w.out.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="` + c + `"/>`)
		}
		if p := strings.Fields(css["padding"]); len(p) > 0 {
			vals := make([]int, 4)
			for i := range vals {
				src := p[0]
				switch len(p) {
				case 2:
					src = p[i%2]
				case 3:
					src = p[[]int{0, 1, 2, 1}[i]]
				case 4:
					src = p[i]
				}
				vals[i], _ = cssTwips(src)
			}
			fmt.Fprintf(w.out, `<w:tcMar><w:top w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/><w:bottom w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tcMar>`,
				vals[0], vals[3], vals[2], vals[1])
		}
		switch strings.ToLower(css["vertical-align"]) {
		case "middle":
			w.out.WriteString(`<w:vAlign w:val="center"/>`)
		case "bottom":
			w.out.WriteString(`<w:vAlign w:val="bottom"/>`)
		}
	}
	w.out.WriteString("</w:tcPr>")
	before := w.blocks
	if n != nil && vm != "cont" {
		pp := pProps{jc: alignOf(n, css)}
		if pp.jc == "left" {
			pp.jc = ""
		}
		if n.tag == "th" {
			rp.b = true
		}
		delete(css, "background-color")
		delete(css, "vertical-align")
		rp = applyInlineStyle(n, css, rp)
		w.container(n.kids, pp, rp)
		w.flush()
	}
	if w.blocks == before || w.lastTable {
		w.out.WriteString("<w:p/>")
		w.lastTable = false
	}
	w.out.WriteString("</w:tc>")
}

// ---- inline ----

var namedColors = map[string]string{
	"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
	"yellow": "FFFF00", "cyan": "00FFFF", "aqua": "00FFFF", "magenta": "FF00FF", "fuchsia": "FF00FF",
	"gray": "808080", "grey": "808080", "silver": "C0C0C0", "maroon": "800000", "olive": "808000",
	"lime": "00FF00", "teal": "008080", "navy": "000080", "purple": "800080", "orange": "FFA500",
	"pink": "FFC0CB", "brown": "A52A2A", "darkred": "8B0000", "darkblue": "00008B",
	"darkgreen": "006400", "lightgray": "D3D3D3", "lightgrey": "D3D3D3", "gold": "FFD700",
}

var rgbRe = regexp.MustCompile(`^rgba?\(\s*([\d.]+)%?\s*[, ]\s*([\d.]+)%?\s*[, ]\s*([\d.]+)%?\s*(?:[,/]\s*([\d.]+%?))?\s*\)$`)

// parseColor returns RRGGBB (uppercase) or "" for transparent/unknown.
func parseColor(s string) string {
	s = strings.ToLower(strings.TrimSpace(s))
	if s == "" || s == "transparent" || s == "inherit" || s == "initial" {
		return ""
	}
	if c, ok := namedColors[s]; ok {
		return c
	}
	if strings.HasPrefix(s, "#") {
		h := s[1:]
		if len(h) == 3 || len(h) == 4 {
			h = string([]byte{h[0], h[0], h[1], h[1], h[2], h[2]})
		}
		if len(h) == 8 {
			h = h[:6]
		}
		if len(h) == 6 && hexRe.MatchString(h) {
			return strings.ToUpper(h)
		}
		return ""
	}
	if m := rgbRe.FindStringSubmatch(s); m != nil {
		if m[4] != "" {
			a, _ := strconv.ParseFloat(strings.TrimSuffix(m[4], "%"), 64)
			if a == 0 {
				return ""
			}
		}
		var out [3]int
		for i := 0; i < 3; i++ {
			f, _ := strconv.ParseFloat(m[i+1], 64)
			out[i] = int(math.Max(0, math.Min(255, math.Round(f))))
		}
		return fmt.Sprintf("%02X%02X%02X", out[0], out[1], out[2])
	}
	return ""
}

var fontKeywordPt = map[string]float64{
	"xx-small": 7, "x-small": 7.5, "small": 10, "medium": 12, "large": 13.5,
	"x-large": 18, "xx-large": 24, "xxx-large": 36,
}

// parseFontSize returns half-points, or 0.
func parseFontSize(s string) int {
	s = strings.ToLower(strings.TrimSpace(s))
	if pt, ok := fontKeywordPt[s]; ok {
		return int(pt * 2)
	}
	num := func(suf string) (float64, bool) {
		if !strings.HasSuffix(s, suf) {
			return 0, false
		}
		f, err := strconv.ParseFloat(strings.TrimSpace(strings.TrimSuffix(s, suf)), 64)
		return f, err == nil && f > 0
	}
	if f, ok := num("pt"); ok {
		return clampSz(f * 2)
	}
	if f, ok := num("px"); ok {
		return clampSz(f * 0.75 * 2)
	}
	if f, ok := num("rem"); ok {
		return clampSz(f * 12 * 2)
	}
	if f, ok := num("em"); ok {
		return clampSz(f * 12 * 2)
	}
	return 0
}

func clampSz(hp float64) int {
	n := int(math.Round(hp))
	if n < 2 {
		n = 2
	}
	if n > 3276 {
		n = 3276
	}
	return n
}

// cssTwips converts a CSS length (also 0 and negative values) to twips.
func cssTwips(s string) (int, bool) {
	s = strings.ToLower(strings.TrimSpace(s))
	if s == "" {
		return 0, false
	}
	if s == "0" {
		return 0, true
	}
	for _, u := range []struct {
		suf string
		mul float64
	}{{"pt", 20}, {"px", 15}, {"in", 1440}, {"cm", 567}, {"mm", 56.7}, {"rem", 240}, {"em", 240}} {
		if strings.HasSuffix(s, u.suf) {
			f, err := strconv.ParseFloat(strings.TrimSpace(strings.TrimSuffix(s, u.suf)), 64)
			if err == nil {
				return int(math.Round(f * u.mul)), true
			}
			return 0, false
		}
	}
	return 0, false
}

// cssLength converts a positive CSS length to twips (0 if unknown).
func cssLength(s string) int {
	if v, ok := cssTwips(s); ok && v > 0 {
		return v
	}
	return 0
}

var fontTagSize = map[string]int{"1": 16, "2": 20, "3": 24, "4": 28, "5": 36, "6": 48, "7": 72}

func applyInlineStyle(n *hnode, css map[string]string, rp rProps) rProps {
	if c := parseColor(css["color"]); c != "" {
		rp.color = c
	}
	bg := css["background-color"]
	if bg == "" {
		bg = css["background"]
	}
	if c := parseColor(bg); c != "" {
		rp.bg = c
	}
	if sz := parseFontSize(css["font-size"]); sz > 0 {
		rp.sz = sz
	}
	if f := css["font-family"]; f != "" {
		f = strings.TrimSpace(strings.Split(f, ",")[0])
		f = strings.Trim(f, `"' `)
		if f != "" && !genericFonts[strings.ToLower(f)] {
			rp.font = f
		}
	}
	switch fw := strings.ToLower(css["font-weight"]); fw {
	case "bold", "bolder", "600", "700", "800", "900":
		rp.b, rp.noB = true, false
	case "normal", "400", "lighter", "100", "200", "300":
		rp.b, rp.noB = false, true
	}
	switch strings.ToLower(css["font-style"]) {
	case "italic", "oblique":
		rp.i, rp.noI = true, false
	case "normal":
		rp.i, rp.noI = false, true
	}
	td := strings.ToLower(css["text-decoration"] + " " + css["text-decoration-line"])
	if strings.Contains(td, "underline") {
		rp.u = true
	}
	if strings.Contains(td, "line-through") {
		rp.s = true
	}
	switch strings.ToLower(css["vertical-align"]) {
	case "super":
		rp.sup, rp.sub = true, false
	case "sub":
		rp.sub, rp.sup = true, false
	}
	switch strings.ToLower(css["text-transform"]) {
	case "uppercase":
		rp.caps = true
	case "none":
		rp.caps = false
	}
	switch strings.ToLower(css["font-variant"]) {
	case "small-caps":
		rp.smallCaps = true
	case "normal":
		rp.smallCaps = false
	}
	if v, ok := cssTwips(css["letter-spacing"]); ok {
		rp.spacing = v
	}
	if ws := strings.ToLower(css["white-space"]); strings.HasPrefix(ws, "pre") {
		rp.pre = true
	}
	if n.tag == "font" {
		if c := parseColor(n.attr("color")); c != "" {
			rp.color = c
		}
		if f := n.attr("face"); f != "" {
			rp.font = strings.Trim(strings.TrimSpace(strings.Split(f, ",")[0]), `"'`)
		}
		if sz, ok := fontTagSize[strings.TrimSpace(n.attr("size"))]; ok {
			rp.sz = sz
		}
	}
	return rp
}

var genericFonts = map[string]bool{"serif": true, "sans-serif": true, "monospace": true,
	"cursive": true, "fantasy": true, "system-ui": true, "inherit": true, "initial": true}

func (w *writer) inline(n *hnode, pp pProps, rp rProps) {
	if n.tag == "" {
		w.text(n.text, pp, rp)
		return
	}
	if skipTags[n.tag] {
		return
	}
	if blockTags[n.tag] {
		w.flush()
		w.block(n, pp, rp)
		return
	}
	css := parseCSS(n.attr("style"))
	switch n.tag {
	case "b", "strong":
		rp.b, rp.noB = true, false
	case "i", "em", "cite", "var":
		rp.i, rp.noI = true, false
	case "u", "ins":
		rp.u = true
	case "s", "strike", "del":
		rp.s = true
	case "sup":
		rp.sup, rp.sub = true, false
	case "sub":
		rp.sub, rp.sup = true, false
	case "code", "kbd", "samp", "tt":
		rp.font = "Courier New"
	case "mark":
		rp.bg = "FFFF00"
	case "br":
		w.ensure(pp)
		w.addRun(`<w:r>`+runPr(rp)+`<w:br/></w:r>`, true)
		w.cur.lastSpace = true
		return
	case "img":
		w.image(n, pp)
		return
	case "a":
		w.link(n, pp, applyInlineStyle(n, css, rp))
		return
	}
	if strings.TrimSpace(strings.ToLower(css["float"])) == "right" {
		delete(css, "float") // right-tab emulation; the tab itself is kept as text
	}
	rp = applyInlineStyle(n, css, rp)
	for _, k := range n.kids {
		w.inline(k, pp, rp)
	}
}

func (w *writer) link(n *hnode, pp pProps, rp rProps) {
	href := strings.TrimSpace(n.attr("href"))
	if href == "" || !safeHref(href) {
		for _, k := range n.kids {
			w.inline(k, pp, rp)
		}
		return
	}
	p := w.ensure(pp)
	start := len(p.runs)
	// links converted from .docx carry their own explicit colour/underline
	rp.link = !(w.page.v2 && hasStyledDesc(n))
	for _, k := range n.kids {
		w.inline(k, pp, rp)
	}
	if w.cur != p || len(p.runs) == start {
		return // link content was split by blocks or empty
	}
	inner := strings.Join(p.runs[start:], "")
	var open string
	if strings.HasPrefix(href, "#") {
		open = `<w:hyperlink w:anchor="` + xesc(href[1:]) + `" w:history="1">`
	} else {
		if u, err := url.Parse(href); err != nil || u.String() == "" {
			return
		}
		id := w.addRel(relHL, href, true)
		open = `<w:hyperlink r:id="` + id + `" w:history="1">`
	}
	p.runs = append(p.runs[:start], open+inner+"</w:hyperlink>")
	p.brRun = append(p.brRun[:start], false)
}

func (w *writer) text(t string, pp pProps, rp rProps) {
	if !rp.pre && !pp.pre {
		var sb strings.Builder
		space := false
		for _, r := range t {
			if r == ' ' || r == '\t' || r == '\n' || r == '\r' || r == '\f' {
				space = true
				continue
			}
			if space {
				sb.WriteByte(' ')
				space = false
			}
			sb.WriteRune(r)
		}
		if space {
			sb.WriteByte(' ')
		}
		t = sb.String()
		if t == "" || (t == " " && w.cur == nil) {
			return
		}
		p := w.ensure(pp)
		if p.lastSpace {
			t = strings.TrimPrefix(t, " ")
		}
		if t == "" {
			return
		}
		p.lastSpace = strings.HasSuffix(t, " ")
	} else {
		w.ensure(pp)
		w.cur.lastSpace = false
		t = strings.ReplaceAll(t, "\r\n", "\n")
	}
	var body strings.Builder
	seg := ""
	emit := func() {
		if seg != "" {
			body.WriteString(`<w:t xml:space="preserve">` + xesc(seg) + `</w:t>`)
			seg = ""
		}
	}
	for _, r := range t {
		switch r {
		case '\t':
			emit()
			body.WriteString("<w:tab/>")
		case '\n':
			emit()
			body.WriteString("<w:br/>")
		default:
			seg += string(r)
		}
	}
	emit()
	if body.Len() > 0 {
		w.addRun("<w:r>"+runPr(rp)+body.String()+"</w:r>", false)
	}
}

func runPr(rp rProps) string {
	var sb strings.Builder
	if rp.link {
		sb.WriteString(`<w:rStyle w:val="Hyperlink"/>`)
	}
	if rp.font != "" {
		f := xesc(rp.font)
		sb.WriteString(`<w:rFonts w:ascii="` + f + `" w:hAnsi="` + f + `" w:eastAsia="` + f + `" w:cs="` + f + `"/>`)
	}
	if rp.b {
		sb.WriteString("<w:b/>")
	} else if rp.noB {
		sb.WriteString(`<w:b w:val="0"/>`)
	}
	if rp.i {
		sb.WriteString("<w:i/>")
	} else if rp.noI {
		sb.WriteString(`<w:i w:val="0"/>`)
	}
	if rp.caps {
		sb.WriteString("<w:caps/>")
	}
	if rp.smallCaps {
		sb.WriteString("<w:smallCaps/>")
	}
	if rp.s {
		sb.WriteString("<w:strike/>")
	}
	if rp.color != "" {
		sb.WriteString(`<w:color w:val="` + rp.color + `"/>`)
	}
	if rp.spacing != 0 {
		fmt.Fprintf(&sb, `<w:spacing w:val="%d"/>`, rp.spacing)
	}
	if rp.sz > 0 {
		fmt.Fprintf(&sb, `<w:sz w:val="%d"/><w:szCs w:val="%d"/>`, rp.sz, rp.sz)
	}
	if rp.u {
		sb.WriteString(`<w:u w:val="single"/>`)
	}
	if rp.bg != "" {
		sb.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="` + rp.bg + `"/>`)
	}
	if rp.sup {
		sb.WriteString(`<w:vertAlign w:val="superscript"/>`)
	} else if rp.sub {
		sb.WriteString(`<w:vertAlign w:val="subscript"/>`)
	}
	if sb.Len() == 0 {
		return ""
	}
	return "<w:rPr>" + sb.String() + "</w:rPr>"
}

// xesc escapes text for XML and drops characters not allowed in XML 1.0.
func xesc(s string) string {
	var sb strings.Builder
	for _, r := range s {
		switch {
		case r == '&':
			sb.WriteString("&amp;")
		case r == '<':
			sb.WriteString("&lt;")
		case r == '>':
			sb.WriteString("&gt;")
		case r == '"':
			sb.WriteString("&quot;")
		case r == '\t' || r == '\n' || r == '\r':
			sb.WriteRune(r)
		case r < 0x20 || r == 0xFFFE || r == 0xFFFF || r == utf8.RuneError || (r >= 0xD800 && r <= 0xDFFF):
		default:
			sb.WriteRune(r)
		}
	}
	return sb.String()
}

var dataURIRe = regexp.MustCompile(`^data:(image/(?:png|jpeg|jpg|gif));base64,`)

func (w *writer) image(n *hnode, pp pProps) {
	src := strings.TrimSpace(n.attr("src"))
	m := dataURIRe.FindStringSubmatch(strings.ToLower(src[:min(len(src), 40)]))
	if m == nil {
		if alt := n.attr("alt"); alt != "" {
			w.text("["+alt+"]", pp, rProps{})
		}
		return
	}
	data, err := base64.StdEncoding.DecodeString(strings.Map(func(r rune) rune {
		if r == ' ' || r == '\n' || r == '\r' || r == '\t' {
			return -1
		}
		return r
	}, src[len(m[0]):]))
	if err != nil || len(data) == 0 {
		return
	}
	cfg, format, err := image.DecodeConfig(bytes.NewReader(data))
	if err != nil {
		return
	}
	ext := map[string]string{"png": "png", "jpeg": "jpeg", "gif": "gif"}[format]
	if ext == "" {
		return
	}
	css := parseCSS(n.attr("style"))
	wpx, hpx := pxAttr(n.attr("width"), css["width"]), pxAttr(n.attr("height"), css["height"])
	switch {
	case wpx > 0 && hpx > 0:
	case wpx > 0 && cfg.Width > 0:
		hpx = wpx * float64(cfg.Height) / float64(cfg.Width)
	case hpx > 0 && cfg.Height > 0:
		wpx = hpx * float64(cfg.Width) / float64(cfg.Height)
	default:
		wpx, hpx = float64(cfg.Width), float64(cfg.Height)
	}
	// fit within the text width of the page
	maxW := float64(w.page.w-w.page.left-w.page.right) / 15
	if w.page.landscape && w.page.h > w.page.w {
		maxW = float64(w.page.h-w.page.left-w.page.right) / 15
	}
	if maxW < 100 {
		maxW = 624
	}
	if wpx > maxW {
		hpx = hpx * maxW / wpx
		wpx = maxW
	}
	if wpx < 1 {
		wpx = 1
	}
	if hpx < 1 {
		hpx = 1
	}
	name := fmt.Sprintf("image%d.%s", len(w.media)+1, ext)
	w.media = append(w.media, [2]string{name, string(data)})
	rid := w.addRel(relIM, "media/"+name, false)
	w.drawID++
	cx, cy := int64(math.Round(wpx*9525)), int64(math.Round(hpx*9525))
	alt := xesc(n.attr("alt"))
	id := w.drawID
	run := fmt.Sprintf(`<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`+
		`<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`+
		`<wp:docPr id="%d" name="Picture %d" descr="%s"/>`+
		`<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`+
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`+
		`<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s" descr="%s"/><pic:cNvPicPr/></pic:nvPicPr>`+
		`<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`+
		`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`+
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`+
		`</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`,
		cx, cy, id, id, alt, id, name, alt, rid, cx, cy)
	p := w.ensure(pp)
	p.lastSpace = false
	w.addRun(run, false)
}

func pxAttr(attr, css string) float64 {
	for _, s := range []string{css, attr} {
		s = strings.ToLower(strings.TrimSpace(s))
		if s == "" || strings.HasSuffix(s, "%") || s == "auto" {
			continue
		}
		mul := 1.0
		switch {
		case strings.HasSuffix(s, "px"):
			s = strings.TrimSuffix(s, "px")
		case strings.HasSuffix(s, "pt"):
			s, mul = strings.TrimSuffix(s, "pt"), 4.0/3
		case strings.HasSuffix(s, "in"):
			s, mul = strings.TrimSuffix(s, "in"), 96
		case strings.HasSuffix(s, "cm"):
			s, mul = strings.TrimSuffix(s, "cm"), 96/2.54
		}
		if f, err := strconv.ParseFloat(strings.TrimSpace(s), 64); err == nil && f > 0 {
			return f * mul
		}
	}
	return 0
}

// ---- packaging ----

func (w *writer) pack(title string) ([]byte, error) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	add := func(name, content string) error {
		f, err := zw.Create(name)
		if err != nil {
			return err
		}
		_, err = f.Write([]byte(content))
		return err
	}
	const hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"
	var ct strings.Builder
	ct.WriteString(hdr + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`<Default Extension="png" ContentType="image/png"/>` +
		`<Default Extension="jpeg" ContentType="image/jpeg"/>` +
		`<Default Extension="gif" ContentType="image/gif"/>` +
		`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
		`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
		`<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
		`<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
		`<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
		`<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`)
	for _, p := range w.parts {
		ct.WriteString(`<Override PartName="/word/` + p.name + `" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.` + p.kind + `+xml"/>`)
	}
	ct.WriteString(`</Types>`)
	rootRels := hdr + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
		`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
		`</Relationships>`
	relsXML := func(rels []outRel, base string) string {
		var dr strings.Builder
		dr.WriteString(hdr + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + base)
		for _, r := range rels {
			mode := ""
			if r.external {
				mode = ` TargetMode="External"`
			}
			dr.WriteString(`<Relationship Id="` + r.id + `" Type="` + r.typ + `" Target="` + xesc(r.target) + `"` + mode + `/>`)
		}
		dr.WriteString(`</Relationships>`)
		return dr.String()
	}
	var sectRefs strings.Builder
	for i, p := range w.parts {
		id := fmt.Sprintf("rIdHF%d", i+1)
		typ, tag := relHD, "headerReference"
		if p.kind == "footer" {
			typ, tag = relFT, "footerReference"
		}
		w.rels = append(w.rels, outRel{id: id, typ: typ, target: p.name})
		sectRefs.WriteString(`<w:` + tag + ` w:type="default" r:id="` + id + `"/>`)
	}
	docRels := relsXML(w.rels,
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`+
			`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`+
			`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`)

	pg := w.page
	orient := ""
	if pg.landscape {
		orient = ` w:orient="landscape"`
	}
	sect := fmt.Sprintf(`<w:sectPr>%s<w:pgSz w:w="%d" w:h="%d"%s/><w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="%d" w:footer="%d" w:gutter="0"/></w:sectPr>`,
		sectRefs.String(), pg.w, pg.h, orient, pg.top, pg.right, pg.bottom, pg.left, pg.header, pg.footer)
	doc := hdr + `<w:document` + nsDecl + `><w:body>` + w.out.String() + sect + `</w:body></w:document>`

	now := time.Now().UTC().Format("2006-01-02T15:04:05Z")
	core := hdr + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
		` xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"` +
		` xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
		`<dc:title>` + xesc(title) + `</dc:title><dc:creator>DocStudio</dc:creator>` +
		`<dcterms:created xsi:type="dcterms:W3CDTF">` + now + `</dcterms:created>` +
		`<dcterms:modified xsi:type="dcterms:W3CDTF">` + now + `</dcterms:modified></cp:coreProperties>`
	app := hdr + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>DocStudio</Application></Properties>`
	settings := hdr + `<w:settings xmlns:w="` + nsW + `">` + fmt.Sprintf(`<w:defaultTabStop w:val="%d"/>`, pg.tab) +
		`<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`

	parts := [][2]string{
		{"[Content_Types].xml", ct.String()}, {"_rels/.rels", rootRels},
		{"word/document.xml", doc}, {"word/_rels/document.xml.rels", docRels},
		{"word/styles.xml", stylesXML(pg.font, pg.sz, pg.v2)}, {"word/numbering.xml", w.numberingXML()},
		{"word/settings.xml", settings}, {"docProps/core.xml", core}, {"docProps/app.xml", app},
	}
	for _, p := range w.parts {
		parts = append(parts, [2]string{"word/" + p.name, p.xml})
		if len(p.rels) > 0 {
			parts = append(parts, [2]string{"word/_rels/" + p.name + ".rels", relsXML(p.rels, "")})
		}
	}
	for _, p := range parts {
		if err := add(p[0], p[1]); err != nil {
			return nil, err
		}
	}
	for _, m := range w.media {
		if err := add("word/media/"+m[0], m[1]); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func (w *writer) numberingXML() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" + `<w:numbering xmlns:w="` + nsW + `">`)
	for _, n := range w.nums {
		fmt.Fprintf(&sb, `<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="hybridMultilevel"/>`, n.id)
		for l, lv := range n.levels {
			font := ""
			if lv.font != "" {
				f := xesc(lv.font)
				font = `<w:rPr><w:rFonts w:ascii="` + f + `" w:hAnsi="` + f + `" w:cs="` + f + `" w:hint="default"/></w:rPr>`
			}
			fmt.Fprintf(&sb, `<w:lvl w:ilvl="%d"><w:start w:val="%d"/><w:numFmt w:val="%s"/><w:lvlText w:val="%s"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr>%s</w:lvl>`,
				l, lv.start, lv.fmt, xesc(lv.text), lv.left, lv.hanging, font)
		}
		sb.WriteString(`</w:abstractNum>`)
	}
	for _, n := range w.nums {
		fmt.Fprintf(&sb, `<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>`, n.id, n.id)
	}
	sb.WriteString(`</w:numbering>`)
	return sb.String()
}

// stylesXML builds styles.xml. neutral (documents that carry explicit
// formatting everywhere) keeps heading styles free of fonts/colours/sizes so
// nothing leaks into runs that rely on the document defaults.
func stylesXML(font string, sz int, neutral bool) string {
	if font == "" {
		font = "Calibri"
	}
	if sz <= 0 {
		sz = 22
	}
	f := xesc(font)
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" + `<w:styles xmlns:w="` + nsW + `">`)
	fmt.Fprintf(&sb, `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>`+
		`<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`, f, f, f, f, sz, sz)
	sb.WriteString(`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`)
	sizes := []int{32, 26, 24, 22, 22, 22}
	for i := 1; i <= 6; i++ {
		ital := ""
		if i >= 4 {
			ital = "<w:i/>"
		}
		if neutral {
			fmt.Fprintf(&sb, `<w:style w:type="paragraph" w:styleId="Heading%d"><w:name w:val="heading %d"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>`+
				`<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="%d"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>`, i, i, i-1)
			continue
		}
		fmt.Fprintf(&sb, `<w:style w:type="paragraph" w:styleId="Heading%d"><w:name w:val="heading %d"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>`+
			`<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="%d"/></w:pPr>`+
			`<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:b/>%s<w:color w:val="1F3864"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:style>`,
			i, i, i-1, ital, sizes[i-1], sizes[i-1])
	}
	sb.WriteString(`<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
		`<w:pPr><w:spacing w:after="80"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>`)
	sb.WriteString(`<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>`)
	sb.WriteString(`<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>`)
	sb.WriteString(`<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`)
	sb.WriteString(`<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>`)
	sb.WriteString(`</w:styles>`)
	return sb.String()
}

func hasStyledDesc(n *hnode) bool {
	for _, k := range n.kids {
		if k.tag == "" {
			continue
		}
		if k.attr("style") != "" || k.tag == "u" || hasStyledDesc(k) {
			return true
		}
	}
	return false
}
