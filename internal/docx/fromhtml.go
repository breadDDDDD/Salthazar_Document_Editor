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
)

type pProps struct {
	style       string
	jc          string
	numID, ilvl int
	indent      int // left indent, twips
	pre         bool
	border      bool // bottom border (horizontal rule)
}

type rProps struct {
	b, i, u, s, sup, sub, link, pre bool
	color, bg, font                 string // colors as RRGGBB
	sz                              int    // half-points
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

type numInst struct{ id, abs, ilvl int }

type writer struct {
	out       strings.Builder
	cur       *para
	blocks    int
	lastTable bool
	rels      []outRel
	media     [][2]string // name, data
	nums      []numInst
	bulletNum int
	drawID    int
	firstHead string
}

// FromHTML converts an HTML fragment or document to a .docx file.
func FromHTML(src string) ([]byte, error) {
	root, title := parseHTML(src)
	w := &writer{}
	w.container(root.kids, pProps{}, rProps{})
	w.flush()
	if w.blocks == 0 || w.lastTable {
		w.out.WriteString("<w:p/>")
	}
	if title == "" {
		title = w.firstHead
	}
	return w.pack(title)
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
	if n := len(p.runs); p.brRun[n-1] {
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
	if pp.border {
		ppr.WriteString(`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>`)
	}
	if pp.pre {
		ppr.WriteString(`<w:spacing w:after="0"/>`)
	}
	if pp.indent > 0 && pp.numID == 0 {
		fmt.Fprintf(&ppr, `<w:ind w:left="%d"/>`, pp.indent)
	}
	if pp.jc != "" {
		ppr.WriteString(`<w:jc w:val="` + pp.jc + `"/>`)
	}
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
		w.writePara(pp, "")
	}
}

func (w *writer) block(n *hnode, pp pProps, rp rProps) {
	if skipTags[n.tag] {
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
	if ml := cssLength(css["margin-left"]); ml > 0 && pp.numID == 0 {
		pp.indent += ml
	} else if pl := cssLength(css["padding-left"]); pl > 0 && pp.numID == 0 && n.tag != "ul" && n.tag != "ol" {
		pp.indent += pl
	}
	rp = applyInlineStyle(n, css, rp)
	switch n.tag {
	case "h1", "h2", "h3", "h4", "h5", "h6":
		pp.style = "Heading" + n.tag[1:]
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
		if strings.Contains(" "+n.attr("class")+" ", " page-break ") ||
			strings.Contains(strings.ToLower(n.attr("style")), "page-break") {
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
		w.table(n, rp)
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

func (w *writer) bullet() int {
	if w.bulletNum == 0 {
		w.bulletNum = len(w.nums) + 1
		w.nums = append(w.nums, numInst{id: w.bulletNum, abs: 0, ilvl: -1})
	}
	return w.bulletNum
}

func (w *writer) list(n *hnode, pp pProps, rp rProps) {
	level := 0
	if pp.numID > 0 {
		level = pp.ilvl + 1
	}
	if level > 8 {
		level = 8
	}
	numID := 0
	if n.tag == "ul" {
		numID = w.bullet()
	} else {
		numID = len(w.nums) + 1
		w.nums = append(w.nums, numInst{id: numID, abs: 1, ilvl: level})
	}
	lp := pProps{style: "ListParagraph", numID: numID, ilvl: level, jc: pp.jc}
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

func (w *writer) table(n *hnode, rp rProps) {
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
		n    *hnode
		span int
	}
	var grid [][]cell
	ncols := 0
	for _, r := range rows {
		var cells []cell
		cols := 0
		for _, c := range r.kids {
			if c.tag != "td" && c.tag != "th" {
				continue
			}
			span, _ := strconv.Atoi(c.attr("colspan"))
			if span < 1 {
				span = 1
			}
			if span > 63 {
				span = 63
			}
			cells = append(cells, cell{c, span})
			cols += span
		}
		if len(cells) == 0 {
			continue
		}
		if cols > ncols {
			ncols = cols
		}
		grid = append(grid, cells)
	}
	if len(grid) == 0 {
		return
	}
	colW := 9360 / ncols
	w.out.WriteString(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>`)
	for _, side := range []string{"top", "left", "bottom", "right", "insideH", "insideV"} {
		w.out.WriteString(`<w:` + side + ` w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
	}
	w.out.WriteString(`</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>`)
	for i := 0; i < ncols; i++ {
		fmt.Fprintf(&w.out, `<w:gridCol w:w="%d"/>`, colW)
	}
	w.out.WriteString(`</w:tblGrid>`)
	for _, cells := range grid {
		w.out.WriteString("<w:tr>")
		used := 0
		for _, c := range cells {
			w.cell(c.n, c.span, colW, rp)
			used += c.span
		}
		if used < ncols {
			w.cell(nil, ncols-used, colW, rp)
		}
		w.out.WriteString("</w:tr>")
	}
	w.out.WriteString("</w:tbl>")
	w.blocks++
	w.lastTable = true
}

func (w *writer) cell(n *hnode, span, colW int, rp rProps) {
	fmt.Fprintf(&w.out, `<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>`, colW*span)
	if span > 1 {
		fmt.Fprintf(&w.out, `<w:gridSpan w:val="%d"/>`, span)
	}
	w.out.WriteString("</w:tcPr>")
	before := w.blocks
	if n != nil {
		css := parseCSS(n.attr("style"))
		pp := pProps{jc: alignOf(n, css)}
		if pp.jc == "left" {
			pp.jc = ""
		}
		if n.tag == "th" {
			rp.b = true
		}
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

// cssLength converts a CSS length to twips (0 if unknown).
func cssLength(s string) int {
	s = strings.ToLower(strings.TrimSpace(s))
	conv := map[string]float64{"pt": 20, "px": 15, "in": 1440, "cm": 567, "mm": 56.7, "em": 240}
	for suf, mul := range conv {
		if strings.HasSuffix(s, suf) {
			f, err := strconv.ParseFloat(strings.TrimSpace(strings.TrimSuffix(s, suf)), 64)
			if err == nil && f > 0 {
				return int(math.Round(f * mul))
			}
		}
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
		rp.b = true
	case "normal", "400", "lighter", "100", "200", "300":
		rp.b = false
	}
	switch strings.ToLower(css["font-style"]) {
	case "italic", "oblique":
		rp.i = true
	case "normal":
		rp.i = false
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
		rp.b = true
	case "i", "em", "cite", "var":
		rp.i = true
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
	rp.link = true
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
		sb.WriteString(`<w:rFonts w:ascii="` + f + `" w:hAnsi="` + f + `" w:cs="` + f + `"/>`)
	}
	if rp.b {
		sb.WriteString("<w:b/>")
	}
	if rp.i {
		sb.WriteString("<w:i/>")
	}
	if rp.s {
		sb.WriteString("<w:strike/>")
	}
	if rp.color != "" {
		sb.WriteString(`<w:color w:val="` + rp.color + `"/>`)
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
	if wpx > 624 { // fit within 6.5in text width
		hpx = hpx * 624 / wpx
		wpx = 624
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
	ct := hdr + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
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
		`<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
		`</Types>`
	rootRels := hdr + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
		`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
		`</Relationships>`
	var dr strings.Builder
	dr.WriteString(hdr + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
		`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`)
	for _, r := range w.rels {
		mode := ""
		if r.external {
			mode = ` TargetMode="External"`
		}
		dr.WriteString(`<Relationship Id="` + r.id + `" Type="` + r.typ + `" Target="` + xesc(r.target) + `"` + mode + `/>`)
	}
	dr.WriteString(`</Relationships>`)

	doc := hdr + `<w:document xmlns:w="` + nsW + `" xmlns:r="` + nsR + `"` +
		` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>` +
		w.out.String() +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>` +
		`</w:body></w:document>`

	now := time.Now().UTC().Format("2006-01-02T15:04:05Z")
	core := hdr + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
		` xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"` +
		` xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
		`<dc:title>` + xesc(title) + `</dc:title><dc:creator>DocStudio</dc:creator>` +
		`<dcterms:created xsi:type="dcterms:W3CDTF">` + now + `</dcterms:created>` +
		`<dcterms:modified xsi:type="dcterms:W3CDTF">` + now + `</dcterms:modified></cp:coreProperties>`
	app := hdr + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>DocStudio</Application></Properties>`
	settings := hdr + `<w:settings xmlns:w="` + nsW + `"><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`

	parts := [][2]string{
		{"[Content_Types].xml", ct}, {"_rels/.rels", rootRels},
		{"word/document.xml", doc}, {"word/_rels/document.xml.rels", dr.String()},
		{"word/styles.xml", stylesXML}, {"word/numbering.xml", w.numberingXML()},
		{"word/settings.xml", settings}, {"docProps/core.xml", core}, {"docProps/app.xml", app},
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
	bullets := []string{"•", "o", "▪"}
	sb.WriteString(`<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>`)
	for l := 0; l < 9; l++ {
		font := ""
		if bullets[l%3] == "o" {
			font = `<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/></w:rPr>`
		}
		fmt.Fprintf(&sb, `<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="%s"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr>%s</w:lvl>`,
			l, bullets[l%3], 720*(l+1), font)
	}
	sb.WriteString(`</w:abstractNum>`)
	formats := []string{"decimal", "lowerLetter", "lowerRoman"}
	sb.WriteString(`<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>`)
	for l := 0; l < 9; l++ {
		fmt.Fprintf(&sb, `<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/><w:lvlText w:val="%%%d."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr></w:lvl>`,
			l, formats[l%3], l+1, 720*(l+1))
	}
	sb.WriteString(`</w:abstractNum>`)
	for _, n := range w.nums {
		fmt.Fprintf(&sb, `<w:num w:numId="%d"><w:abstractNumId w:val="%d"/>`, n.id, n.abs)
		if n.ilvl >= 0 {
			fmt.Fprintf(&sb, `<w:lvlOverride w:ilvl="%d"><w:startOverride w:val="1"/></w:lvlOverride>`, n.ilvl)
		}
		sb.WriteString(`</w:num>`)
	}
	sb.WriteString(`</w:numbering>`)
	return sb.String()
}

var stylesXML = func() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" + `<w:styles xmlns:w="` + nsW + `">`)
	sb.WriteString(`<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>` +
		`<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`)
	sb.WriteString(`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`)
	sizes := []int{32, 26, 24, 22, 22, 22}
	for i := 1; i <= 6; i++ {
		ital := ""
		if i >= 4 {
			ital = "<w:i/>"
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
}()
