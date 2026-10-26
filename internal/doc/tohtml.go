package doc

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"html"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf16"
)

// ToHTML converts a Word 97-2003 .doc file to the same HTML dialect that
// internal/docx ToHTML produces. It returns the HTML and the document title.
func ToHTML(b []byte) (h string, title string, err error) {
	defer func() {
		if r := recover(); r != nil {
			h, title, err = "", "", errors.New("cannot read .doc file: the file is damaged or uses unsupported features")
		}
	}()
	return toHTML(b)
}

func toHTML(b []byte) (string, string, error) {
	if len(b) >= 4 && string(b[:4]) == "PK\x03\x04" {
		// A .docx saved with a .doc extension.
		if _, err := zip.NewReader(bytes.NewReader(b), int64(len(b))); err == nil {
			return docxFallback(b)
		}
	}
	if !isCFB(b) {
		head := strings.ToLower(string(b[:min(len(b), 512)]))
		switch {
		case strings.HasPrefix(head, "{\\rtf"):
			return "", "", errors.New("this .doc file is really an RTF document; open it in Word and save it as .docx")
		case strings.Contains(head, "<html") || strings.Contains(head, "mime-version"):
			return "", "", errors.New("this .doc file is really a web page (HTML/MHT); open it in Word and save it as .docx")
		}
		return "", "", errors.New("not a Word 97-2003 .doc file")
	}
	cf, err := openCFB(b)
	if err != nil {
		return "", "", errors.New("cannot read .doc file: " + err.Error())
	}
	if !cf.has("WordDocument") {
		switch {
		case cf.has("EncryptedPackage"):
			return "", "", errEncrypted
		case cf.has("Workbook") || cf.has("Book"):
			return "", "", errors.New("this file is an Excel workbook, not a Word document")
		case cf.has("PowerPoint Document"):
			return "", "", errors.New("this file is a PowerPoint presentation, not a Word document")
		}
		return "", "", errors.New("this compound file is not a Word document")
	}
	d, err := parseWord(cf)
	if err != nil {
		return "", "", err
	}
	c := newConv(d)
	out := c.document()
	title := summaryTitle(cf.stream("\x05SummaryInformation"))
	if title == "" {
		title = c.firstHeading
	}
	return out, title, nil
}

const pageBreakMark = "\x00PB\x00"

type field struct {
	sep   bool
	instr []uint16
	link  bool
}

type listFrame struct {
	tag  string
	key  int
	left int
}

type blockWriter struct {
	out   strings.Builder
	lists []listFrame
}

func (w *blockWriter) closeLists() {
	for len(w.lists) > 0 {
		w.popList()
	}
}

func (w *blockWriter) popList() {
	t := w.lists[len(w.lists)-1]
	w.lists = w.lists[:len(w.lists)-1]
	w.out.WriteString("</li></" + t.tag + ">")
}

type cellOut struct {
	html string
}

type rowOut struct {
	cells []string
	td    *tdef
}

type conv struct {
	d            *wordDoc
	fields       []field
	linkOpen     bool
	firstHeading string
	chpCache     map[[3]int]chp
	counters     map[int32]*[9]int
	started      map[int32]*[9]bool
	pending      []string // blocks to emit after the current paragraph (text boxes)
	notes        []string
	ftnRef       []int
	ftnTxt       []int
	endRef       []int
	endTxt       []int
	noteNum      string // number to print for the 0x02 mark inside a note story
	story        int    // 0 main, 1 other
	usedTxbx     map[int]bool
	normal       chp
}

func newConv(d *wordDoc) *conv {
	c := &conv{d: d, chpCache: map[[3]int]chp{}, counters: map[int32]*[9]int{},
		started: map[int32]*[9]bool{}, usedTxbx: map[int]bool{}}
	_, c.normal = d.resolveStyle(0)
	readCPs := func(i int, stride int) []int {
		b := d.tblBytes(i)
		if len(b) < 4 {
			return nil
		}
		n := (len(b) - 4) / (4 + stride)
		if stride == 0 {
			n = len(b)/4 - 1
		}
		var cps []int
		for k := 0; k <= n; k++ {
			cps = append(cps, i32(b, 4*k))
		}
		return cps
	}
	c.ftnRef = readCPs(iPlcffndRef, 2)
	c.ftnTxt = readCPs(iPlcffndTxt, 0)
	c.endRef = readCPs(iPlcfendRef, 2)
	c.endTxt = readCPs(iPlcfendTxt, 0)
	return c
}

func pt(twips int) string {
	return strconv.FormatFloat(float64(twips)/20, 'f', -1, 64) + "pt"
}

func (c *conv) document() string {
	d := c.d
	s := defaultSep()
	if len(d.sections) > 0 {
		s = d.sections[0].sep
	}
	orient := "portrait"
	if s.landscape || s.pageW > s.pageH {
		orient = "landscape"
	}
	var sb strings.Builder
	fmt.Fprintf(&sb, `<div class="docx-doc" data-page-w="%d" data-page-h="%d" data-margin-top="%d" data-margin-right="%d" data-margin-bottom="%d" data-margin-left="%d" data-header-dist="%d" data-footer-dist="%d" data-orient="%s"`,
		s.pageW, s.pageH, abs(s.top), s.right, abs(s.bottom), s.left, s.hdr, s.ftr, orient)
	if f := d.fontName(c.normal.ftc); f != "" {
		sb.WriteString(` data-font="` + html.EscapeString(f) + `"`)
	}
	fmt.Fprintf(&sb, ` data-font-size="%s" data-tab="%s">`, strconv.FormatFloat(float64(c.normal.hps)/2, 'f', -1, 64),
		strconv.FormatFloat(float64(d.dxaTab)/20, 'f', -1, 64))

	c.story = 1
	if h := c.headerFooter(7, 11); h != "" {
		sb.WriteString(`<header class="docx-header">` + h + `</header>`)
	}
	if f := c.headerFooter(9, 12); f != "" {
		sb.WriteString(`<footer class="docx-footer">` + f + `</footer>`)
	}
	c.story = 0
	c.fields = nil
	sb.WriteString(c.render(0, d.ccp[0]))
	c.story = 1
	if len(c.notes) > 0 {
		sb.WriteString(`<div class="docx-footnotes">` + strings.Join(c.notes, "") + `</div>`)
	}
	sb.WriteString("</div>")
	return sb.String()
}

// headerFooter renders the first section's odd (or first-page) header/footer.
func (c *conv) headerFooter(odd, first int) string {
	d := c.d
	b := d.tblBytes(iPlcfHdd)
	if len(b) < 8 {
		return ""
	}
	n := len(b)/4 - 1
	base := d.ccp[0] + d.ccp[1]
	for _, i := range []int{odd, first - 1} {
		if i+1 > n {
			continue
		}
		s, e := i32(b, 4*i), i32(b, 4*(i+1))
		if e-s <= 1 {
			continue
		}
		// Each header/footer story ends with an extra guard paragraph mark.
		if ch := d.chars(base+e-2, base+e); len(ch) == 2 && ch[0].r == 0x0D && ch[1].r == 0x0D {
			e--
		}
		c.fields = nil
		h := c.render(base+s, base+e)
		if strings.TrimSpace(stripTags(h)) != "" || strings.Contains(h, "<img") {
			return h
		}
	}
	return ""
}

var tagRe = regexp.MustCompile(`<[^>]*>`)

func stripTags(s string) string { return html.UnescapeString(tagRe.ReplaceAllString(s, "")) }

func abs(v int) int {
	if v < 0 {
		return -v
	}
	return v
}

// render converts the story range [cpStart, cpEnd) to block HTML.
func (c *conv) render(cpStart, cpEnd int) string {
	d := c.d
	chars := d.chars(cpStart, cpEnd)
	w := &blockWriter{}
	var rows []rowOut
	var cur rowOut
	var cell *blockWriter
	flushTable := func() {
		if cell != nil && cell.out.Len() > 0 {
			cell.closeLists()
			cur.cells = append(cur.cells, cell.out.String())
		}
		if len(cur.cells) > 0 {
			rows = append(rows, cur)
		}
		if len(rows) > 0 {
			w.closeLists()
			w.out.WriteString(c.table(rows))
		}
		rows, cur, cell = nil, rowOut{}, nil
	}
	start := 0
	for i := 0; i < len(chars); i++ {
		ch := chars[i].r
		secBreak, bkc := false, 0
		if ch == 0x0C {
			secBreak, bkc = d.sectionBreakAt(cpStart + i)
		}
		if ch != 0x0D && ch != 0x07 && !secBreak {
			continue
		}
		para := chars[start:i]
		end := chars[i]
		start = i + 1
		p, stChp := c.papAt(end)
		if p.itap >= 1 && (p.inTable || p.itap > 0) && ch != 0x0C {
			if p.ttp || (p.itap == 1 && p.innerTtp) {
				if cell != nil && cell.out.Len() > 0 {
					cell.closeLists()
					cur.cells = append(cur.cells, cell.out.String())
				}
				cell = nil
				cur.td = p.tdef
				rows = append(rows, cur)
				cur = rowOut{}
				continue
			}
			if p.itap > 1 && p.innerTtp {
				continue // end of a nested row: nested tables are flattened
			}
			if cell == nil {
				cell = &blockWriter{}
			}
			c.paragraph(cell, para, p, stChp, true)
			c.flushPending(cell)
			if ch == 0x07 && p.itap <= 1 {
				cell.closeLists()
				cur.cells = append(cur.cells, cell.out.String())
				cell = nil
			}
			continue
		}
		if len(rows) > 0 || len(cur.cells) > 0 || cell != nil {
			flushTable()
		}
		if p.pageBreakBefore && w.out.Len() > 0 {
			w.closeLists()
			w.out.WriteString(`<hr class="page-break">`)
		}
		if !(secBreak && len(para) == 0) {
			c.paragraph(w, para, p, stChp, false)
			c.flushPending(w)
		}
		if secBreak && bkc != 0 {
			w.closeLists()
			w.out.WriteString(`<hr class="page-break">`)
		}
	}
	if start < len(chars) {
		end := chars[len(chars)-1]
		p, stChp := c.papAt(end)
		p.inTable, p.itap = false, 0
		if len(rows) > 0 || cell != nil {
			flushTable()
		}
		c.paragraph(w, chars[start:], p, stChp, false)
		c.flushPending(w)
	}
	if len(rows) > 0 || len(cur.cells) > 0 || cell != nil {
		flushTable()
	}
	w.closeLists()
	return w.out.String()
}

func (c *conv) flushPending(w *blockWriter) {
	if len(c.pending) == 0 {
		return
	}
	w.closeLists()
	for _, p := range c.pending {
		w.out.WriteString(p)
	}
	c.pending = nil
}

// papAt computes the paragraph properties for the paragraph ending at ch.
func (c *conv) papAt(ch dchar) (pap, chp) {
	d := c.d
	istd := 0
	var g []byte
	if i := findRun(d.papRuns, ch.fc); i >= 0 {
		istd, g = d.papRuns[i].istd, d.papRuns[i].grpprl
	}
	sp, sc := d.resolveStyle(istd)
	pg := d.pieceGrpprl(ch.piece)
	p := sp
	p.tdef = nil
	p.istd = istd
	d.applyPapx(&p, g)
	d.applyPapx(&p, pg)
	if l := d.listFor(p.ilfo); l != nil && p.ilvl >= 0 && p.ilvl < len(l.lvls) {
		// Level indents sit between the style and direct formatting.
		q := sp
		q.tdef = nil
		q.istd = istd
		d.applyPapx(&q, l.lvls[p.ilvl].papx)
		d.applyPapx(&q, g)
		d.applyPapx(&q, pg)
		p = q
	}
	return p, sc
}

// chpAt computes the character properties of ch in a paragraph whose style
// character properties are st.
func (c *conv) chpAt(ch dchar, istd int, st chp) chp {
	d := c.d
	ri := findRun(d.chpRuns, ch.fc)
	key := [3]int{ri, ch.piece, istd}
	if v, ok := c.chpCache[key]; ok {
		return v
	}
	var g []byte
	if ri >= 0 {
		g = d.chpRuns[ri].grpprl
	}
	base := st
	eachSprm(g, func(op uint16, a []byte) {
		if op == 0x4A30 {
			c.applyCharStyle(&base, int(argU(a)), 0)
		}
	})
	v := base
	d.applyChpx(&v, g, base)
	d.applyChpx(&v, d.pieceGrpprl(ch.piece), base)
	c.chpCache[key] = v
	return v
}

func (c *conv) applyCharStyle(x *chp, istd, depth int) {
	s := c.d.style(istd)
	if s == nil || s.stk != 2 || depth > 10 {
		return
	}
	if s.base != 0x0FFF && s.base != istd {
		c.applyCharStyle(x, s.base, depth+1)
	}
	c.d.applyChpx(x, s.chpx, *x)
}

func headingLevel(s *style) (int, bool) {
	if s == nil {
		return 0, false
	}
	if s.sti >= 1 && s.sti <= 9 {
		return min(s.sti, 6), false
	}
	n := strings.ToLower(strings.TrimSpace(s.name))
	if strings.HasPrefix(n, "heading ") {
		if v, err := strconv.Atoi(strings.TrimSpace(n[8:])); err == nil && v >= 1 && v <= 9 {
			return min(v, 6), false
		}
	}
	if s.sti == 62 || n == "title" {
		return 1, true
	}
	return 0, false
}

func (c *conv) fontCSS(ftc int) string {
	return cssFontName(c.d.fontName(ftc))
}

func cssFontName(f string) string {
	f = strings.Map(func(r rune) rune {
		if r == '"' || r == '\'' || r == ';' || r == '<' || r == '>' || r == '&' || r < 0x20 {
			return -1
		}
		return r
	}, f)
	if strings.ContainsRune(f, ' ') {
		return "'" + f + "'"
	}
	return f
}

func fnum(v float64) string { return strconv.FormatFloat(v, 'f', -1, 64) }

func (c *conv) blockCSS(p pap, st chp, heading, listItem bool) []string {
	var css []string
	css = append(css, "margin-top:"+pt(p.before), "margin-bottom:"+pt(p.after))
	if !listItem {
		if p.left != 0 {
			css = append(css, "margin-left:"+pt(p.left))
		}
		if p.left1 != 0 {
			css = append(css, "text-indent:"+pt(p.left1))
		}
	}
	if p.right != 0 {
		css = append(css, "margin-right:"+pt(p.right))
	}
	switch p.jc {
	case 1:
		css = append(css, "text-align:center")
	case 2:
		css = append(css, "text-align:right")
	case 3, 4, 5, 6, 7, 8, 9:
		css = append(css, "text-align:justify")
	}
	if p.shade != "" {
		css = append(css, "background-color:"+p.shade)
	}
	switch {
	case !p.hasLine || (p.lineMult && p.line == 0):
		css = append(css, "line-height:1.15")
	case p.lineMult:
		css = append(css, "line-height:"+fnum(math.Round(1150*float64(p.line)/240)/1000))
	default:
		css = append(css, "line-height:"+pt(abs(p.line)))
	}
	if f := c.fontCSS(st.ftc); f != "" {
		css = append(css, "font-family:"+f)
	}
	css = append(css, "font-size:"+fnum(float64(st.hps)/2)+"pt")
	if st.color != "" {
		css = append(css, "color:"+st.color)
	} else {
		css = append(css, "color:inherit")
	}
	if heading {
		if st.b {
			css = append(css, "font-weight:bold")
		} else {
			css = append(css, "font-weight:normal")
		}
		if st.i {
			css = append(css, "font-style:italic")
		}
	}
	return css
}

// ---- inline content ----

type runKey struct {
	b, i, u, s, sup, sub bool
	css                  string
}

type segment struct {
	k       runKey
	content string
	raw     bool
}

type inline struct {
	segs []segment
}

func (in *inline) add(k runKey, s string) {
	if n := len(in.segs); n > 0 && !in.segs[n-1].raw && in.segs[n-1].k == k {
		in.segs[n-1].content += s
		return
	}
	in.segs = append(in.segs, segment{k: k, content: s})
}

func (in *inline) raw(s string) { in.segs = append(in.segs, segment{content: s, raw: true}) }

func (in *inline) html() string {
	var sb strings.Builder
	for _, s := range in.segs {
		if s.raw {
			sb.WriteString(s.content)
			continue
		}
		open, close := "", ""
		wrap := func(o, c string) { open += o; close = c + close }
		if s.k.css != "" {
			wrap(`<span style="`+html.EscapeString(s.k.css)+`">`, "</span>")
		}
		if s.k.b {
			wrap("<strong>", "</strong>")
		}
		if s.k.i {
			wrap("<em>", "</em>")
		}
		if s.k.u {
			wrap("<u>", "</u>")
		}
		if s.k.s {
			wrap("<s>", "</s>")
		}
		if s.k.sup {
			wrap("<sup>", "</sup>")
		} else if s.k.sub {
			wrap("<sub>", "</sub>")
		}
		sb.WriteString(open + s.content + close)
	}
	return sb.String()
}

func (c *conv) runKey(x, st chp) runKey {
	k := runKey{b: x.b && !st.b, i: x.i && !st.i, u: x.u, s: x.strike, sup: x.iss == 1, sub: x.iss == 2}
	var css []string
	if x.color != st.color {
		if x.color == "" {
			css = append(css, "color:inherit")
		} else {
			css = append(css, "color:"+x.color)
		}
	}
	if x.highlight != "" {
		css = append(css, "background-color:"+x.highlight)
	} else if x.shade != "" {
		css = append(css, "background-color:"+x.shade)
	}
	if x.hps != st.hps && x.hps > 0 {
		css = append(css, "font-size:"+fnum(float64(x.hps)/2)+"pt")
	}
	if x.ftc != st.ftc {
		if f := c.fontCSS(x.ftc); f != "" && c.d.fontName(x.ftc) != c.d.fontName(st.ftc) {
			css = append(css, "font-family:"+f)
		}
	}
	if !x.b && st.b {
		css = append(css, "font-weight:normal")
	}
	if !x.i && st.i {
		css = append(css, "font-style:normal")
	}
	if x.caps {
		css = append(css, "text-transform:uppercase")
	}
	if x.smallCaps {
		css = append(css, "font-variant:small-caps")
	}
	if abs(x.spacing) > 10 { // ignore sub-half-point kerning tweaks
		css = append(css, "letter-spacing:"+pt(x.spacing))
	}
	k.css = strings.Join(css, ";")
	return k
}

var symbolMap = map[uint16]rune{
	0xB7: '•', 0xD7: '×', 0xB0: '°', 0xB1: '±', 0xA3: '≤', 0xB3: '≥', 0xB9: '≠', 0xBB: '≈', 0xAE: '→',
	0xAC: '←', 0xAD: '↑', 0xAF: '↓', 0xDE: '⇒', 0xDC: '⇐', 0xA5: '∞', 0xD6: '√', 0xE5: '∑', 0xF2: '∫',
	0xB6: '∂', 0xC6: '∅', 0xCE: '∈', 0x22: '∀', 0x24: '∃', 0xB4: '÷', 0xA2: '′', 0xB2: '″', 0xBC: '…',
	0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε', 0x66: 'φ', 0x67: 'γ', 0x68: 'η', 0x69: 'ι',
	0x6B: 'κ', 0x6C: 'λ', 0x6D: 'μ', 0x6E: 'ν', 0x6F: 'ο', 0x70: 'π', 0x71: 'θ', 0x72: 'ρ', 0x73: 'σ',
	0x74: 'τ', 0x75: 'υ', 0x77: 'ω', 0x78: 'ξ', 0x79: 'ψ', 0x7A: 'ζ', 0x44: 'Δ', 0x46: 'Φ', 0x47: 'Γ',
	0x4C: 'Λ', 0x50: 'Π', 0x51: 'Θ', 0x53: 'Σ', 0x57: 'Ω', 0x59: 'Ψ',
}

var wingdingsMap = map[uint16]rune{
	0xFC: '✓', 0xFB: '✗', 0xFD: '☒', 0xFE: '☑', 0xA8: '◻', 0x6F: '□', 0x6E: '■', 0xA7: '▪', 0x9F: '•',
	0x6C: '●', 0xA1: '○', 0xD8: '➢', 0xE0: '➔', 0xE8: '➔', 0x76: '❖', 0x71: '❑', 0x75: '◆', 0x77: '⬥',
	0x4A: '☺', 0x4C: '☹', 0x28: '☎', 0x2A: '✉', 0x3F: '✍', 0x46: '☞', 0xAB: '★', 0x9E: '·',
}

func mapSymbol(font string, ch uint16) (rune, bool) {
	code := ch
	if code >= 0xF000 && code <= 0xF0FF {
		code -= 0xF000
	}
	f := strings.ToLower(font)
	switch {
	case f == "symbol":
		if r, ok := symbolMap[code]; ok {
			return r, true
		}
		return rune(code), true
	case strings.HasPrefix(f, "wingdings"):
		if r, ok := wingdingsMap[code]; ok {
			return r, true
		}
	}
	if ch >= 0xF000 && ch <= 0xF0FF {
		return rune(code), false
	}
	return rune(ch), false
}

func (c *conv) hidden() bool {
	for _, f := range c.fields {
		if !f.sep {
			return true
		}
	}
	return false
}

var hrefRe = regexp.MustCompile(`(?i)HYPERLINK\s+(?:\\l\s+)?"([^"]*)"|(?i)HYPERLINK\s+(\S+)`)

func fieldHref(instr string) string {
	s := strings.TrimSpace(instr)
	if !strings.HasPrefix(strings.ToUpper(s), "HYPERLINK") {
		return ""
	}
	anchor := strings.Contains(s, `\l`)
	m := hrefRe.FindStringSubmatch(s)
	if m == nil {
		return ""
	}
	h := m[1]
	if h == "" {
		h = m[2]
	}
	if anchor && !strings.Contains(h, "://") {
		h = "#" + h
	}
	if l := strings.ToLower(strings.TrimSpace(h)); h == "" || strings.HasPrefix(l, "javascript:") || strings.HasPrefix(l, "vbscript:") || strings.HasPrefix(l, "data:") {
		return ""
	}
	return h
}

func (c *conv) inlineHTML(chars []dchar, istd int, st chp) string {
	in := &inline{}
	d := c.d
	for idx := 0; idx < len(chars); idx++ {
		ch := chars[idx]
		switch ch.r {
		case 0x13:
			c.fields = append(c.fields, field{})
			continue
		case 0x14:
			if n := len(c.fields); n > 0 && !c.fields[n-1].sep {
				f := &c.fields[n-1]
				f.sep = true
				if !c.hidden() {
					if h := fieldHref(string(utf16.Decode(f.instr))); h != "" && !c.linkOpen {
						in.raw(`<a href="` + html.EscapeString(h) + `">`)
						f.link, c.linkOpen = true, true
					}
				}
			}
			continue
		case 0x15:
			if n := len(c.fields); n > 0 {
				f := c.fields[n-1]
				c.fields = c.fields[:n-1]
				if f.link && c.linkOpen {
					in.raw("</a>")
					c.linkOpen = false
				}
			}
			continue
		}
		if c.hidden() {
			f := &c.fields[len(c.fields)-1]
			if !f.sep {
				f.instr = append(f.instr, ch.r)
			}
			continue
		}
		x := c.chpAt(ch, istd, st)
		if x.vanish || x.del {
			continue
		}
		k := c.runKey(x, st)
		switch r := ch.r; {
		case r == 0x09:
			in.add(k, `<span style="white-space:pre">`+"\t</span>")
		case r == 0x0B:
			in.add(k, "<br>")
		case r == 0x0C:
			in.raw(pageBreakMark)
		case r == 0x0E:
			in.add(k, "<br>")
		case r == 0x1E:
			in.add(k, "‑")
		case r == 0x1F:
		case r == 0x01 && x.fSpec:
			if x.fData || x.picLoc < 0 {
				continue
			}
			in.raw(imgHTML(d.inlinePicture(x.picLoc)))
		case r == 0x08 && x.fSpec:
			c.drawnObject(in, ch, idx, chars)
		case r == 0x02 && x.fSpec:
			c.noteRef(in, ch)
		case r == 0x28 && x.fSpec && x.symChar != 0:
			font := d.fontName(x.symFtc)
			sr, mapped := mapSymbol(font, x.symChar)
			if mapped {
				in.add(k, html.EscapeString(string(sr)))
			} else {
				kk := k
				kk.css = strings.Trim(kk.css+";font-family:"+cssFontName(font), ";")
				in.add(kk, html.EscapeString(string(sr)))
			}
		case r < 0x20:
			// other control characters (annotation marks, separators...)
		case r >= 0xD800 && r <= 0xDBFF && idx+1 < len(chars):
			lo := chars[idx+1].r
			if lo >= 0xDC00 && lo <= 0xDFFF {
				in.add(k, html.EscapeString(string(utf16.DecodeRune(rune(r), rune(lo)))))
				idx++
			}
		case r >= 0xD800 && r <= 0xDFFF, r == 0xFFFF, r == 0xFFFE:
		case r >= 0xF000 && r <= 0xF0FF:
			sr, _ := mapSymbol(d.fontName(x.ftc), r)
			in.add(k, html.EscapeString(string(sr)))
		default:
			in.add(k, html.EscapeString(string(rune(r))))
		}
	}
	if c.linkOpen {
		in.raw("</a>")
		c.linkOpen = false
		for i := range c.fields {
			c.fields[i].link = false
		}
	}
	return in.html()
}

func imgHTML(p picture) string {
	if !p.ok {
		label := "picture"
		if p.kind != "" {
			label = p.kind + " picture"
		}
		return `<span style="color:#808080">[` + label + `]</span>`
	}
	attrs := ""
	if p.w > 0 && p.h > 0 {
		attrs = fmt.Sprintf(` width="%d" height="%d"`, p.w, p.h)
	}
	return `<img src="` + p.url + `"` + attrs + ` alt="">`
}

func (c *conv) drawnObject(in *inline, ch dchar, idx int, chars []dchar) {
	d := c.d
	if c.story != 0 {
		return
	}
	cp := c.cpOf(ch)
	k := sort.SearchInts(d.spaCP, cp)
	if k >= len(d.spa) || d.spaCP[k] != cp {
		return
	}
	f := d.spa[k]
	s := d.shapes[f.spid]
	if tb := c.textboxIndex(f.spid, s); tb >= 0 {
		if c.usedTxbx[tb] {
			return
		}
		c.usedTxbx[tb] = true
		base := d.ccp[0] + d.ccp[1] + d.ccp[2] + d.ccp[4] + d.ccp[5]
		saved, savedStory, savedLink := c.fields, c.story, c.linkOpen
		c.fields, c.story, c.linkOpen = nil, 1, false
		body := c.render(base+d.txbxCP[tb], base+d.txbxCP[tb+1])
		c.fields, c.story, c.linkOpen = saved, savedStory, savedLink
		css := "width:" + pt(f.right-f.left)
		if s != nil && !s.noLine {
			css += ";border:1px solid #000000"
		}
		c.pending = append(c.pending, `<div class="docx-textbox" style="`+css+`">`+body+`</div>`)
		return
	}
	if s != nil && s.pib > 0 {
		in.raw(imgHTML(d.floatingPicture(k)))
	}
}

// cpOf maps a character back to its CP.
func (c *conv) cpOf(ch dchar) int {
	p := c.d.pieces[ch.piece]
	if p.compressed {
		return p.cpStart + int(ch.fc-p.fc)
	}
	return p.cpStart + int(ch.fc-p.fc)/2
}

func (c *conv) textboxIndex(spid uint32, s *shapeInfo) int {
	d := c.d
	n := len(d.txbxCP) - 1
	if n <= 0 {
		return -1
	}
	for i, lid := range d.txbxLid {
		if uint32(lid) == spid && i < n {
			return i
		}
	}
	if s != nil && s.txid != 0 {
		if i := int(s.txid>>16) - 1; i >= 0 && i < n {
			return i
		}
	}
	return -1
}

func (c *conv) noteRef(in *inline, ch dchar) {
	d := c.d
	if c.story != 0 {
		if c.noteNum != "" {
			in.raw("<sup>" + c.noteNum + "</sup>")
		}
		return
	}
	cp := c.cpOf(ch)
	try := func(ref, txt []int, base int, label func(int) string) bool {
		if len(ref) < 2 {
			return false
		}
		k := sort.SearchInts(ref[:len(ref)-1], cp)
		if k >= len(ref)-1 || ref[k] != cp || k+1 >= len(txt) {
			return false
		}
		num := label(k + 1)
		in.raw("<sup>" + num + "</sup>")
		saved, savedLink := c.fields, c.linkOpen
		c.fields, c.story, c.noteNum, c.linkOpen = nil, 1, num, false
		body := c.render(base+txt[k], base+txt[k+1])
		c.fields, c.story, c.noteNum, c.linkOpen = saved, 0, "", savedLink
		c.notes = append(c.notes, body)
		return true
	}
	if try(c.ftnRef, c.ftnTxt, d.ccp[0], strconv.Itoa) {
		return
	}
	try(c.endRef, c.endTxt, d.ccp[0]+d.ccp[1]+d.ccp[2]+d.ccp[4], func(n int) string { return strings.ToLower(roman(n)) })
}

// ---- paragraphs and lists ----

func (c *conv) paragraph(w *blockWriter, chars []dchar, p pap, st chp, inCell bool) {
	content := c.inlineHTML(chars, p.istd, st)
	level, isTitle := headingLevel(c.d.style(p.istd))
	l := c.d.listFor(p.ilfo)
	if l != nil && (p.ilvl < 0 || p.ilvl >= len(l.lvls)) {
		l = nil
	}
	if l != nil && level == 0 {
		c.listItem(w, l, p, st, content)
		return
	}
	w.closeLists()
	label := ""
	if l != nil {
		label, _ = c.nextLabel(l, p.ilvl)
	}
	tag := "p"
	if level > 0 {
		tag = "h" + strconv.Itoa(level)
		if c.firstHeading == "" && c.story == 0 {
			c.firstHeading = strings.TrimSpace(stripTags(strings.ReplaceAll(content, pageBreakMark, "")))
		}
	}
	css := c.blockCSS(p, st, level > 0, false)
	attrs := ` style="` + html.EscapeString(strings.Join(css, ";")) + `"`
	if isTitle {
		attrs = ` class="title"` + attrs
	}
	if label != "" {
		content = html.EscapeString(label) + " " + content
	}
	parts := strings.Split(content, pageBreakMark)
	for i, part := range parts {
		if i > 0 {
			w.out.WriteString(`<hr class="page-break">`)
		}
		if strings.TrimSpace(part) == "" {
			if len(parts) > 1 {
				continue
			}
			part = "<br>"
		}
		w.out.WriteString("<" + tag + attrs + ">" + part + "</" + tag + ">")
	}
}

var nfcNames = map[int]string{0: "decimal", 1: "upperRoman", 2: "lowerRoman", 3: "upperLetter", 4: "lowerLetter",
	5: "ordinal", 22: "decimalZero", 23: "bullet", 255: "none"}

var nfcCSS = map[int]string{0: "decimal", 1: "upper-roman", 2: "lower-roman", 3: "upper-alpha", 4: "lower-alpha",
	22: "decimal-leading-zero", 255: "none"}

func roman(n int) string {
	if n <= 0 || n > 3999 {
		return strconv.Itoa(n)
	}
	vals := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	syms := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
	var sb strings.Builder
	for i, v := range vals {
		for n >= v {
			sb.WriteString(syms[i])
			n -= v
		}
	}
	return sb.String()
}

func letters(n int) string {
	if n <= 0 {
		return strconv.Itoa(n)
	}
	ch := byte('A' + (n-1)%26)
	return strings.Repeat(string(ch), (n-1)/26+1)
}

func formatNum(n, nfc int) string {
	switch nfc {
	case 1:
		return roman(n)
	case 2:
		return strings.ToLower(roman(n))
	case 3:
		return letters(n)
	case 4:
		return strings.ToLower(letters(n))
	case 22:
		if n < 10 {
			return "0" + strconv.Itoa(n)
		}
	case 255:
		return ""
	}
	return strconv.Itoa(n)
}

// nextLabel advances the list counters and returns the item label and the
// counter value at ilvl.
func (c *conv) nextLabel(l *lst, ilvl int) (string, int) {
	cnt, ok := c.counters[l.lsid]
	if !ok {
		cnt = &[9]int{}
		c.counters[l.lsid] = cnt
		c.started[l.lsid] = &[9]bool{}
	}
	st := c.started[l.lsid]
	if !st[ilvl] {
		cnt[ilvl] = l.lvls[ilvl].start
		st[ilvl] = true
	} else {
		cnt[ilvl]++
	}
	for k := ilvl + 1; k < 9; k++ {
		st[k] = false
	}
	v := l.lvls[ilvl]
	if v.nfc == 23 {
		return bulletChar(v.text), cnt[ilvl]
	}
	var sb strings.Builder
	for _, u := range v.text {
		if u < 9 {
			k := int(u)
			n := cnt[k]
			if !st[k] && k < len(l.lvls) {
				n = l.lvls[k].start
			}
			nfc := 0
			if k < len(l.lvls) {
				nfc = l.lvls[k].nfc
			}
			sb.WriteString(formatNum(n, nfc))
		} else {
			sb.WriteRune(rune(u))
		}
	}
	return sb.String(), cnt[ilvl]
}

func bulletChar(t []uint16) string {
	if len(t) == 0 {
		return "•"
	}
	r, _ := mapSymbol("symbol", t[0])
	if t[0] >= 0xF000 {
		if m, ok := wingdingsMap[t[0]-0xF000]; ok && t[0] != 0xF0B7 {
			r = m
		}
	}
	return string(r)
}

func (c *conv) listItem(w *blockWriter, l *lst, p pap, st chp, content string) {
	v := l.lvls[p.ilvl]
	label, n := c.nextLabel(l, p.ilvl)
	tag := "ol"
	if v.nfc == 23 {
		tag = "ul"
	}
	key := int(l.lsid)*16 + p.ilvl
	ilvl := p.ilvl
	for len(w.lists) > ilvl+1 {
		w.popList()
	}
	if len(w.lists) == ilvl+1 && (w.lists[ilvl].tag != tag || w.lists[ilvl].key != key) {
		w.popList()
	}
	if len(w.lists) == ilvl+1 {
		w.out.WriteString("</li>")
	} else {
		for len(w.lists) < ilvl+1 {
			lv := len(w.lists)
			parentLeft := 0
			if lv > 0 {
				parentLeft = w.lists[lv-1].left
			}
			left := p.left
			if lv < ilvl {
				left = parentLeft + 720
			}
			pad := left - parentLeft
			if pad < 0 {
				pad = 0
			}
			lt := "disc"
			numfmt := nfcNames[v.nfc]
			if numfmt == "" {
				numfmt = "decimal"
			}
			if tag == "ol" {
				lt = nfcCSS[v.nfc]
				if lt == "" {
					lt = "decimal"
				}
			} else {
				switch label {
				case "o", "◦":
					lt = "circle"
				case "▪", "■", "§":
					lt = "square"
				}
			}
			open := "<" + tag
			if tag == "ol" {
				open += fmt.Sprintf(` start="%d"`, n)
			}
			lvltext := ""
			for _, u := range v.text {
				if u < 9 {
					lvltext += "%" + strconv.Itoa(int(u)+1)
				} else if v.nfc == 23 {
					lvltext += label
				} else {
					lvltext += string(rune(u))
				}
			}
			open += ` data-numfmt="` + numfmt + `" data-lvltext="` + html.EscapeString(lvltext) + `"`
			open += ` style="list-style-type:` + lt + `;margin-top:0;margin-bottom:0;padding-left:` + pt(pad) + `">`
			w.out.WriteString(open)
			w.lists = append(w.lists, listFrame{tag: tag, key: key, left: left})
			if len(w.lists) < ilvl+1 {
				w.out.WriteString(`<li style="list-style-type:none">`)
			}
		}
	}
	css := c.blockCSS(p, st, false, true)
	if override := labelOverride(tag, label, v, p.ilvl); override != "" {
		css = append(css, `list-style-type:"`+strings.ReplaceAll(override, `"`, `\"`)+` "`)
	}
	if strings.TrimSpace(content) == "" {
		content = "<br>"
	}
	content = strings.ReplaceAll(content, pageBreakMark, "")
	w.out.WriteString(`<li style="` + html.EscapeString(strings.Join(css, ";")) + `">` + content)
}

// labelOverride returns the literal label when CSS list-style-type alone
// cannot reproduce it.
func labelOverride(tag, label string, v lvl, ilvl int) string {
	if tag == "ul" {
		switch label {
		case "•", "o", "◦", "▪", "■", "§", "":
			return ""
		}
		return label
	}
	if len(v.text) == 2 && int(v.text[0]) == ilvl && v.text[1] == '.' && nfcCSS[v.nfc] != "" {
		return ""
	}
	return label
}

// ---- tables ----

func (c *conv) table(rows []rowOut) string {
	// Gather column boundaries across all rows.
	var bounds []int
	for i := range rows {
		td := rows[i].td
		if td == nil || len(td.centers) < len(rows[i].cells)+1 {
			// Synthesize equal-width cells.
			n := len(rows[i].cells)
			nt := &tdef{}
			if td != nil {
				*nt = *td
			}
			nt.centers = nil
			for k := 0; k <= n; k++ {
				nt.centers = append(nt.centers, k*9000/max(n, 1))
			}
			for len(nt.cells) < n {
				nt.cells = append(nt.cells, tc{})
			}
			rows[i].td = nt
		}
		bounds = append(bounds, rows[i].td.centers...)
	}
	sort.Ints(bounds)
	var grid []int
	for _, b := range bounds {
		if len(grid) == 0 || b-grid[len(grid)-1] > 5 {
			grid = append(grid, b)
		}
	}
	col := func(x int) int {
		best, bd := 0, 1<<30
		for i, g := range grid {
			if dd := abs(g - x); dd < bd {
				best, bd = i, dd
			}
		}
		return best
	}
	type cellR struct {
		html            string
		col, span, rows int
		width           int
		t               tc
		skip            bool
		first, last     bool
	}
	out := make([][]*cellR, len(rows))
	open := map[int]*cellR{}
	for ri, r := range rows {
		td := r.td
		var prev *cellR
		for ci, h := range r.cells {
			if ci+1 >= len(td.centers) {
				break
			}
			t := tc{}
			if ci < len(td.cells) {
				t = td.cells[ci]
			}
			c0, c1 := col(td.centers[ci]), col(td.centers[ci+1])
			cr := &cellR{html: h, col: c0, span: max(1, c1-c0), rows: 1, width: td.centers[ci+1] - td.centers[ci], t: t}
			if t.hMerge == 2 && prev != nil {
				prev.span += cr.span
				prev.width += cr.width
				if strings.TrimSpace(stripTags(h)) != "" {
					prev.html += h
				}
				continue
			}
			if t.vMerge == 2 {
				if o := open[c0]; o != nil {
					o.rows++
					cr.skip = true
				}
			} else if t.vMerge == 1 {
				open[c0] = cr
			} else {
				delete(open, c0)
			}
			out[ri] = append(out[ri], cr)
			prev = cr
		}
		if n := len(out[ri]); n > 0 {
			out[ri][0].first = true
			out[ri][n-1].last = true
		}
	}
	var sb strings.Builder
	tblCSS := "border-collapse:collapse"
	if t := rows[0].td; t != nil && len(t.centers) > 0 {
		if ind := t.centers[0] + t.gapHalf; ind > 0 {
			tblCSS += ";margin-left:" + pt(ind)
		}
	}
	sb.WriteString(`<table style="` + tblCSS + `"><tbody>`)
	for ri, row := range out {
		sb.WriteString("<tr>")
		td := rows[ri].td
		for _, cr := range row {
			if cr.skip {
				continue
			}
			attrs := ""
			if cr.span > 1 {
				attrs += fmt.Sprintf(` colspan="%d"`, cr.span)
			}
			if cr.rows > 1 {
				attrs += fmt.Sprintf(` rowspan="%d"`, cr.rows)
			}
			css := []string{"width:" + pt(cr.width)}
			names := []string{"top", "left", "bottom", "right"}
			for k := 0; k < 4; k++ {
				b := cr.t.borders[k]
				if !b.set {
					switch k {
					case 0:
						b = td.tblBrc[4]
						if ri == 0 {
							b = td.tblBrc[0]
						}
					case 2:
						b = td.tblBrc[4]
						if ri+cr.rows >= len(out) {
							b = td.tblBrc[2]
						}
					case 1:
						b = td.tblBrc[5]
						if cr.first {
							b = td.tblBrc[1]
						}
					case 3:
						b = td.tblBrc[5]
						if cr.last {
							b = td.tblBrc[3]
						}
					}
				}
				css = append(css, "border-"+names[k]+":"+brcCSS(b))
			}
			pad := td.gapHalf
			if pad < 0 {
				pad = 0
			}
			css = append(css, "padding:0 "+pt(pad))
			if cr.t.shade != "" {
				css = append(css, "background-color:"+cr.t.shade)
			}
			css = append(css, "vertical-align:"+[]string{"top", "middle", "bottom", "top"}[cr.t.vAlign&3])
			h := cr.html
			if strings.TrimSpace(h) == "" {
				h = "<p><br></p>"
			}
			sb.WriteString(`<td` + attrs + ` style="` + strings.Join(css, ";") + `">` + h + `</td>`)
		}
		sb.WriteString("</tr>")
	}
	sb.WriteString("</tbody></table>")
	return sb.String()
}

func brcCSS(b brc) string {
	if !b.set || b.none {
		return "none"
	}
	w := b.width
	if w < 0.5 {
		w = 0.5
	}
	return fnum(w) + "pt solid " + b.color
}

// ---- summary information ----

func summaryTitle(b []byte) string {
	if len(b) < 48 {
		return ""
	}
	if u32(b, 24) < 1 {
		return ""
	}
	so := int(u32(b, 44))
	if so <= 0 || so+8 > len(b) {
		return ""
	}
	n := int(u32(b, so+4))
	codepage := 1252
	title := []byte(nil)
	for i := 0; i < n && so+8+8*i+8 <= len(b); i++ {
		id := u32(b, so+8+8*i)
		off := so + int(u32(b, so+12+8*i))
		typ := u32(b, off)
		switch {
		case id == 1 && typ == 2:
			codepage = int(u16(b, off+4))
		case id == 2 && typ == 30:
			cch := int(u32(b, off+4))
			title = sliceN(b, off+8, cch)
		case id == 2 && typ == 31:
			cch := int(u32(b, off+4))
			s, _ := utf16z(b, off+8, cch)
			return strings.TrimSpace(s)
		}
	}
	if title == nil {
		return ""
	}
	title = bytes.TrimRight(title, "\x00")
	if codepage == 65001 {
		return strings.TrimSpace(string(title))
	}
	var sb strings.Builder
	for _, ch := range title {
		if ch >= 0x80 && ch < 0xA0 {
			sb.WriteRune(rune(cp1252[ch-0x80]))
		} else {
			sb.WriteRune(rune(ch))
		}
	}
	return strings.TrimSpace(sb.String())
}
