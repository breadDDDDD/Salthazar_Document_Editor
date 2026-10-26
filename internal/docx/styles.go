package docx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// This file resolves WordprocessingML formatting inheritance:
// docDefaults -> table style -> paragraph style chain (basedOn) -> numbering
// level -> direct paragraph properties, and for runs additionally the
// character style chain and direct run properties. Theme fonts and colours
// are resolved through theme1.xml.

// tri is a tri-state toggle: unset, on or off.
type tri uint8

const (
	unset tri = iota
	on
	off
)

func (t tri) is() bool { return t == on }

type optInt struct {
	v  int
	ok bool
}

func (o *optInt) set(v int) { o.v, o.ok = v, true }

// rpr holds (possibly partial) run properties.
type rpr struct {
	ascii, hAnsi, eastAsia, cs string // resolved font names
	sz                         int    // half-points, 0 = unset
	b, i, caps, smallCaps      tri
	strike, vanish             tri
	u                          string // "" unset, "none", or a Word underline style
	color                      string // "" unset, "auto", or lowercase rrggbb
	bg                         string // "" unset, "none", or lowercase rrggbb (highlight / shading)
	vert                       string // "" unset, "baseline", "superscript", "subscript"
	spacing                    optInt // character spacing, twips
}

func (a rpr) merge(o rpr) rpr {
	if o.ascii != "" {
		a.ascii = o.ascii
	}
	if o.hAnsi != "" {
		a.hAnsi = o.hAnsi
	}
	if o.eastAsia != "" {
		a.eastAsia = o.eastAsia
	}
	if o.cs != "" {
		a.cs = o.cs
	}
	if o.sz > 0 {
		a.sz = o.sz
	}
	for _, p := range [][2]*tri{{&a.b, &o.b}, {&a.i, &o.i}, {&a.caps, &o.caps}, {&a.smallCaps, &o.smallCaps},
		{&a.strike, &o.strike}, {&a.vanish, &o.vanish}} {
		if *p[1] != unset {
			*p[0] = *p[1]
		}
	}
	if o.u != "" {
		a.u = o.u
	}
	if o.color != "" {
		a.color = o.color
	}
	if o.bg != "" {
		a.bg = o.bg
	}
	if o.vert != "" {
		a.vert = o.vert
	}
	if o.spacing.ok {
		a.spacing = o.spacing
	}
	return a
}

// font returns the primary (Latin) font name.
func (a rpr) font() string {
	if a.ascii != "" {
		return a.ascii
	}
	return a.hAnsi
}

// ppr holds (possibly partial) paragraph properties.
type ppr struct {
	jc                          string
	left, right, first, hanging optInt // twips
	before, after, line         optInt
	lineRule                    string
	beforeAuto, afterAuto       tri
	pageBreakBefore, contextual tri
	numID, ilvl                 optInt
	shd                         string // lowercase rrggbb or "none"
	rightTab                    tri    // has a right-aligned tab stop
	outline                     optInt
	borderBottom                string // CSS border or "" (paragraph bottom border)
	borderTop                   string
}

func (a ppr) merge(o ppr) ppr {
	if o.jc != "" {
		a.jc = o.jc
	}
	for _, p := range [][2]*optInt{{&a.left, &o.left}, {&a.right, &o.right}, {&a.before, &o.before},
		{&a.after, &o.after}, {&a.line, &o.line}, {&a.numID, &o.numID}, {&a.ilvl, &o.ilvl}, {&a.outline, &o.outline}} {
		if p[1].ok {
			*p[0] = *p[1]
		}
	}
	// firstLine and hanging are mutually exclusive: setting one clears the other.
	if o.first.ok {
		a.first, a.hanging = o.first, optInt{}
	}
	if o.hanging.ok {
		a.hanging, a.first = o.hanging, optInt{}
	}
	if o.lineRule != "" {
		a.lineRule = o.lineRule
	}
	for _, p := range [][2]*tri{{&a.beforeAuto, &o.beforeAuto}, {&a.afterAuto, &o.afterAuto},
		{&a.pageBreakBefore, &o.pageBreakBefore}, {&a.contextual, &o.contextual}, {&a.rightTab, &o.rightTab}} {
		if *p[1] != unset {
			*p[0] = *p[1]
		}
	}
	if o.shd != "" {
		a.shd = o.shd
	}
	if o.borderBottom != "" {
		a.borderBottom = o.borderBottom
	}
	if o.borderTop != "" {
		a.borderTop = o.borderTop
	}
	return a
}

type styleDef struct {
	typ, id, name, basedOn string
	p                      ppr
	r                      rpr
	tblPr, tcPr, node      *xnode
	// resolved (inherited) values
	done  bool
	effP  ppr
	effR  rpr
	chain []*styleDef // self first, then ancestors
}

type themeInfo struct {
	fonts  map[string]string // majorHAnsi, minorHAnsi, majorEastAsia, ... -> typeface
	colors map[string]string // dk1, lt1, accent1 ... -> rrggbb
}

func (r *reader) loadTheme(root *xnode) {
	r.theme = themeInfo{fonts: map[string]string{}, colors: map[string]string{}}
	if root == nil {
		return
	}
	walkAll(root, "fontScheme", func(fs *xnode) {
		for _, kind := range []string{"major", "minor"} {
			f := fs.child(kind + "Font")
			if f == nil {
				continue
			}
			if l := f.child("latin"); l != nil {
				r.theme.fonts[kind+"HAnsi"] = l.attr("typeface")
			}
			if e := f.child("ea"); e != nil && e.attr("typeface") != "" {
				r.theme.fonts[kind+"EastAsia"] = e.attr("typeface")
			}
			if c := f.child("cs"); c != nil && c.attr("typeface") != "" {
				r.theme.fonts[kind+"Bidi"] = c.attr("typeface")
			}
		}
	})
	walkAll(root, "clrScheme", func(cs *xnode) {
		for _, c := range cs.kids {
			for _, k := range c.kids {
				v := ""
				switch k.name {
				case "srgbClr":
					v = k.attr("val")
				case "sysClr":
					v = k.attr("lastClr")
					if v == "" {
						v = map[string]string{"windowText": "000000", "window": "FFFFFF"}[k.attr("val")]
					}
				}
				if hexRe.MatchString(v) {
					r.theme.colors[c.name] = strings.ToLower(v)
				}
			}
		}
	})
}

func (r *reader) themeFont(name string) string {
	if f := r.theme.fonts[name]; f != "" {
		return f
	}
	switch {
	case strings.HasPrefix(name, "major"):
		if f := r.theme.fonts["majorHAnsi"]; f != "" {
			return f
		}
		return "Calibri Light"
	case strings.HasPrefix(name, "minor"):
		if f := r.theme.fonts["minorHAnsi"]; f != "" {
			return f
		}
		return "Calibri"
	}
	return ""
}

var themeColorAlias = map[string]string{"text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2",
	"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2", "hyperlink": "hlink", "followedHyperlink": "folHlink"}

// themeColor resolves a w:themeColor with optional themeTint/themeShade (hex bytes).
func (r *reader) themeColor(name, tint, shade string) string {
	if a, ok := themeColorAlias[name]; ok {
		name = a
	}
	c := r.theme.colors[name]
	if c == "" {
		return ""
	}
	rgb, _ := strconv.ParseUint(c, 16, 32)
	ch := [3]float64{float64(rgb >> 16 & 255), float64(rgb >> 8 & 255), float64(rgb & 255)}
	if t, err := strconv.ParseUint(tint, 16, 8); err == nil && tint != "" {
		f := float64(t) / 255
		for i := range ch {
			ch[i] = ch[i]*f + 255*(1-f)
		}
	}
	if s, err := strconv.ParseUint(shade, 16, 8); err == nil && shade != "" {
		f := float64(s) / 255
		for i := range ch {
			ch[i] *= f
		}
	}
	return fmt.Sprintf("%02x%02x%02x", int(math.Round(ch[0])), int(math.Round(ch[1])), int(math.Round(ch[2])))
}

// colorOf reads a colour from an element with val/color + themeColor attributes.
func (r *reader) colorOf(n *xnode, valAttr string) string {
	if n == nil {
		return ""
	}
	v := n.attr(valAttr)
	if hexRe.MatchString(v) {
		return strings.ToLower(v)
	}
	if tc := n.attr("themeColor"); tc != "" {
		if c := r.themeColor(tc, n.attr("themeTint"), n.attr("themeShade")); c != "" {
			return c
		}
	}
	if v == "auto" {
		return "auto"
	}
	return ""
}

func parseTri(n *xnode, name string) tri {
	c := n.child(name)
	if c == nil {
		return unset
	}
	switch c.attr("val") {
	case "0", "false", "off", "none":
		return off
	}
	return on
}

func atoiOpt(s string) (optInt, bool) {
	if s == "" {
		return optInt{}, false
	}
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return optInt{}, false
	}
	return optInt{v: int(math.Round(f)), ok: true}, true
}

func (r *reader) parseRPr(n *xnode) rpr {
	var o rpr
	if n == nil {
		return o
	}
	if rf := n.child("rFonts"); rf != nil {
		pick := func(theme, plain string) string {
			if t := rf.attr(theme); t != "" {
				if f := r.themeFont(t); f != "" {
					return f
				}
			}
			return rf.attr(plain)
		}
		o.ascii = pick("asciiTheme", "ascii")
		o.hAnsi = pick("hAnsiTheme", "hAnsi")
		o.eastAsia = pick("eastAsiaTheme", "eastAsia")
		o.cs = pick("cstheme", "cs")
	}
	if v, ok := n.val("sz"); ok {
		if f, err := strconv.ParseFloat(v, 64); err == nil && f > 0 {
			o.sz = int(math.Round(f))
		}
	}
	o.b = parseTri(n, "b")
	o.i = parseTri(n, "i")
	o.caps = parseTri(n, "caps")
	o.smallCaps = parseTri(n, "smallCaps")
	o.vanish = parseTri(n, "vanish")
	if s := parseTri(n, "strike"); s != unset {
		o.strike = s
	}
	if s := parseTri(n, "dstrike"); s == on {
		o.strike = on
	}
	if u := n.child("u"); u != nil {
		v := u.attr("val")
		if v == "" {
			v = "single"
		}
		if v == "0" || v == "false" {
			v = "none"
		}
		o.u = v
	}
	if c := n.child("color"); c != nil {
		o.color = r.colorOf(c, "val")
	}
	if v, ok := n.val("highlight"); ok {
		if c, ok := highlightCSS[v]; ok {
			o.bg = strings.TrimPrefix(c, "#")
		} else if v == "none" {
			o.bg = "none"
		}
	}
	if shd := n.child("shd"); shd != nil && o.bg == "" {
		if c := r.colorOf(shd, "fill"); c != "" && c != "auto" {
			o.bg = c
		}
	}
	if v, ok := n.val("vertAlign"); ok {
		o.vert = v
	}
	if v, ok := n.val("spacing"); ok {
		if x, ok := atoiOpt(v); ok {
			o.spacing = x
		}
	}
	return o
}

func (r *reader) borderCSS(b *xnode) string {
	if b == nil {
		return ""
	}
	v := b.attr("val")
	switch v {
	case "", "nil", "none":
		return "none"
	}
	sz, _ := strconv.ParseFloat(b.attr("sz"), 64) // eighths of a point
	w := sz / 8
	if w < 0.5 {
		w = 0.5
	}
	style := "solid"
	switch {
	case strings.HasPrefix(v, "double") || v == "triple":
		style = "double"
		if w < 2.25 {
			w = 2.25
		}
	case strings.HasPrefix(v, "dash"):
		style = "dashed"
	case strings.HasPrefix(v, "dot"):
		style = "dotted"
	case v == "inset" || v == "outset":
		style = v
	}
	c := r.colorOf(b, "color")
	if c == "" || c == "auto" {
		c = "000000"
	}
	return fmt.Sprintf("%gpt %s #%s", math.Round(w*100)/100, style, c)
}

func (r *reader) parsePPr(n *xnode) ppr {
	var o ppr
	if n == nil {
		return o
	}
	if v, ok := n.val("jc"); ok {
		o.jc = v
	}
	if ind := n.child("ind"); ind != nil {
		for _, p := range []struct {
			dst   *optInt
			attrs []string
		}{{&o.left, []string{"left", "start"}}, {&o.right, []string{"right", "end"}},
			{&o.first, []string{"firstLine"}}, {&o.hanging, []string{"hanging"}}} {
			for _, a := range p.attrs {
				if x, ok := atoiOpt(ind.attr(a)); ok {
					*p.dst = x
					break
				}
			}
		}
		if o.first.ok && o.hanging.ok {
			o.first = optInt{}
		}
	}
	if sp := n.child("spacing"); sp != nil {
		if x, ok := atoiOpt(sp.attr("before")); ok {
			o.before = x
		}
		if x, ok := atoiOpt(sp.attr("after")); ok {
			o.after = x
		}
		if x, ok := atoiOpt(sp.attr("line")); ok {
			o.line = x
			o.lineRule = sp.attr("lineRule")
			if o.lineRule == "" {
				o.lineRule = "auto"
			}
		}
		switch sp.attr("beforeAutospacing") {
		case "1", "true", "on":
			o.beforeAuto = on
		case "0", "false", "off":
			o.beforeAuto = off
		}
		switch sp.attr("afterAutospacing") {
		case "1", "true", "on":
			o.afterAuto = on
		case "0", "false", "off":
			o.afterAuto = off
		}
	}
	o.pageBreakBefore = parseTri(n, "pageBreakBefore")
	if tabs := n.child("tabs"); tabs != nil {
		o.rightTab = off
		for _, t := range tabs.kids {
			if v := t.attr("val"); t.name == "tab" && (v == "right" || v == "end") {
				o.rightTab = on
			}
		}
	}
	o.contextual = parseTri(n, "contextualSpacing")
	if np := n.child("numPr"); np != nil {
		if v, ok := np.val("numId"); ok {
			if x, ok := atoiOpt(v); ok {
				o.numID = x
			}
		}
		if v, ok := np.val("ilvl"); ok {
			if x, ok := atoiOpt(v); ok {
				o.ilvl = x
			}
		}
	}
	if shd := n.child("shd"); shd != nil {
		if c := r.colorOf(shd, "fill"); c != "" {
			if c == "auto" {
				c = "none"
			}
			o.shd = c
		}
	}
	if v, ok := n.val("outlineLvl"); ok {
		if x, ok := atoiOpt(v); ok {
			o.outline = x
		}
	}
	if bd := n.child("pBdr"); bd != nil {
		o.borderBottom = r.borderCSS(bd.child("bottom"))
		o.borderTop = r.borderCSS(bd.child("top"))
	}
	return o
}

func (r *reader) loadStyles(root *xnode) {
	r.styles = map[string]*styleDef{}
	if root == nil {
		return
	}
	if dd := root.child("styles").child("docDefaults"); dd != nil {
		r.defR = r.parseRPr(dd.child("rPrDefault").child("rPr"))
		r.defP = r.parsePPr(dd.child("pPrDefault").child("pPr"))
	}
	walkAll(root, "style", func(s *xnode) {
		d := &styleDef{typ: s.attr("type"), id: s.attr("styleId")}
		d.name, _ = s.val("name")
		d.basedOn, _ = s.val("basedOn")
		d.p = r.parsePPr(s.child("pPr"))
		d.r = r.parseRPr(s.child("rPr"))
		d.tblPr = s.child("tblPr")
		d.tcPr = s.child("tcPr")
		d.node = s
		r.styles[d.id] = d
		if def := s.attr("default"); def == "1" || def == "true" {
			switch d.typ {
			case "paragraph":
				r.defPara = d.id
			case "table":
				r.defTable = d.id
			}
		}
	})
}

// style returns the resolved style (with inherited properties) or nil.
func (r *reader) style(id string) *styleDef {
	d := r.styles[id]
	if d == nil {
		return nil
	}
	if d.done {
		return d
	}
	d.done = true // guards against basedOn cycles
	d.effP, d.effR, d.chain = d.p, d.r, []*styleDef{d}
	if parent := r.style(d.basedOn); parent != nil && parent != d {
		d.effP = parent.effP.merge(d.p)
		d.effR = parent.effR.merge(d.r)
		d.chain = append(d.chain, parent.chain...)
	}
	return d
}

// headingLevel returns 1..9 for heading styles (0 if not a heading) and
// whether the style is the Title style.
func (r *reader) headingLevel(id string) (int, bool) {
	d := r.style(id)
	if d == nil {
		return 0, false
	}
	for _, s := range d.chain {
		if m := headingRe.FindStringSubmatch(s.name); m != nil {
			n, _ := strconv.Atoi(m[1])
			return n, false
		}
		if m := headingRe.FindStringSubmatch(s.id); m != nil {
			n, _ := strconv.Atoi(m[1])
			return n, false
		}
		if strings.EqualFold(s.name, "title") || strings.EqualFold(s.id, "title") {
			return 1, true
		}
	}
	if d.effP.outline.ok && d.effP.outline.v < 9 {
		return d.effP.outline.v + 1, false
	}
	return 0, false
}

// ---- numbering ----

type numLevel struct {
	start   int
	fmt     string
	text    string
	isLgl   bool
	restart int // lvlRestart, -1 = default
	p       ppr
	r       rpr
}

type numDef struct {
	abs       string
	levels    [9]*numLevel
	overrides map[int]int // ilvl -> startOverride
}

func (r *reader) loadNumbering(root *xnode) {
	r.nums = map[int]*numDef{}
	if root == nil {
		return
	}
	parseLvl := func(l *xnode) *numLevel {
		nl := &numLevel{start: 1, fmt: "decimal", restart: -1}
		if v, ok := l.val("start"); ok {
			nl.start, _ = strconv.Atoi(v)
		}
		if v, ok := l.val("numFmt"); ok {
			nl.fmt = v
		}
		nl.text, _ = l.val("lvlText")
		nl.isLgl = l.child("isLgl") != nil && parseTri(l, "isLgl") == on
		if v, ok := l.val("lvlRestart"); ok {
			nl.restart, _ = strconv.Atoi(v)
		}
		nl.p = r.parsePPr(l.child("pPr"))
		nl.r = r.parseRPr(l.child("rPr"))
		return nl
	}
	abs := map[string][9]*numLevel{}
	walkAll(root, "abstractNum", func(a *xnode) {
		var lv [9]*numLevel
		for _, l := range a.kids {
			if l.name == "lvl" {
				il, _ := strconv.Atoi(l.attr("ilvl"))
				if il >= 0 && il < 9 {
					lv[il] = parseLvl(l)
				}
			}
		}
		abs[a.attr("abstractNumId")] = lv
	})
	walkAll(root, "num", func(n *xnode) {
		id, _ := strconv.Atoi(n.attr("numId"))
		a, _ := n.val("abstractNumId")
		nd := &numDef{abs: a, levels: abs[a], overrides: map[int]int{}}
		for _, o := range n.kids {
			if o.name != "lvlOverride" {
				continue
			}
			il, _ := strconv.Atoi(o.attr("ilvl"))
			if il < 0 || il > 8 {
				continue
			}
			if l := o.child("lvl"); l != nil {
				nd.levels[il] = parseLvl(l)
			}
			if v, ok := o.val("startOverride"); ok {
				nd.overrides[il], _ = strconv.Atoi(v)
			}
		}
		r.nums[id] = nd
	})
}

func (r *reader) numLevel(numID, ilvl int) *numLevel {
	nd := r.nums[numID]
	if nd == nil || ilvl < 0 || ilvl > 8 {
		return nil
	}
	return nd.levels[ilvl]
}

// numCounter tracks list counters per abstract numbering definition.
type numCounter struct {
	vals    [9]int
	started [9]bool
}

// nextLabel advances the counter for (numID, ilvl) and returns the label text.
func (r *reader) nextLabel(numID, ilvl int) (string, *numLevel, int) {
	nd := r.nums[numID]
	lvl := r.numLevel(numID, ilvl)
	if nd == nil || lvl == nil {
		return "", nil, 0
	}
	key := nd.abs
	c := r.counters[key]
	if c == nil {
		c = &numCounter{}
		r.counters[key] = c
	}
	if !r.numUsed[numID] {
		r.numUsed[numID] = true
		for il, s := range nd.overrides {
			c.vals[il], c.started[il] = s-1, true
		}
	}
	if !c.started[ilvl] {
		c.vals[ilvl], c.started[ilvl] = lvl.start-1, true
	}
	c.vals[ilvl]++
	for l := ilvl + 1; l < 9; l++ {
		if ll := nd.levels[l]; ll == nil || ll.restart != 0 {
			c.started[l] = false
		}
	}
	text := lvl.text
	if lvl.fmt == "bullet" {
		return bulletChar(text, lvl.r.font()), lvl, c.vals[ilvl]
	}
	if lvl.fmt == "none" {
		text = numPlaceholder.ReplaceAllString(text, "")
	}
	text = numPlaceholder.ReplaceAllStringFunc(text, func(m string) string {
		k := int(m[1] - '1')
		l := nd.levels[k]
		v := c.vals[k]
		if l == nil {
			return strconv.Itoa(v)
		}
		if !c.started[k] {
			v = l.start
		}
		f := l.fmt
		if lvl.isLgl && k != ilvl {
			f = "decimal"
		}
		return formatNum(v, f)
	})
	return text, lvl, c.vals[ilvl]
}

func formatNum(v int, f string) string {
	switch f {
	case "lowerLetter":
		return strings.ToLower(letters(v))
	case "upperLetter":
		return letters(v)
	case "lowerRoman":
		return strings.ToLower(roman(v))
	case "upperRoman":
		return roman(v)
	case "decimalZero":
		if v < 10 {
			return "0" + strconv.Itoa(v)
		}
	case "ordinal":
		suf := "th"
		if v%100 < 11 || v%100 > 13 {
			switch v % 10 {
			case 1:
				suf = "st"
			case 2:
				suf = "nd"
			case 3:
				suf = "rd"
			}
		}
		return strconv.Itoa(v) + suf
	case "none":
		return ""
	}
	return strconv.Itoa(v)
}

func letters(v int) string {
	if v <= 0 {
		return strconv.Itoa(v)
	}
	// Word repeats the letter: a..z, aa..zz, aaa..
	ch := byte('A' + (v-1)%26)
	return strings.Repeat(string(ch), (v-1)/26+1)
}

func roman(v int) string {
	if v <= 0 || v >= 4000 {
		return strconv.Itoa(v)
	}
	vals := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	syms := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
	var sb strings.Builder
	for i, n := range vals {
		for v >= n {
			sb.WriteString(syms[i])
			v -= n
		}
	}
	return sb.String()
}

// bulletChar maps symbol-font bullet characters to Unicode equivalents.
func bulletChar(text, font string) string {
	if text == "" {
		return "•"
	}
	rs := []rune(text)
	c := rs[0]
	if c >= 0xF000 && c <= 0xF0FF {
		c -= 0xF000
	}
	f := strings.ToLower(font)
	switch {
	case strings.Contains(f, "wingdings"):
		if m, ok := map[rune]string{0xA7: "▪", 0xD8: "➢", 0xFC: "✓", 'v': "❖", 'q': "❑", 'n': "■", 'l': "●",
			0xA8: "◻", 0x6F: "□", 0xE0: "➔", 0xE8: "➔", 0xF0: "⇨", 0x9F: "•"}[c]; ok {
			return m
		}
		return "•"
	case strings.Contains(f, "symbol"):
		if m, ok := map[rune]string{0xB7: "•", 0x2D: "–", 0xA8: "♦", 0xDE: "⇒", 0xAE: "→"}[c]; ok {
			return m
		}
		return "•"
	}
	if c == 0xB7 || c == 0xF0B7 {
		return "•"
	}
	return string(rs)
}
