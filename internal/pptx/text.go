package pptx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// rProps is a layered set of character properties (nil = not specified).
type rProps struct {
	Size      *float64
	B, I      *bool
	U, Strike *bool
	Baseline  *int
	Color     *string
	Font      *string
	Highlight *string
}

func (r *rProps) merge(o *rProps) {
	if o == nil {
		return
	}
	if o.Size != nil {
		r.Size = o.Size
	}
	if o.B != nil {
		r.B = o.B
	}
	if o.I != nil {
		r.I = o.I
	}
	if o.U != nil {
		r.U = o.U
	}
	if o.Strike != nil {
		r.Strike = o.Strike
	}
	if o.Baseline != nil {
		r.Baseline = o.Baseline
	}
	if o.Color != nil {
		r.Color = o.Color
	}
	if o.Font != nil {
		r.Font = o.Font
	}
	if o.Highlight != nil {
		r.Highlight = o.Highlight
	}
}

func boolp(b bool) *bool      { return &b }
func f64p(f float64) *float64 { return &f }
func intp(i int) *int         { return &i }
func i64p(i int64) *int64     { return &i }

func parseBool(s string) bool { return s == "1" || s == "true" || s == "on" }

func resolveFont(t *themeInfo, f string) string {
	switch {
	case strings.HasPrefix(f, "+mj"):
		return t.major
	case strings.HasPrefix(f, "+mn"):
		return t.minor
	}
	return f
}

// apply overlays an a:rPr / a:defRPr / a:endParaRPr element.
func (r *rProps) apply(n *Node, cc *colorCtx) {
	if n == nil {
		return
	}
	if v, ok := n.AttrOK("sz"); ok {
		r.Size = f64p(atof(v) / 100)
	}
	if v, ok := n.AttrOK("b"); ok {
		r.B = boolp(parseBool(v))
	}
	if v, ok := n.AttrOK("i"); ok {
		r.I = boolp(parseBool(v))
	}
	if v, ok := n.AttrOK("u"); ok {
		r.U = boolp(v != "none" && v != "")
	}
	if v, ok := n.AttrOK("strike"); ok {
		r.Strike = boolp(v != "noStrike" && v != "")
	}
	if v, ok := n.AttrOK("baseline"); ok {
		r.Baseline = intp(int(math.Round(atof(v) / 1000)))
	}
	if c, kind, _ := cc.fillOf(n, ""); kind == "solid" {
		r.Color = strp(c)
	}
	if l := n.C("a:latin"); l != nil && l.A("typeface") != "" {
		r.Font = strp(resolveFont(cc.theme, l.A("typeface")))
	}
	if h := n.C("a:highlight"); h != nil {
		r.Highlight = strp(cc.resolveColor(firstColor(h), ""))
	}
}

type spacing struct {
	pts   float64
	pct   float64
	isPct bool
}

// pProps is a layered set of paragraph properties.
type pProps struct {
	Algn    *string
	MarL    *int64
	Indent  *int64
	LnPct   *float64
	LnPts   *float64
	Bef     *spacing
	Aft     *spacing
	BuType  *string // none | char | auto
	BuChar  *string
	BuAuto  *string
	BuStart *int
	BuFont  *string // "" = follow text
	BuColor *string // "" = follow text
	BuSize  *float64
	Def     rProps
}

func readSpacing(n *Node) *spacing {
	if n == nil {
		return nil
	}
	if p := n.C("a:spcPct"); p != nil {
		return &spacing{pct: atof(p.A("val")) / 1000, isPct: true}
	}
	if p := n.C("a:spcPts"); p != nil {
		return &spacing{pts: atof(p.A("val")) / 100}
	}
	return nil
}

func (p *pProps) apply(n *Node, cc *colorCtx) {
	if n == nil {
		return
	}
	if v, ok := n.AttrOK("algn"); ok {
		p.Algn = strp(v)
	}
	if v, ok := n.AttrOK("marL"); ok {
		p.MarL = i64p(atoi64(v))
	}
	if v, ok := n.AttrOK("indent"); ok {
		p.Indent = i64p(atoi64(v))
	}
	for _, c := range n.Elems() {
		switch c.Local {
		case "lnSpc":
			if s := readSpacing(c); s != nil {
				if s.isPct {
					p.LnPct, p.LnPts = f64p(s.pct), nil
				} else {
					p.LnPts, p.LnPct = f64p(s.pts), nil
				}
			}
		case "spcBef":
			if s := readSpacing(c); s != nil {
				p.Bef = s
			}
		case "spcAft":
			if s := readSpacing(c); s != nil {
				p.Aft = s
			}
		case "buClrTx":
			p.BuColor = strp("")
		case "buClr":
			p.BuColor = strp(cc.resolveColor(firstColor(c), ""))
		case "buSzTx":
			p.BuSize = nil
		case "buSzPct":
			p.BuSize = f64p(atof(c.A("val")) / 1000)
		case "buFontTx":
			p.BuFont = strp("")
		case "buFont":
			p.BuFont = strp(resolveFont(cc.theme, c.A("typeface")))
		case "buNone":
			p.BuType = strp("none")
		case "buAutoNum":
			p.BuType = strp("auto")
			p.BuAuto = strp(c.A("type"))
			if s := c.A("startAt"); s != "" {
				p.BuStart = intp(int(atoi64(s)))
			} else {
				p.BuStart = nil
			}
		case "buChar":
			p.BuType = strp("char")
			p.BuChar = strp(c.A("char"))
		case "buBlip":
			p.BuType = strp("char")
			p.BuChar = strp("•")
		case "defRPr":
			p.Def.apply(c, cc)
		}
	}
}

// layer is one level of text style inheritance.
type layer struct {
	list *Node   // a:lstStyle, p:titleStyle, p:defaultTextStyle ...
	r    *rProps // character overrides (fontRef color, table style text)
}

type textChain struct {
	cc     *colorCtx
	layers []layer
	cache  map[int]*pProps
}

func newChain(cc *colorCtx, layers ...layer) *textChain {
	base := &rProps{Size: f64p(18), Color: strp(cc.scheme("tx1")), Font: strp(cc.theme.minor),
		B: boolp(false), I: boolp(false), U: boolp(false), Strike: boolp(false), Baseline: intp(0)}
	ls := append([]layer{{r: base}}, layers...)
	return &textChain{cc: cc, layers: ls, cache: map[int]*pProps{}}
}

func (tc *textChain) with(l ...layer) *textChain {
	return &textChain{cc: tc.cc, layers: append(append([]layer{}, tc.layers...), l...), cache: map[int]*pProps{}}
}

// eff returns the inherited paragraph properties for a level (0-based).
func (tc *textChain) eff(level int) pProps {
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	if p, ok := tc.cache[level]; ok {
		return copyP(*p)
	}
	p := &pProps{}
	for _, l := range tc.layers {
		if l.list != nil {
			p.apply(l.list.C("a:defPPr"), tc.cc)
			p.apply(l.list.C("a:lvl"+strconv.Itoa(level+1)+"pPr"), tc.cc)
		}
		if l.r != nil {
			p.Def.merge(l.r)
		}
	}
	tc.cache[level] = p
	return copyP(*p)
}

func copyP(p pProps) pProps { return p }

// --- model conversion -------------------------------------------------

func runFromProps(r rProps) Run {
	out := Run{}
	if r.Size != nil {
		out.Size = round2(*r.Size)
	}
	if r.B != nil {
		out.B = *r.B
	}
	if r.I != nil {
		out.I = *r.I
	}
	if r.U != nil {
		out.U = *r.U
	}
	if r.Strike != nil {
		out.Strike = *r.Strike
	}
	if r.Baseline != nil {
		out.Baseline = *r.Baseline
	}
	if r.Color != nil {
		out.Color = *r.Color
	}
	if r.Font != nil {
		out.Font = *r.Font
	}
	if r.Highlight != nil {
		h := *r.Highlight
		out.Highlight = &h
	}
	return out
}

func round2(f float64) float64 { return math.Round(f*100) / 100 }

func spacingPts(s *spacing, size float64) float64 {
	if s == nil {
		return 0
	}
	if s.isPct {
		return round2(s.pct / 100 * size * 1.2)
	}
	return round2(s.pts)
}

func algnModel(a *string) string {
	if a == nil {
		return "l"
	}
	switch *a {
	case "ctr", "r", "just":
		return *a
	case "dist", "justLow", "thaiDist":
		return "just"
	}
	return "l"
}

func bulletModel(p pProps) *Bullet {
	if p.BuType == nil {
		return nil
	}
	switch *p.BuType {
	case "char":
		b := &Bullet{Char: "•"}
		if p.BuChar != nil && *p.BuChar != "" {
			b.Char = *p.BuChar
		}
		if p.BuFont != nil && *p.BuFont != "" {
			b.Font = *p.BuFont
		}
		if p.BuColor != nil && *p.BuColor != "" {
			b.Color = *p.BuColor
		}
		if p.BuSize != nil {
			b.Size = round2(*p.BuSize)
		}
		return b
	case "auto":
		b := &Bullet{Auto: "arabicPeriod"}
		if p.BuAuto != nil && *p.BuAuto != "" {
			b.Auto = *p.BuAuto
		}
		if p.BuStart != nil {
			b.StartAt = *p.BuStart
		}
		if p.BuColor != nil && *p.BuColor != "" {
			b.Color = *p.BuColor
		}
		if p.BuSize != nil {
			b.Size = round2(*p.BuSize)
		}
		return b
	}
	return nil
}

// paraShell fills the paragraph-level model fields from effective props.
func paraShell(p pProps, level int, sizeBasis float64) *Para {
	out := &Para{Align: algnModel(p.Algn), Level: level, Bullet: bulletModel(p), LineSpacing: 100, Runs: []*Run{}}
	if p.MarL != nil {
		out.MarL = *p.MarL
	}
	if p.Indent != nil {
		out.Indent = *p.Indent
	}
	if p.LnPct != nil {
		out.LineSpacing = round2(*p.LnPct)
	}
	if p.LnPts != nil {
		out.LineSpacingPts = round2(*p.LnPts)
	}
	out.SpaceBefore = spacingPts(p.Bef, sizeBasis)
	out.SpaceAfter = spacingPts(p.Aft, sizeBasis)
	return out
}

func sizeBasisFromRuns(runs []*Run, end *Run, def float64) float64 {
	for _, r := range runs {
		if r.Size > 0 {
			return r.Size
		}
	}
	if end != nil && end.Size > 0 {
		return end.Size
	}
	return def
}

// readParas converts a txBody into model paragraphs.
func readParas(txBody *Node, tc *textChain, rels map[string]Rel) []*Para {
	if txBody == nil {
		return nil
	}
	if ls := txBody.C("a:lstStyle"); ls != nil && len(ls.Elems()) > 0 {
		tc = tc.with(layer{list: ls})
	}
	var out []*Para
	for _, p := range txBody.All("a:p") {
		pPr := p.C("a:pPr")
		level := int(atoi64(pPr.A("lvl")))
		ep := tc.eff(level)
		inheritedDef := ep.Def
		ep.apply(pPr, tc.cc)
		ep.Def = inheritedDef // a paragraph's own defRPr does not style its runs
		var runs []*Run
		addRun := func(rPr *Node, text, field string) {
			rp := ep.Def
			rp.apply(rPr, tc.cc)
			r := runFromProps(rp)
			r.Text = text
			r.Field = field
			if h := rPr.C("a:hlinkClick"); h != nil {
				if rel, ok := rels[h.A("r:id")]; ok && rel.External {
					l := rel.Target
					r.Link = &l
				}
			}
			runs = append(runs, &r)
		}
		for _, c := range p.Elems() {
			switch {
			case c.Is("a:r"):
				addRun(c.C("a:rPr"), c.C("a:t").TextContent(), "")
			case c.Is("a:br"):
				addRun(c.C("a:rPr"), "\n", "")
			case c.Is("a:fld"):
				addRun(c.C("a:rPr"), c.C("a:t").TextContent(), c.A("type"))
			case c.Is("mc:AlternateContent"):
				// e.g. inline math: use fallback text if any.
				if fb := c.C("mc:Fallback"); fb != nil {
					for _, r := range fb.All("a:r") {
						addRun(r.C("a:rPr"), r.C("a:t").TextContent(), "")
					}
				}
			}
		}
		rp := ep.Def
		rp.apply(p.C("a:endParaRPr"), tc.cc)
		endRun := runFromProps(rp)
		end := &endRun
		para := paraShell(ep, level, sizeBasisFromRuns(runs, end, valOr(ep.Def.Size, 18)))
		if runs != nil {
			para.Runs = runs
		}
		para.End = end
		out = append(out, para)
	}
	return out
}

func valOr(f *float64, d float64) float64 {
	if f == nil {
		return d
	}
	return *f
}

// --- writing ------------------------------------------------------------

// relAlloc allocates relationships for the part being written.
type relAlloc interface {
	hyperlink(url string) string
}

func writeParas(sb *strings.Builder, paras []*Para, tc *textChain, ra relAlloc) {
	if len(paras) == 0 {
		sb.WriteString(`<a:p><a:endParaRPr lang="en-US"/></a:p>`)
		return
	}
	for _, p := range paras {
		writePara(sb, p, tc, ra)
	}
}

func writePara(sb *strings.Builder, p *Para, tc *textChain, ra relAlloc) {
	if p == nil {
		sb.WriteString(`<a:p/>`)
		return
	}
	level := p.Level
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	ep := tc.eff(level)
	inh := paraShell(ep, level, sizeBasisFromRuns(p.Runs, p.End, valOr(ep.Def.Size, 18)))
	var attrs, kids strings.Builder
	if level > 0 {
		fmt.Fprintf(&attrs, ` lvl="%d"`, level)
	}
	if p.MarL != inh.MarL {
		fmt.Fprintf(&attrs, ` marL="%d"`, p.MarL)
	}
	if p.Indent != inh.Indent {
		fmt.Fprintf(&attrs, ` indent="%d"`, p.Indent)
	}
	// An empty align means "not set": inherit from the layout/master.
	align := p.Align
	if align != "" && align != inh.Align {
		fmt.Fprintf(&attrs, ` algn="%s"`, esc(align))
	}
	if p.LineSpacingPts > 0 {
		if p.LineSpacingPts != inh.LineSpacingPts {
			fmt.Fprintf(&kids, `<a:lnSpc><a:spcPts val="%d"/></a:lnSpc>`, int(math.Round(p.LineSpacingPts*100)))
		}
	} else if p.LineSpacing > 0 && (p.LineSpacing != inh.LineSpacing || inh.LineSpacingPts > 0) {
		fmt.Fprintf(&kids, `<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>`, int(math.Round(p.LineSpacing*1000)))
	}
	if p.SpaceBefore != inh.SpaceBefore {
		fmt.Fprintf(&kids, `<a:spcBef><a:spcPts val="%d"/></a:spcBef>`, int(math.Round(p.SpaceBefore*100)))
	}
	if p.SpaceAfter != inh.SpaceAfter {
		fmt.Fprintf(&kids, `<a:spcAft><a:spcPts val="%d"/></a:spcAft>`, int(math.Round(p.SpaceAfter*100)))
	}
	if !bulletEq(p.Bullet, inh.Bullet) {
		b := p.Bullet
		switch {
		case b == nil:
			kids.WriteString(`<a:buNone/>`)
		default:
			ib := inh.Bullet
			if ib == nil {
				ib = &Bullet{}
			}
			if b.Color != "" {
				fmt.Fprintf(&kids, `<a:buClr><a:srgbClr val="%s"/></a:buClr>`, hexVal(b.Color))
			} else if ib.Color != "" {
				kids.WriteString(`<a:buClrTx/>`)
			}
			if b.Size > 0 {
				fmt.Fprintf(&kids, `<a:buSzPct val="%d"/>`, int(math.Round(b.Size*1000)))
			} else if ib.Size > 0 {
				kids.WriteString(`<a:buSzTx/>`)
			}
			if b.Auto != "" {
				start := ""
				if b.StartAt > 0 {
					start = fmt.Sprintf(` startAt="%d"`, b.StartAt)
				}
				if b.Font != "" {
					fmt.Fprintf(&kids, `<a:buFont typeface="%s"/>`, esc(b.Font))
				} else if ib.Font != "" {
					kids.WriteString(`<a:buFontTx/>`)
				}
				fmt.Fprintf(&kids, `<a:buAutoNum type="%s"%s/>`, esc(b.Auto), start)
			} else {
				ch := b.Char
				if ch == "" {
					ch = "•"
				}
				if b.Font != "" {
					fmt.Fprintf(&kids, `<a:buFont typeface="%s"/>`, esc(b.Font))
				} else if ib.Font != "" {
					kids.WriteString(`<a:buFontTx/>`)
				}
				fmt.Fprintf(&kids, `<a:buChar char="%s"/>`, esc(ch))
			}
		}
	}
	sb.WriteString(`<a:p>`)
	if attrs.Len() > 0 || kids.Len() > 0 {
		sb.WriteString(`<a:pPr` + attrs.String())
		if kids.Len() > 0 {
			sb.WriteString(">" + kids.String() + "</a:pPr>")
		} else {
			sb.WriteString("/>")
		}
	}
	def := runFromProps(ep.Def)
	for _, r := range p.Runs {
		if r == nil || r.Text == "" {
			continue
		}
		if r.Field != "" && !strings.Contains(r.Text, "\n") {
			fmt.Fprintf(sb, `<a:fld id="%s" type="%s">`, fieldGUID(r.Field, r.Text), esc(r.Field))
			writeRPr(sb, "a:rPr", r, &def, ra, true)
			sb.WriteString(`<a:t>` + escText(r.Text) + `</a:t></a:fld>`)
			continue
		}
		parts := strings.Split(strings.ReplaceAll(r.Text, "\r\n", "\n"), "\n")
		for i, t := range parts {
			if i > 0 {
				sb.WriteString(`<a:br>`)
				writeRPr(sb, "a:rPr", r, &def, ra, true)
				sb.WriteString(`</a:br>`)
			}
			if t == "" {
				continue
			}
			sb.WriteString(`<a:r>`)
			writeRPr(sb, "a:rPr", r, &def, ra, true)
			sb.WriteString(`<a:t>` + escText(t) + `</a:t></a:r>`)
		}
	}
	end := p.End
	if end == nil {
		if len(p.Runs) > 0 {
			end = p.Runs[len(p.Runs)-1]
		} else {
			end = &def
		}
	}
	e := *end
	e.Link = nil
	writeRPr(sb, "a:endParaRPr", &e, &def, ra, true)
	sb.WriteString(`</a:p>`)
}

func bulletEq(a, b *Bullet) bool {
	if a == nil || b == nil {
		return a == nil && b == nil
	}
	return *a == *b
}

func hexVal(c string) string {
	c = strings.ToUpper(strings.TrimPrefix(c, "#"))
	if len(c) == 3 {
		c = string([]byte{c[0], c[0], c[1], c[1], c[2], c[2]})
	}
	if len(c) != 6 {
		return "000000"
	}
	if _, err := strconv.ParseUint(c, 16, 32); err != nil {
		return "000000"
	}
	return c
}

// writeRPr writes only the properties of r that differ from def.
func writeRPr(sb *strings.Builder, tag string, r *Run, def *Run, ra relAlloc, always bool) {
	var attrs, kids strings.Builder
	attrs.WriteString(` lang="en-US"`)
	if r.Size > 0 && r.Size != def.Size {
		fmt.Fprintf(&attrs, ` sz="%d"`, int(math.Round(r.Size*100)))
	}
	if r.B != def.B {
		fmt.Fprintf(&attrs, ` b="%s"`, b01(r.B))
	}
	if r.I != def.I {
		fmt.Fprintf(&attrs, ` i="%s"`, b01(r.I))
	}
	if r.U != def.U {
		if r.U {
			attrs.WriteString(` u="sng"`)
		} else {
			attrs.WriteString(` u="none"`)
		}
	}
	if r.Strike != def.Strike {
		if r.Strike {
			attrs.WriteString(` strike="sngStrike"`)
		} else {
			attrs.WriteString(` strike="noStrike"`)
		}
	}
	if r.Baseline != def.Baseline {
		fmt.Fprintf(&attrs, ` baseline="%d"`, r.Baseline*1000)
	}
	attrs.WriteString(` dirty="0"`)
	if r.Color != "" && !strings.EqualFold(r.Color, def.Color) {
		fmt.Fprintf(&kids, `<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, hexVal(r.Color))
	}
	if r.Highlight != nil && *r.Highlight != "" && (def.Highlight == nil || !strings.EqualFold(*def.Highlight, *r.Highlight)) {
		fmt.Fprintf(&kids, `<a:highlight><a:srgbClr val="%s"/></a:highlight>`, hexVal(*r.Highlight))
	}
	if r.Font != "" && r.Font != def.Font {
		fmt.Fprintf(&kids, `<a:latin typeface="%s"/><a:cs typeface="%s"/>`, esc(r.Font), esc(r.Font))
	}
	if r.Link != nil && *r.Link != "" && ra != nil {
		fmt.Fprintf(&kids, `<a:hlinkClick r:id="%s"/>`, ra.hyperlink(*r.Link))
	}
	sb.WriteString("<" + tag + attrs.String())
	if kids.Len() > 0 {
		sb.WriteString(">" + kids.String() + "</" + tag + ">")
	} else {
		sb.WriteString("/>")
	}
}

func b01(b bool) string {
	if b {
		return "1"
	}
	return "0"
}

func fieldGUID(field, text string) string {
	h := uint64(1469598103934665603)
	for _, c := range field + "|" + text {
		h ^= uint64(c)
		h *= 1099511628211
	}
	s := fmt.Sprintf("%016X%016X", h, h*2654435761)
	return "{" + s[0:8] + "-" + s[8:12] + "-" + s[12:16] + "-" + s[16:20] + "-" + s[20:32] + "}"
}
