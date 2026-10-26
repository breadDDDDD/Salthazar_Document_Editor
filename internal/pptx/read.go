package pptx

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"math"
	"path"
	"strconv"
	"strings"
)

type masterInfo struct {
	part   string
	root   *Node
	rels   map[string]Rel
	theme  *themeInfo
	themeP string
	cc     *colorCtx
}

type layoutInfo struct {
	part   string
	name   string
	root   *Node
	rels   map[string]Rel
	master *masterInfo
	cc     *colorCtx
}

// origShape remembers where a parsed slide shape came from so the writer can
// re-emit untouched shapes verbatim.
type origShape struct {
	node *Node
	rels map[string]Rel
	json string
}

type reader struct {
	pkg       *Pkg
	model     *Model
	presPart  string
	pres      *Node
	defText   *Node
	masters   map[string]*masterInfo
	layouts   map[string]*layoutInfo
	layoutOrd []*layoutInfo
	tblStyles map[string]*Node
	orig      map[string]map[string]*origShape // slide part -> shape id
	origSlide map[string]*slideOrig
}

type slideOrig struct {
	root      *Node
	rels      map[string]Rel
	notes     string
	notesPart string
}

// sctx is the context for parsing one shape tree.
type sctx struct {
	part   string
	rels   map[string]Rel
	cc     *colorCtx
	level  string // slide | layout | master
	layout *layoutInfo
	master *masterInfo
	idPfx  string
	decor  bool
}

// ToModel parses a .pptx into the editor model. The second result is the
// presentation title (from core properties or the first slide title).
func ToModel(b []byte) (*Model, string, error) {
	rd, err := newReader(b)
	if err != nil {
		return nil, "", err
	}
	if err := rd.read(); err != nil {
		return nil, "", err
	}
	return rd.model, rd.title(), nil
}

func newReader(b []byte) (*reader, error) {
	p, err := openPkg(b)
	if err != nil {
		return nil, err
	}
	return newReaderPkg(p)
}

func newReaderPkg(p *Pkg) (*reader, error) {
	rd := &reader{pkg: p, masters: map[string]*masterInfo{}, layouts: map[string]*layoutInfo{},
		tblStyles: map[string]*Node{}, orig: map[string]map[string]*origShape{}, origSlide: map[string]*slideOrig{}}
	rd.presPart = p.mainPart()
	rd.pres = p.XML(rd.presPart)
	if rd.pres == nil || rd.pres.Local != "presentation" {
		return nil, fmt.Errorf("not a PowerPoint presentation (missing %s)", rd.presPart)
	}
	rd.defText = rd.pres.C("p:defaultTextStyle")
	prels := p.Rels(rd.presPart)
	// masters & layouts
	for _, mid := range rd.pres.C("p:sldMasterIdLst").All("p:sldMasterId") {
		rel, ok := prels[mid.A("r:id")]
		if !ok {
			continue
		}
		rd.loadMaster(rel.Target)
	}
	if len(rd.masters) == 0 { // fall back to any master rel
		for _, id := range sortedRelIDs(prels) {
			if sameRelKind(prels[id].Type, rtSlideMaster) {
				rd.loadMaster(prels[id].Target)
			}
		}
	}
	if ts, ok := relByType(prels, rtTableStyles); ok {
		if root := p.XML(ts.Target); root != nil {
			for _, st := range root.All("a:tblStyle") {
				rd.tblStyles[strings.ToUpper(st.A("styleId"))] = st
			}
		}
	}
	return rd, nil
}

func (rd *reader) loadMaster(part string) *masterInfo {
	if m, ok := rd.masters[part]; ok {
		return m
	}
	root := rd.pkg.XML(part)
	if root == nil {
		return nil
	}
	m := &masterInfo{part: part, root: root, rels: rd.pkg.Rels(part)}
	if th, ok := relByType(m.rels, rtTheme); ok {
		m.themeP = th.Target
		m.theme = parseTheme(rd.pkg.XML(th.Target))
	} else {
		m.theme = parseTheme(nil)
	}
	m.cc = &colorCtx{theme: m.theme, clrMap: readClrMap(root.C("p:clrMap"), defaultClrMap())}
	rd.masters[part] = m
	for _, lid := range root.C("p:sldLayoutIdLst").All("p:sldLayoutId") {
		if rel, ok := m.rels[lid.A("r:id")]; ok {
			rd.loadLayout(rel.Target, m)
		}
	}
	return m
}

func (rd *reader) loadLayout(part string, m *masterInfo) *layoutInfo {
	if l, ok := rd.layouts[part]; ok {
		return l
	}
	root := rd.pkg.XML(part)
	if root == nil {
		return nil
	}
	l := &layoutInfo{part: part, root: root, rels: rd.pkg.Rels(part), master: m}
	l.name = root.C("p:cSld").A("name")
	if l.name == "" {
		l.name = root.A("type")
	}
	if m == nil {
		if mr, ok := relByType(l.rels, rtSlideMaster); ok {
			m = rd.loadMaster(mr.Target)
			l.master = m
		}
	}
	if m == nil {
		return nil
	}
	l.cc = &colorCtx{theme: m.theme, clrMap: overrideClrMap(root, m.cc.clrMap)}
	rd.layouts[part] = l
	rd.layoutOrd = append(rd.layoutOrd, l)
	return l
}

func overrideClrMap(root *Node, base map[string]string) map[string]string {
	if ov := root.Path("p:clrMapOvr", "a:overrideClrMapping"); ov != nil {
		return readClrMap(ov, base)
	}
	return base
}

func (rd *reader) firstTheme() *themeInfo {
	for _, l := range rd.layoutOrd {
		return l.master.theme
	}
	for _, m := range rd.masters {
		return m.theme
	}
	return parseTheme(nil)
}

func (rd *reader) read() error {
	m := &Model{Media: map[string]string{}}
	rd.model = m
	sz := rd.pres.C("p:sldSz")
	m.Size = Size{CX: atoi64(sz.A("cx")), CY: atoi64(sz.A("cy"))}
	if m.Size.CX <= 0 || m.Size.CY <= 0 {
		m.Size = Size{CX: 12192000, CY: 6858000}
	}
	th := rd.firstTheme()
	m.Theme = Theme{MajorFont: th.major, MinorFont: th.minor, Colors: map[string]string{}}
	for k, v := range th.colors {
		m.Theme.Colors[k] = v
	}
	m.Layouts = []*Layout{}
	for _, l := range rd.layoutOrd {
		m.Layouts = append(m.Layouts, rd.layoutModel(l))
	}
	m.Slides = []*Slide{}
	prels := rd.pkg.Rels(rd.presPart)
	n := 0
	for _, sid := range rd.pres.C("p:sldIdLst").All("p:sldId") {
		rel, ok := prels[sid.A("r:id")]
		if !ok {
			continue
		}
		n++
		s, err := rd.readSlide(rel.Target, n)
		if err != nil {
			return fmt.Errorf("slide %d: %v", n, err)
		}
		m.Slides = append(m.Slides, s)
	}
	return nil
}

func (rd *reader) title() string {
	for _, s := range rd.model.Slides {
		for _, sh := range s.Shapes {
			if sh.Ph != nil && (*sh.Ph == "title" || *sh.Ph == "ctrTitle") {
				if t := strings.Join(strings.Fields(shapeText(sh)), " "); t != "" {
					return t
				}
			}
		}
		break
	}
	if core := rd.pkg.XML("docProps/core.xml"); core != nil {
		for _, c := range core.Elems() {
			if c.Local == "title" {
				if t := strings.TrimSpace(c.TextContent()); t != "" {
					return t
				}
			}
		}
	}
	return ""
}

func shapeText(sh *Shape) string {
	var sb strings.Builder
	for i, p := range sh.Paras {
		if i > 0 {
			sb.WriteString("\n")
		}
		for _, r := range p.Runs {
			sb.WriteString(r.Text)
		}
	}
	return sb.String()
}

func (rd *reader) layoutModel(l *layoutInfo) *Layout {
	out := &Layout{ID: l.part, Name: l.name}
	out.Bg = rd.resolveBg(nil, nil, l)
	sc := &sctx{part: l.part, rels: l.rels, cc: l.cc, level: "layout", layout: l, master: l.master, idPfx: "lp"}
	for _, sh := range rd.parseTree(l.root.Path("p:cSld", "p:spTree"), sc, identXf) {
		if sh.Ph == nil {
			continue
		}
		// placeholders carry no prompt text
		for _, p := range sh.Paras {
			p.Runs = []*Run{}
		}
		if len(sh.Paras) > 1 {
			sh.Paras = sh.Paras[:1]
		}
		out.Placeholders = append(out.Placeholders, sh)
	}
	out.Decor = rd.decorFor(nil, l)
	return out
}

func (rd *reader) decorFor(slide *Node, l *layoutInfo) []*Shape {
	out := []*Shape{}
	showMaster := slide == nil || slide.A("showMasterSp") != "0"
	if showMaster && l.root.A("showMasterSp") != "0" {
		m := l.master
		sc := &sctx{part: m.part, rels: m.rels, cc: l.cc, level: "master", master: m, idPfx: "dm", decor: true}
		out = append(out, rd.parseTree(m.root.Path("p:cSld", "p:spTree"), sc, identXf)...)
	}
	sc := &sctx{part: l.part, rels: l.rels, cc: l.cc, level: "layout", layout: l, master: l.master, idPfx: "dl", decor: true}
	out = append(out, rd.parseTree(l.root.Path("p:cSld", "p:spTree"), sc, identXf)...)
	return out
}

func (rd *reader) readSlide(part string, n int) (*Slide, error) {
	root := rd.pkg.XML(part)
	if root == nil {
		return nil, fmt.Errorf("missing or invalid part %s", part)
	}
	rels := rd.pkg.Rels(part)
	var l *layoutInfo
	if lr, ok := relByType(rels, rtSlideLayout); ok {
		l = rd.layouts[lr.Target]
		if l == nil {
			l = rd.loadLayout(lr.Target, nil)
		}
	}
	if l == nil {
		if len(rd.layoutOrd) == 0 {
			return nil, fmt.Errorf("slide has no layout")
		}
		l = rd.layoutOrd[0]
	}
	cc := &colorCtx{theme: l.master.theme, clrMap: overrideClrMap(root, l.cc.clrMap)}
	src := part
	s := &Slide{ID: "s" + strconv.Itoa(n), Src: &src, Layout: l.part, Hidden: root.A("show") == "0", Shapes: []*Shape{}}
	s.Bg = rd.resolveBg(root, rels, l)
	s.Decor = rd.decorFor(root, l)
	sc := &sctx{part: part, rels: rels, cc: cc, level: "slide", layout: l, master: l.master, idPfx: "sp"}
	rd.orig[part] = map[string]*origShape{}
	s.Shapes = rd.parseTree(root.Path("p:cSld", "p:spTree"), sc, identXf)
	so := &slideOrig{root: root, rels: rels}
	if nr, ok := relByType(rels, rtNotesSlide); ok {
		so.notesPart = nr.Target
		s.Notes = rd.notesText(nr.Target)
		so.notes = s.Notes
	}
	rd.origSlide[part] = so
	return s, nil
}

func (rd *reader) notesText(part string) string {
	root := rd.pkg.XML(part)
	if root == nil {
		return ""
	}
	var paras []string
	for _, sp := range root.Path("p:cSld", "p:spTree").All("p:sp") {
		ph := sp.Path("p:nvSpPr", "p:nvPr", "p:ph")
		if ph == nil || ph.A("type") != "body" {
			continue
		}
		for _, p := range sp.C("p:txBody").All("a:p") {
			var sb strings.Builder
			for _, c := range p.Elems() {
				switch c.Local {
				case "r", "fld":
					sb.WriteString(c.C("a:t").TextContent())
				case "br":
					sb.WriteString("\n")
				}
			}
			paras = append(paras, sb.String())
		}
	}
	return strings.TrimRight(strings.Join(paras, "\n"), "\n")
}

// --- backgrounds ---------------------------------------------------------

func (rd *reader) resolveBg(slide *Node, slideRels map[string]Rel, l *layoutInfo) *Bg {
	type src struct {
		root *Node
		part string
		rels map[string]Rel
		cc   *colorCtx
	}
	var chain []src
	if slide != nil {
		chain = append(chain, src{slide, "", nil, nil})
	}
	chain = append(chain, src{l.root, l.part, l.rels, l.cc}, src{l.master.root, l.master.part, l.master.rels, l.master.cc})
	for i, s := range chain {
		bg := s.root.Path("p:cSld", "p:bg")
		if bg == nil {
			continue
		}
		rels := s.rels
		cc := s.cc
		if i == 0 && slide != nil {
			rels = slideRels
			cc = &colorCtx{theme: l.master.theme, clrMap: overrideClrMap(slide, l.cc.clrMap)}
		}
		if pr := bg.C("p:bgPr"); pr != nil {
			col, kind, f := cc.fillOf(pr, "")
			switch kind {
			case "solid":
				return &Bg{Color: col}
			case "image":
				if img := rd.blipRef(f.C("a:blip"), rels); img != "" {
					return &Bg{Image: img}
				}
				return nil
			}
			return nil
		}
		if ref := bg.C("p:bgRef"); ref != nil {
			idx := int(atoi64(ref.A("idx")))
			col := cc.resolveColor(firstColor(ref), "")
			var entry *Node
			th := l.master.theme
			if idx >= 1001 && idx-1001 < len(th.bgFill) {
				entry = th.bgFill[idx-1001]
			} else if idx >= 1 && idx-1 < len(th.fillStyle) {
				entry = th.fillStyle[idx-1]
			}
			if entry != nil && entry.Local == "blipFill" {
				if img := rd.blipRef(entry.C("a:blip"), rd.pkg.Rels(l.master.themeP)); img != "" {
					return &Bg{Image: img}
				}
			}
			if entry != nil {
				c, kind, _ := cc.fillOf(&Node{Children: []*Node{entry}}, col)
				if kind == "solid" {
					return &Bg{Color: c}
				}
				if kind == "none" {
					return nil
				}
			}
			return &Bg{Color: col}
		}
	}
	return nil
}

// blipRef returns a media reference for an a:blip, "" if not displayable.
func (rd *reader) blipRef(blip *Node, rels map[string]Rel) string {
	if blip == nil {
		return ""
	}
	rel, ok := rels[blip.A("r:embed")]
	if !ok || rel.External {
		return ""
	}
	return rd.mediaRef(rel.Target)
}

func (rd *reader) mediaRef(part string) string {
	part, ok := rd.pkg.findFile(part)
	if !ok {
		return ""
	}
	ext := path.Ext(part)
	if !browserImage(ext) {
		return ""
	}
	if _, ok := rd.model.Media[part]; !ok {
		rd.model.Media[part] = "data:" + mimeForExt(ext) + ";base64," + base64.StdEncoding.EncodeToString(rd.pkg.Files[part])
	}
	return "media:" + part
}

// --- shapes ----------------------------------------------------------------

type xform struct{ ax, bx, ay, by float64 }

var identXf = xform{1, 0, 1, 0}

func (x xform) box(off, ext *Node) (int64, int64, int64, int64) {
	X, Y := atof(off.A("x")), atof(off.A("y"))
	W, H := atof(ext.A("cx")), atof(ext.A("cy"))
	return int64(math.Round(x.ax*X + x.bx)), int64(math.Round(x.ay*Y + x.by)), int64(math.Round(x.ax * W)), int64(math.Round(x.ay * H))
}

func phOf(n *Node) *Node {
	for _, c := range n.Elems() {
		if strings.HasPrefix(c.Local, "nv") {
			return c.Path("p:nvPr", "p:ph")
		}
	}
	return nil
}

func cNvPrOf(n *Node) *Node {
	for _, c := range n.Elems() {
		if strings.HasPrefix(c.Local, "nv") {
			return c.C("p:cNvPr")
		}
	}
	return nil
}

func phType(ph *Node) string {
	t := ph.A("type")
	if t == "" {
		return "body" // ECMA default is obj; treat it as body content
	}
	if t == "obj" {
		return "body"
	}
	return t
}

func phGroup(t string) string {
	switch t {
	case "title", "ctrTitle":
		return "title"
	case "body", "subTitle", "obj", "":
		return "body"
	}
	return t
}

// matchPh finds the placeholder in tree that a placeholder (typ, idx) inherits from.
func matchPh(tree *Node, typ, idx string, byIdx bool) *Node {
	if tree == nil {
		return nil
	}
	var cands []*Node
	for _, c := range tree.Elems() {
		if ph := phOf(c); ph != nil {
			cands = append(cands, c)
		}
	}
	if byIdx && idx != "" {
		for _, c := range cands {
			if phOf(c).A("idx") == idx {
				return c
			}
		}
	}
	for _, c := range cands {
		if phType(phOf(c)) == typ {
			return c
		}
	}
	g := phGroup(typ)
	for _, c := range cands {
		if phGroup(phType(phOf(c))) == g {
			return c
		}
	}
	return nil
}

// phParents returns [layoutPh, masterPh] (nearest first) for a placeholder.
func phParents(sc *sctx, typ, idx string) []*Node {
	var out []*Node
	switch sc.level {
	case "slide":
		lp := matchPh(sc.layout.root.Path("p:cSld", "p:spTree"), typ, idx, true)
		mt := typ
		if lp != nil {
			out = append(out, lp)
			mt = phType(phOf(lp))
		}
		if mp := matchPh(sc.master.root.Path("p:cSld", "p:spTree"), mt, "", false); mp != nil {
			out = append(out, mp)
		}
	case "layout":
		if mp := matchPh(sc.master.root.Path("p:cSld", "p:spTree"), typ, "", false); mp != nil {
			out = append(out, mp)
		}
	}
	return out
}

func (rd *reader) parseTree(tree *Node, sc *sctx, xf xform) []*Shape {
	out := []*Shape{}
	for _, c := range tree.Elems() {
		sh := rd.parseShape(c, sc, xf)
		if sh == nil {
			continue
		}
		out = append(out, sh)
		if sc.level == "slide" && !sc.decor && xf == identXf {
			js, _ := json.Marshal(sh)
			if _, dup := rd.orig[sc.part][sh.ID]; !dup {
				rd.orig[sc.part][sh.ID] = &origShape{node: c, rels: sc.rels, json: string(js)}
			}
		}
	}
	return out
}

func (rd *reader) parseShape(n *Node, sc *sctx, xf xform) *Shape {
	if n.URI != nsP && n.URI != nsMC {
		return nil
	}
	if sc.decor {
		if phOf(n) != nil {
			return nil
		}
	}
	switch {
	case n.Is("p:sp"), n.Is("p:cxnSp"):
		return rd.parseSp(n, sc, xf)
	case n.Is("p:pic"):
		return rd.parsePic(n, sc, xf)
	case n.Is("p:graphicFrame"):
		return rd.parseFrame(n, sc, xf)
	case n.Is("p:grpSp"):
		return rd.parseGroup(n, sc, xf)
	case n.Is("mc:AlternateContent"):
		return rd.parseAlt(n, sc, xf)
	case n.Is("p:contentPart"):
		return nil
	}
	return nil
}

func (rd *reader) baseShape(n *Node, sc *sctx) *Shape {
	sh := &Shape{}
	cnv := cNvPrOf(n)
	sh.ID = sc.idPfx + cnv.A("id")
	sh.Name = cnv.A("name")
	sh.Descr = cnv.A("descr")
	if ph := phOf(n); ph != nil {
		t := ph.A("type")
		if t == "" {
			t = "body"
		}
		sh.Ph = &t
		sh.PhIdx = ph.A("idx")
	}
	return sh
}

// xfrmFrom finds geometry in spPr/xfrm (or p:xfrm) and inherited placeholders.
func xfrmOf(n *Node) *Node {
	if sp := n.C("p:spPr"); sp != nil {
		if x := sp.C("a:xfrm"); x != nil {
			return x
		}
	}
	if x := n.C("p:xfrm"); x != nil {
		return x
	}
	if g := n.C("p:grpSpPr"); g != nil {
		return g.C("a:xfrm")
	}
	return nil
}

func (sh *Shape) setXfrm(x *Node, xf xform) bool {
	if x == nil || x.C("a:off") == nil || x.C("a:ext") == nil {
		return false
	}
	sh.X, sh.Y, sh.W, sh.H = xf.box(x.C("a:off"), x.C("a:ext"))
	sh.Rot = round2(atof(x.A("rot")) / 60000)
	sh.FlipH = parseBool(x.A("flipH"))
	sh.FlipV = parseBool(x.A("flipV"))
	return true
}

type bodyProps struct {
	ins       [4]int64
	anchor    string
	wrap      string
	vert      string
	autofit   string
	fontScale float64
	lnRed     float64
}

func bodyEff(nodes ...*Node) bodyProps {
	b := bodyProps{ins: [4]int64{91440, 45720, 91440, 45720}, anchor: "t", wrap: "square", vert: "horz", autofit: "none"}
	for _, n := range nodes {
		if n == nil {
			continue
		}
		for i, k := range []string{"lIns", "tIns", "rIns", "bIns"} {
			if v, ok := n.AttrOK(k); ok {
				b.ins[i] = atoi64(v)
			}
		}
		if v := n.A("anchor"); v != "" {
			b.anchor = v
		}
		if v := n.A("wrap"); v != "" {
			b.wrap = v
		}
		if v := n.A("vert"); v != "" {
			b.vert = v
		}
		if a := n.C("a:normAutofit"); a != nil {
			b.autofit = "norm"
			b.fontScale = atof(a.A("fontScale")) / 1000
			b.lnRed = atof(a.A("lnSpcReduction")) / 1000
		} else if n.C("a:spAutoFit") != nil {
			b.autofit, b.fontScale, b.lnRed = "shape", 0, 0
		} else if n.C("a:noAutofit") != nil {
			b.autofit, b.fontScale, b.lnRed = "none", 0, 0
		}
	}
	return b
}

func (sh *Shape) setBody(b bodyProps) {
	sh.Anchor = b.anchor
	if sh.Anchor == "just" || sh.Anchor == "dist" {
		sh.Anchor = "ctr"
	}
	sh.Inset = []int64{b.ins[0], b.ins[1], b.ins[2], b.ins[3]}
	sh.Autofit = b.autofit
	sh.FontScale = round2(b.fontScale)
	sh.LnSpcReduction = round2(b.lnRed)
	sh.Wrap = b.wrap
	if b.vert != "horz" {
		sh.Vert = b.vert
	}
}

// textChainFor builds the inherited text style chain for a shape.
func (rd *reader) textChainFor(sc *sctx, ph *string, parents []*Node, style *Node, isTxBox bool) *textChain {
	layers := []layer{{list: rd.defText}}
	if ph != nil {
		var ts *Node
		txs := sc.master.root.C("p:txStyles")
		switch phGroup(*ph) {
		case "title":
			ts = txs.C("p:titleStyle")
		case "body":
			ts = txs.C("p:bodyStyle")
		default:
			ts = txs.C("p:otherStyle")
		}
		layers = append(layers, layer{list: ts})
		for i := len(parents) - 1; i >= 0; i-- {
			if ls := parents[i].Path("p:txBody", "a:lstStyle"); ls != nil {
				layers = append(layers, layer{list: ls})
			}
		}
	} else {
		th := sc.cc.theme
		def := th.spDef
		if isTxBox {
			def = th.txDef
		}
		if ls := def.C("a:lstStyle"); ls != nil {
			layers = append(layers, layer{list: ls})
		}
	}
	if fr := style.C("a:fontRef"); fr != nil {
		if c := firstColor(fr); c != nil {
			layers = append(layers, layer{r: &rProps{Color: strp(sc.cc.resolveColor(c, ""))}})
		}
	}
	return newChain(sc.cc, layers...)
}

func (rd *reader) parseSp(n *Node, sc *sctx, xf xform) *Shape {
	sh := rd.baseShape(n, sc)
	var parents []*Node
	if sh.Ph != nil {
		parents = phParents(sc, phType(phOf(n)), sh.PhIdx)
	}
	if !sh.setXfrm(xfrmOf(n), xf) {
		for _, p := range parents {
			if sh.setXfrm(xfrmOf(p), xf) {
				break
			}
		}
	}
	spPr := n.C("p:spPr")
	sh.Geom = "rect"
	if pg := spPr.C("a:prstGeom"); pg != nil {
		sh.Geom = pg.A("prst")
	} else if cg := spPr.C("a:custGeom"); cg != nil {
		sh.Geom = "custom"
		sh.CustGeom = cg.Serialize()
		sh.Path = custGeomPath(cg, sh.W, sh.H)
	} else {
		for _, p := range parents {
			if pg := p.Path("p:spPr", "a:prstGeom"); pg != nil {
				sh.Geom = pg.A("prst")
				break
			}
		}
	}
	style := n.C("p:style")
	// fill
	fillDone := false
	for _, src := range append([]*Node{n}, parents...) {
		c, kind, f := sc.cc.fillOf(src.C("p:spPr"), "")
		switch kind {
		case "solid":
			sh.Fill = strp(c)
		case "image":
			sh.FillImage = rd.blipRef(f.C("a:blip"), relsFor(rd, src, sc))
		case "none", "group":
		default:
			continue
		}
		fillDone = true
		break
	}
	if !fillDone && style != nil {
		sh.Fill, _ = fillFromRef(style.C("a:fillRef"), sc.cc)
	}
	// line
	var line *Line
	if style != nil {
		line = lineFromRef(style.C("a:lnRef"), sc.cc)
	}
	for i := len(parents) - 1; i >= 0; i-- {
		line = readLine(parents[i].Path("p:spPr", "a:ln"), sc.cc, "", line)
	}
	sh.Line = readLine(spPr.C("a:ln"), sc.cc, "", line)
	isTxBox := parseBool(n.Path("p:nvSpPr", "p:cNvSpPr").A("txBox"))
	if sh.Ph != nil || isTxBox {
		sh.Kind = "text"
	} else {
		sh.Kind = "shape"
	}
	if n.Is("p:cxnSp") {
		sh.Kind = "shape"
		sh.Fill = nil
	}
	var bodies []*Node
	for i := len(parents) - 1; i >= 0; i-- {
		bodies = append(bodies, parents[i].Path("p:txBody", "a:bodyPr"))
	}
	if !isTxBox && sh.Ph == nil {
		bodies = append([]*Node{nil}, bodies...)
	}
	tx := n.C("p:txBody")
	bodies = append(bodies, tx.C("a:bodyPr"))
	sh.setBody(bodyEff(bodies...))
	if isTxBox && tx.C("a:bodyPr") != nil && tx.C("a:bodyPr").A("wrap") == "" {
		sh.Wrap = "square"
	}
	if tx != nil {
		tc := rd.textChainFor(sc, sh.Ph, parents, style, isTxBox)
		sh.Paras = readParas(tx, tc, sc.rels)
	}
	if sh.Paras == nil && (sh.Kind == "text") {
		sh.Paras = []*Para{}
	}
	return sh
}

// relsFor returns the rels of the part that owns node src.
func relsFor(rd *reader, src *Node, sc *sctx) map[string]Rel {
	root := src
	for root.Parent != nil {
		root = root.Parent
	}
	for name, n := range rd.pkg.dom {
		if n == root {
			return rd.pkg.Rels(name)
		}
	}
	return sc.rels
}

func (rd *reader) parsePic(n *Node, sc *sctx, xf xform) *Shape {
	sh := rd.baseShape(n, sc)
	var parents []*Node
	if sh.Ph != nil {
		parents = phParents(sc, phType(phOf(n)), sh.PhIdx)
	}
	if !sh.setXfrm(xfrmOf(n), xf) {
		for _, p := range parents {
			if sh.setXfrm(xfrmOf(p), xf) {
				break
			}
		}
	}
	bf := n.C("p:blipFill")
	blip := bf.C("a:blip")
	img := rd.blipRef(blip, sc.rels)
	if img == "" {
		return rd.opaque(n, sh, sc)
	}
	sh.Kind = "image"
	sh.Src = img
	if sr := bf.C("a:srcRect"); sr != nil {
		c := &Crop{L: pct(sr.A("l")), T: pct(sr.A("t")), R: pct(sr.A("r")), B: pct(sr.A("b"))}
		if *c != (Crop{}) {
			sh.Crop = c
		}
	}
	sh.Geom = "rect"
	if pg := n.Path("p:spPr", "a:prstGeom"); pg != nil {
		sh.Geom = pg.A("prst")
	}
	var line *Line
	if st := n.C("p:style"); st != nil {
		line = lineFromRef(st.C("a:lnRef"), sc.cc)
	}
	sh.Line = readLine(n.Path("p:spPr", "a:ln"), sc.cc, "", line)
	return sh
}

func (rd *reader) parseFrame(n *Node, sc *sctx, xf xform) *Shape {
	sh := rd.baseShape(n, sc)
	sh.setXfrm(n.C("p:xfrm"), xf)
	gd := n.Path("a:graphic", "a:graphicData")
	if tbl := gd.C("a:tbl"); tbl != nil {
		if xf != identXf {
			// tables inside groups: scale columns for display
			rd.parseTable(sh, tbl, sc)
			for i := range sh.ColW {
				sh.ColW[i] = int64(float64(sh.ColW[i]) * xf.ax)
			}
			for i := range sh.RowH {
				sh.RowH[i] = int64(float64(sh.RowH[i]) * xf.ay)
			}
			return sh
		}
		rd.parseTable(sh, tbl, sc)
		return sh
	}
	return rd.opaque(n, sh, sc)
}

func (rd *reader) parseGroup(n *Node, sc *sctx, xf xform) *Shape {
	sh := rd.baseShape(n, sc)
	x := n.Path("p:grpSpPr", "a:xfrm")
	child := xf
	if sh.setXfrm(x, xf) {
		off, ext := x.C("a:off"), x.C("a:ext")
		chOff, chExt := x.C("a:chOff"), x.C("a:chExt")
		sx, sy := 1.0, 1.0
		if chExt != nil && atof(chExt.A("cx")) > 0 {
			sx = atof(ext.A("cx")) / atof(chExt.A("cx"))
		}
		if chExt != nil && atof(chExt.A("cy")) > 0 {
			sy = atof(ext.A("cy")) / atof(chExt.A("cy"))
		}
		cox, coy := atof(chOff.A("x")), atof(chOff.A("y"))
		gx := xform{ax: sx, bx: atof(off.A("x")) - cox*sx, ay: sy, by: atof(off.A("y")) - coy*sy}
		child = xform{ax: xf.ax * gx.ax, bx: xf.ax*gx.bx + xf.bx, ay: xf.ay * gx.ay, by: xf.ay*gx.by + xf.by}
	}
	csc := *sc
	csc.idPfx = sc.idPfx + cNvPrOf(n).A("id") + "_"
	csc.decor = false
	sh.Children = rd.parseTree(n, &csc, child)
	if sh.Children == nil {
		sh.Children = []*Shape{}
	}
	return rd.opaque(n, sh, sc)
}

func (rd *reader) parseAlt(n *Node, sc *sctx, xf xform) *Shape {
	choice := n.C("mc:Choice")
	fb := n.C("mc:Fallback")
	var inner *Node
	for _, c := range []*Node{choice, fb} {
		for _, e := range c.Elems() {
			inner = e
			break
		}
		if inner != nil {
			break
		}
	}
	if inner == nil {
		return nil
	}
	if sc.decor && phOf(inner) != nil {
		return nil
	}
	sh := rd.baseShape(inner, sc)
	if !sh.setXfrm(xfrmOf(inner), xf) {
		if sh.Ph != nil {
			for _, p := range phParents(sc, phType(phOf(inner)), sh.PhIdx) {
				if sh.setXfrm(xfrmOf(p), xf) {
					break
				}
			}
		}
	}
	// preview from the fallback's picture (e.g. equations rendered as images)
	if fb != nil {
		fb.Walk(func(e *Node) {
			if sh.Preview == "" && e.Is("a:blip") {
				sh.Preview = rd.blipRef(e, sc.rels)
			}
		})
	}
	if fb != nil && sh.Preview == "" {
		// text-only fallback: parse it so the frontend can render its text
		var fsh *Shape
		for _, e := range fb.Elems() {
			fsh = rd.parseShape(e, &sctx{part: sc.part, rels: sc.rels, cc: sc.cc, level: sc.level, layout: sc.layout, master: sc.master, idPfx: sc.idPfx + "fb"}, xf)
			break
		}
		if fsh != nil {
			sh.Children = []*Shape{fsh}
		}
	}
	return rd.opaque(n, sh, sc)
}

// opaque finalizes sh as an opaque shape carrying n verbatim.
func (rd *reader) opaque(n *Node, sh *Shape, sc *sctx) *Shape {
	sh.Kind = "opaque"
	sh.XML = n.Serialize()
	sh.Rels = collectRels(n, sc.rels)
	if sh.Preview == "" {
		n.Walk(func(e *Node) {
			if sh.Preview == "" && e.Is("a:blip") {
				sh.Preview = rd.blipRef(e, sc.rels)
			}
		})
	}
	return sh
}

// collectRels returns the relationships referenced by r:* attributes in n.
func collectRels(n *Node, rels map[string]Rel) map[string]Rel {
	out := map[string]Rel{}
	n.Walk(func(e *Node) {
		for _, a := range e.Attrs {
			if a.URI != nsR || a.Value == "" {
				continue
			}
			if rel, ok := rels[a.Value]; ok {
				out[a.Value] = rel
			}
		}
	})
	if len(out) == 0 {
		return nil
	}
	return out
}
