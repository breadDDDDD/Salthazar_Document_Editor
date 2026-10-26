package pptx

import (
	"crypto/sha1"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

type writer struct {
	rd        *reader
	out       *Pkg
	model     *Model
	media     map[string]string // sha1 -> part
	srcMap    map[string]string // original slide part -> new slide part
	newCT     map[string]string // part -> content type override
	usedOrig  map[string]int    // original slide part -> times used
	slideRefs []string          // new slide parts in order

	addNotesMaster string // notes master part added from the template
}

// FromModel builds a .pptx from the model. When original is given, its
// masters, layouts, theme, media and other parts are kept and only the
// slides are regenerated.
func FromModel(m *Model, original []byte) ([]byte, error) {
	if m == nil {
		return nil, fmt.Errorf("missing model")
	}
	var base *Pkg
	var err error
	if len(original) > 0 {
		base, err = openPkg(original)
		if err != nil {
			return nil, err
		}
	} else {
		base = templatePkg()
	}
	rd, err := newReaderPkg(base)
	if err != nil {
		return nil, err
	}
	rd.model = &Model{Media: map[string]string{}}
	prels := base.Rels(rd.presPart)
	var origSlides []string
	n := 0
	for _, sid := range rd.pres.C("p:sldIdLst").All("p:sldId") {
		if rel, ok := prels[sid.A("r:id")]; ok {
			n++
			origSlides = append(origSlides, rel.Target)
			if _, err := rd.readSlide(rel.Target, n); err != nil {
				return nil, fmt.Errorf("original slide %d: %v", n, err)
			}
		}
	}
	w := &writer{rd: rd, model: m, media: map[string]string{}, srcMap: map[string]string{},
		newCT: map[string]string{}, usedOrig: map[string]int{}}
	w.out = newPkg()
	removed := map[string]bool{}
	for _, s := range origSlides {
		removed[s] = true
		removed[relsPartName(s)] = true
		if so := rd.origSlide[s]; so != nil && so.notesPart != "" {
			removed[so.notesPart] = true
			removed[relsPartName(so.notesPart)] = true
		}
	}
	// Also drop any other slide/notes rels pointing at removed parts.
	for _, name := range base.Order {
		if removed[name] {
			continue
		}
		w.out.Order = append(w.out.Order, name)
		w.out.Files[name] = base.Files[name]
	}
	for name, data := range base.Files {
		if !removed[name] {
			if _, ok := w.out.Files[name]; !ok {
				w.out.Files[name] = data
				w.out.Order = append(w.out.Order, name)
			}
		}
	}
	// new part names
	for i, s := range m.Slides {
		part := fmt.Sprintf("ppt/slides/slide%d.xml", i+1)
		w.slideRefs = append(w.slideRefs, part)
		if s != nil && s.Src != nil {
			if _, ok := w.srcMap[*s.Src]; !ok {
				w.srcMap[*s.Src] = part
			}
		}
	}
	for i, s := range m.Slides {
		if s == nil {
			return nil, fmt.Errorf("slide %d is null", i+1)
		}
		if _, err := w.writeSlide(s, i); err != nil {
			return nil, fmt.Errorf("slide %d: %v", i+1, err)
		}
	}
	if err := w.writePresentation(); err != nil {
		return nil, err
	}
	w.writeContentTypes(removed)
	w.writeAppProps()
	return w.out.zip()
}

// --- slide -------------------------------------------------------------

type slideWriter struct {
	w      *writer
	rd     *reader
	part   string
	rels   []Rel
	relKey map[string]string
	cc     *colorCtx
	layout *layoutInfo
	sc     *sctx
	used   map[int]bool
	next   int
}

func (sw *slideWriter) addRel(typ, target string, external bool) string {
	key := typ + "|" + target + "|" + strconv.FormatBool(external)
	if id, ok := sw.relKey[key]; ok {
		return id
	}
	id := "rId" + strconv.Itoa(len(sw.rels)+1)
	sw.rels = append(sw.rels, Rel{ID: id, Type: typ, Target: target, External: external})
	sw.relKey[key] = id
	return id
}

func (sw *slideWriter) hyperlink(url string) string {
	return sw.addRel(rtHyperlink, url, true)
}

// mapRel re-creates an original relationship for the new slide part.
func (sw *slideWriter) mapRel(r Rel) (string, bool) {
	if r.External {
		return sw.addRel(r.Type, r.Target, true), true
	}
	if sameRelKind(r.Type, rtSlide) {
		if np, ok := sw.w.srcMap[r.Target]; ok {
			return sw.addRel(rtSlide, np, false), true
		}
		if len(sw.w.slideRefs) > 0 {
			return sw.addRel(rtSlide, sw.w.slideRefs[0], false), true
		}
		return "", false
	}
	if sameRelKind(r.Type, rtNotesSlide) || sameRelKind(r.Type, rtSlideLayout) {
		return "", false
	}
	if _, ok := sw.w.out.Files[r.Target]; !ok {
		return "", false
	}
	return sw.addRel(r.Type, r.Target, false), true
}

// remap parses a fragment, rewrites r:* attributes through rels and
// returns the serialized XML plus the cNvPr ids used inside.
func (sw *slideWriter) remap(xmlText string, rels map[string]Rel) (string, []int, error) {
	n, err := parseXML([]byte(xmlText))
	if err != nil {
		return "", nil, err
	}
	var ids []int
	n.Walk(func(e *Node) {
		if e.Local == "cNvPr" {
			if v, err := strconv.Atoi(e.A("id")); err == nil {
				ids = append(ids, v)
			}
		}
		for i, a := range e.Attrs {
			if a.URI != nsR || a.Value == "" {
				continue
			}
			if r, ok := rels[a.Value]; ok {
				if nid, ok := sw.mapRel(r); ok {
					e.Attrs[i].Value = nid
					continue
				}
			}
			e.Attrs[i].Value = ""
		}
	})
	return n.Serialize(), ids, nil
}

func (sw *slideWriter) newID(pref string) int {
	if v, err := strconv.Atoi(strings.TrimLeft(pref, "abcdefghijklmnopqrstuvwxyz_")); err == nil && v > 1 && !sw.used[v] && !strings.Contains(pref, "_") {
		sw.used[v] = true
		return v
	}
	for sw.used[sw.next] || sw.next < 2 {
		sw.next++
	}
	sw.used[sw.next] = true
	return sw.next
}

func (w *writer) findLayout(s *Slide) *layoutInfo {
	rd := w.rd
	if l, ok := rd.layouts[s.Layout]; ok {
		return l
	}
	name := ""
	for _, ml := range w.model.Layouts {
		if ml != nil && ml.ID == s.Layout {
			name = ml.Name
		}
	}
	if name == "" {
		name = s.Layout
	}
	for _, l := range rd.layoutOrd {
		if strings.EqualFold(l.name, name) {
			return l
		}
	}
	hasTitle, hasCtr, hasBody := false, false, false
	for _, sh := range s.Shapes {
		if sh != nil && sh.Ph != nil {
			switch phGroup(*sh.Ph) {
			case "title":
				hasTitle = true
				if *sh.Ph == "ctrTitle" {
					hasCtr = true
				}
			case "body":
				hasBody = true
			}
		}
	}
	want := "Blank"
	if hasCtr {
		want = "Title Slide"
	} else if hasTitle && hasBody {
		want = "Title and Content"
	} else if hasTitle {
		want = "Title Only"
	}
	for _, l := range rd.layoutOrd {
		if strings.EqualFold(l.name, want) {
			return l
		}
	}
	for _, l := range rd.layoutOrd {
		if strings.EqualFold(l.root.A("type"), map[string]string{"Title Slide": "title", "Title and Content": "obj", "Blank": "blank", "Title Only": "titleOnly"}[want]) {
			return l
		}
	}
	return rd.layoutOrd[0]
}

type piece struct {
	text string
	sh   *Shape
}

func (w *writer) writeSlide(s *Slide, idx int) (string, error) {
	rd := w.rd
	part := w.slideRefs[idx]
	l := w.findLayout(s)
	var so *slideOrig
	if s.Src != nil {
		so = rd.origSlide[*s.Src]
		w.usedOrig[*s.Src]++
	}
	first := so != nil && w.usedOrig[*s.Src] == 1
	cc := &colorCtx{theme: l.master.theme, clrMap: l.cc.clrMap}
	if so != nil {
		cc.clrMap = overrideClrMap(so.root, l.cc.clrMap)
	}
	sw := &slideWriter{w: w, rd: rd, part: part, relKey: map[string]string{}, cc: cc, layout: l,
		used: map[int]bool{1: true}, next: 2}
	sw.sc = &sctx{part: part, cc: cc, level: "slide", layout: l, master: l.master, idPfx: "sp"}
	sw.addRel(rtSlideLayout, l.part, false)

	// pass 1: verbatim pieces
	var pieces []piece
	for _, sh := range s.Shapes {
		if sh == nil {
			continue
		}
		if s.Src != nil {
			if o := rd.orig[*s.Src][sh.ID]; o != nil {
				js, _ := json.Marshal(sh)
				if string(js) == o.json {
					x, ids, err := sw.remap(o.node.Serialize(), o.rels)
					if err == nil {
						for _, id := range ids {
							sw.used[id] = true
						}
						pieces = append(pieces, piece{text: x})
						continue
					}
				}
			}
		}
		if sh.Kind == "opaque" {
			if sh.XML == "" {
				continue
			}
			x, ids, err := sw.remap(sh.XML, sh.Rels)
			if err != nil {
				continue
			}
			for _, id := range ids {
				sw.used[id] = true
			}
			pieces = append(pieces, piece{text: x})
			continue
		}
		pieces = append(pieces, piece{sh: sh})
	}
	// pass 2: generated shapes
	var tree strings.Builder
	for _, p := range pieces {
		if p.sh == nil {
			tree.WriteString(p.text)
			continue
		}
		id := sw.newID(p.sh.ID)
		if err := sw.writeShape(&tree, p.sh, id); err != nil {
			return "", err
		}
	}

	var sb strings.Builder
	sb.WriteString(xmlDecl)
	sb.WriteString(`<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`)
	if so != nil {
		if v := so.root.A("showMasterSp"); v != "" {
			sb.WriteString(` showMasterSp="` + esc(v) + `"`)
		}
	}
	if s.Hidden {
		sb.WriteString(` show="0"`)
	}
	sb.WriteString(`><p:cSld`)
	if so != nil {
		if nm := so.root.C("p:cSld").A("name"); nm != "" {
			sb.WriteString(` name="` + esc(nm) + `"`)
		}
	}
	sb.WriteString(`>`)
	inhBg := rd.resolveBg(nil, nil, l)
	if s.Bg != nil && (s.Bg.Color != "" || s.Bg.Image != "") && !jsonEq(s.Bg, inhBg) {
		if s.Bg.Image != "" {
			rid, err := sw.imageRel(s.Bg.Image)
			if err != nil {
				return "", err
			}
			fmt.Fprintf(&sb, `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`, rid)
		} else {
			fmt.Fprintf(&sb, `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`, hexVal(s.Bg.Color))
		}
	} else if so != nil {
		// keep an explicit original background that resolves to the same thing
		if bg := so.root.Path("p:cSld", "p:bg"); bg != nil && jsonEq(s.Bg, rd.resolveBg(so.root, so.rels, l)) && !jsonEq(s.Bg, inhBg) {
			if x, _, err := sw.remap(bg.Serialize(), so.rels); err == nil {
				sb.WriteString(x)
			}
		}
	}
	sb.WriteString(`<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`)
	sb.WriteString(tree.String())
	sb.WriteString(`</p:spTree>`)
	if first {
		if ext := so.root.Path("p:cSld", "p:extLst"); ext != nil {
			sb.WriteString(ext.Serialize())
		}
	}
	sb.WriteString(`</p:cSld>`)
	wroteClr := false
	if so != nil {
		for _, c := range so.root.Elems() {
			switch {
			case c.Is("p:clrMapOvr"):
				sb.WriteString(c.Serialize())
				wroteClr = true
			case c.Is("p:transition"), c.Is("mc:AlternateContent") && c.Find("p:transition") != nil:
				if !wroteClr {
					sb.WriteString(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`)
					wroteClr = true
				}
				if x, _, err := sw.remap(c.Serialize(), so.rels); err == nil {
					sb.WriteString(x)
				}
			case c.Is("p:extLst"):
				if first {
					if !wroteClr {
						sb.WriteString(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`)
						wroteClr = true
					}
					sb.WriteString(c.Serialize())
				}
			}
		}
	}
	if !wroteClr {
		// must come before transition/timing/extLst; only reached when none were written
		sb.WriteString(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`)
	}
	sb.WriteString(`</p:sld>`)
	slideXML := sb.String()

	// notes
	notesPart := ""
	if strings.TrimSpace(s.Notes) != "" {
		notesPart = fmt.Sprintf("ppt/notesSlides/notesSlide%d.xml", idx+1)
		if err := w.writeNotes(notesPart, part, s, so); err != nil {
			return "", err
		}
		sw.addRel(rtNotesSlide, notesPart, false)
	}
	w.out.put(part, []byte(slideXML))
	w.out.put(relsPartName(part), writeRels(sw.rels, part))
	w.newCT[part] = ctSlide
	return notesPart, nil
}

func (p *Pkg) put(name string, data []byte) {
	if _, ok := p.Files[name]; !ok {
		p.Order = append(p.Order, name)
	}
	p.Files[name] = data
	delete(p.dom, name)
	delete(p.rels, strings.TrimSuffix(strings.Replace(name, "_rels/", "", 1), ".rels"))
}

func nameOr(s, d string) string {
	if s == "" {
		return d
	}
	return s
}

func (sw *slideWriter) writeShape(sb *strings.Builder, sh *Shape, id int) error {
	switch sh.Kind {
	case "image":
		return sw.writePic(sb, sh, id)
	case "table":
		sw.writeTable(sb, sh, id)
		return nil
	default:
		return sw.writeSp(sb, sh, id)
	}
}

func xfrmXML(sh *Shape) string {
	var attrs string
	if sh.Rot != 0 {
		attrs += fmt.Sprintf(` rot="%d"`, int64(sh.Rot*60000))
	}
	if sh.FlipH {
		attrs += ` flipH="1"`
	}
	if sh.FlipV {
		attrs += ` flipV="1"`
	}
	w, h := sh.W, sh.H
	if w < 0 {
		w = 0
	}
	if h < 0 {
		h = 0
	}
	return fmt.Sprintf(`<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, attrs, sh.X, sh.Y, w, h)
}

func (sw *slideWriter) phXML(sh *Shape, parents []*Node) string {
	if sh.Ph == nil {
		return ""
	}
	t := *sh.Ph
	attrs := ""
	omitType := t == "body" && len(parents) > 0 && phOf(parents[0]).A("type") == "" && sh.PhIdx != ""
	if t == "obj" {
		omitType = true
	}
	if !omitType {
		attrs += ` type="` + esc(t) + `"`
	}
	if sh.PhIdx != "" {
		attrs += ` idx="` + esc(sh.PhIdx) + `"`
	} else if len(parents) > 0 && phOf(parents[0]).A("idx") != "" && phGroup(t) != "title" {
		attrs += ` idx="` + esc(phOf(parents[0]).A("idx")) + `"`
	}
	return `<p:ph` + attrs + `/>`
}

func (sw *slideWriter) inheritedGeom(sh *Shape, parents []*Node) (*Shape, bool) {
	tmp := &Shape{}
	for _, p := range parents {
		if tmp.setXfrm(xfrmOf(p), identXf) {
			return tmp, true
		}
	}
	return tmp, false
}

func (sw *slideWriter) writeSp(sb *strings.Builder, sh *Shape, id int) error {
	var parents []*Node
	if sh.Ph != nil {
		typ := *sh.Ph
		if typ == "obj" {
			typ = "body"
		}
		parents = phParents(sw.sc, typ, sh.PhIdx)
	}
	isTxBox := sh.Kind == "text" && sh.Ph == nil
	fmt.Fprintf(sb, `<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"`, id, esc(nameOr(sh.Name, fmt.Sprintf("Shape %d", id-1))))
	if sh.Descr != "" {
		fmt.Fprintf(sb, ` descr="%s"`, esc(sh.Descr))
	}
	sb.WriteString(`/><p:cNvSpPr`)
	if isTxBox {
		sb.WriteString(` txBox="1"`)
	}
	if sh.Ph != nil {
		sb.WriteString(`><a:spLocks noGrp="1"/></p:cNvSpPr>`)
	} else {
		sb.WriteString(`/>`)
	}
	sb.WriteString(`<p:nvPr>` + sw.phXML(sh, parents) + `</p:nvPr></p:nvSpPr><p:spPr>`)
	inhG, hasInh := sw.inheritedGeom(sh, parents)
	sameAsInh := hasInh && inhG.X == sh.X && inhG.Y == sh.Y && inhG.W == sh.W && inhG.H == sh.H
	noGeom := sh.W == 0 && sh.H == 0 && sh.X == 0 && sh.Y == 0
	if !(sh.Ph != nil && (sameAsInh || noGeom) && sh.Rot == 0 && !sh.FlipH && !sh.FlipV) {
		sb.WriteString(xfrmXML(sh))
	}
	geom := sh.Geom
	if geom == "" {
		geom = "rect"
	}
	inhGeom := "rect"
	for _, p := range parents {
		if pg := p.Path("p:spPr", "a:prstGeom"); pg != nil {
			inhGeom = pg.A("prst")
			break
		}
	}
	switch {
	case geom == "custom" && sh.CustGeom != "":
		if x, _, err := sw.remap(sh.CustGeom, nil); err == nil {
			sb.WriteString(x)
		} else {
			sb.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
		}
	case geom == "custom":
		sb.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
	case sh.Ph == nil || geom != inhGeom:
		fmt.Fprintf(sb, `<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>`, esc(geom))
	}
	// fill
	var inhFill *string
	inhFillImage := ""
	var inhLine *Line
	for i := len(parents) - 1; i >= 0; i-- {
		inhLine = readLine(parents[i].Path("p:spPr", "a:ln"), sw.cc, "", inhLine)
	}
	for _, p := range parents {
		c, kind, f := sw.cc.fillOf(p.C("p:spPr"), "")
		if kind == "solid" {
			inhFill = strp(c)
		} else if kind == "image" {
			inhFillImage = sw.rd.blipRef(f.C("a:blip"), relsFor(sw.rd, p, sw.sc))
		} else if kind == "" {
			continue
		}
		break
	}
	if sh.FillImage != "" && sh.FillImage != inhFillImage {
		rid, err := sw.imageRel(sh.FillImage)
		if err != nil {
			return err
		}
		fmt.Fprintf(sb, `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, rid)
	} else if sh.FillImage == "" && (!strPtrEq(sh.Fill, inhFill) || inhFillImage != "") {
		if sh.Fill == nil {
			sb.WriteString(`<a:noFill/>`)
		} else {
			fmt.Fprintf(sb, `<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, hexVal(*sh.Fill))
		}
	}
	if !linePtrEq(sh.Line, inhLine) {
		sb.WriteString(lineXML("a:ln", sh.Line))
	}
	sb.WriteString(`</p:spPr>`)
	if len(sh.Paras) > 0 || sh.Ph != nil || isTxBox {
		var bodies []*Node
		for i := len(parents) - 1; i >= 0; i-- {
			bodies = append(bodies, parents[i].Path("p:txBody", "a:bodyPr"))
		}
		inh := bodyEff(bodies...)
		var attrs strings.Builder
		names := []string{"lIns", "tIns", "rIns", "bIns"}
		for i := 0; i < 4 && i < len(sh.Inset); i++ {
			if sh.Inset[i] != inh.ins[i] {
				fmt.Fprintf(&attrs, ` %s="%d"`, names[i], sh.Inset[i])
			}
		}
		if sh.Wrap != "" && sh.Wrap != inh.wrap {
			fmt.Fprintf(&attrs, ` wrap="%s"`, esc(sh.Wrap))
		} else if isTxBox && sh.Wrap == "" {
			attrs.WriteString(` wrap="square"`)
		}
		vert := sh.Vert
		if vert == "" {
			vert = "horz"
		}
		if vert != inh.vert {
			fmt.Fprintf(&attrs, ` vert="%s"`, esc(vert))
		}
		if sh.Anchor != "" && sh.Anchor != inh.anchor {
			fmt.Fprintf(&attrs, ` anchor="%s"`, esc(sh.Anchor))
		}
		auto := ""
		af := sh.Autofit
		if af == "" {
			af = inh.autofit
		}
		if af != inh.autofit || sh.FontScale > 0 || sh.LnSpcReduction > 0 || (sh.Ph == nil && af != "none") {
			switch af {
			case "norm":
				auto = `<a:normAutofit`
				if sh.FontScale > 0 && sh.FontScale < 100 {
					auto += fmt.Sprintf(` fontScale="%d"`, int64(sh.FontScale*1000))
				}
				if sh.LnSpcReduction > 0 {
					auto += fmt.Sprintf(` lnSpcReduction="%d"`, int64(sh.LnSpcReduction*1000))
				}
				auto += `/>`
			case "shape":
				auto = `<a:spAutoFit/>`
			default:
				auto = `<a:noAutofit/>`
			}
		}
		sb.WriteString(`<p:txBody><a:bodyPr` + attrs.String() + `>` + auto + `</a:bodyPr><a:lstStyle/>`)
		tc := sw.rd.textChainFor(sw.sc, sh.Ph, parents, nil, isTxBox)
		writeParas(sb, sh.Paras, tc, sw)
		sb.WriteString(`</p:txBody>`)
	}
	sb.WriteString(`</p:sp>`)
	return nil
}

func (sw *slideWriter) phParentsOf(sh *Shape) []*Node {
	if sh.Ph == nil {
		return nil
	}
	typ := *sh.Ph
	if typ == "obj" {
		typ = "body"
	}
	return phParents(sw.sc, typ, sh.PhIdx)
}

func (sw *slideWriter) writePic(sb *strings.Builder, sh *Shape, id int) error {
	rid, err := sw.imageRel(sh.Src)
	if err != nil {
		return err
	}
	var parents []*Node
	if sh.Ph != nil {
		parents = phParents(sw.sc, *sh.Ph, sh.PhIdx)
	}
	fmt.Fprintf(sb, `<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"`, id, esc(nameOr(sh.Name, fmt.Sprintf("Picture %d", id-1))))
	if sh.Descr != "" {
		fmt.Fprintf(sb, ` descr="%s"`, esc(sh.Descr))
	}
	sb.WriteString(`/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr>` + sw.phXML(sh, parents) + `</p:nvPr></p:nvPicPr>`)
	fmt.Fprintf(sb, `<p:blipFill><a:blip r:embed="%s"/>`, rid)
	if c := sh.Crop; c != nil && *c != (Crop{}) {
		fmt.Fprintf(sb, `<a:srcRect l="%d" t="%d" r="%d" b="%d"/>`, int64(c.L*100000), int64(c.T*100000), int64(c.R*100000), int64(c.B*100000))
	}
	sb.WriteString(`<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>`)
	sb.WriteString(xfrmXML(sh))
	geom := sh.Geom
	if geom == "" || geom == "custom" {
		geom = "rect"
	}
	fmt.Fprintf(sb, `<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>`, esc(geom))
	if sh.Line != nil {
		sb.WriteString(lineXML("a:ln", sh.Line))
	}
	sb.WriteString(`</p:spPr></p:pic>`)
	return nil
}

var dataURLRe = regexp.MustCompile(`^data:([a-zA-Z0-9.+/-]+)?(;[^,]*)?,`)

// imageRel returns a relationship id for an image src (media: ref or data URL).
func (sw *slideWriter) imageRel(src string) (string, error) {
	part, err := sw.w.mediaPart(src)
	if err != nil {
		return "", err
	}
	return sw.addRel(rtImage, part, false), nil
}

func (w *writer) mediaPart(src string) (string, error) {
	if strings.HasPrefix(src, "media:") {
		p := strings.TrimPrefix(src, "media:")
		if name, ok := w.out.findFile(p); ok {
			return name, nil
		}
		if d, ok := w.model.Media[p]; ok && strings.HasPrefix(d, "data:") {
			return w.mediaPart(d)
		}
		return "", fmt.Errorf("image %s not found (send the original .pptx or a data: URL)", p)
	}
	m := dataURLRe.FindStringSubmatch(src)
	if m == nil {
		return "", fmt.Errorf("unsupported image source (expected data: URL or media: reference)")
	}
	payload := src[len(m[0]):]
	var data []byte
	if strings.Contains(m[2], "base64") {
		b, err := base64.StdEncoding.DecodeString(strings.TrimSpace(payload))
		if err != nil {
			b, err = base64.RawStdEncoding.DecodeString(strings.TrimRight(strings.TrimSpace(payload), "="))
			if err != nil {
				return "", fmt.Errorf("invalid base64 image data")
			}
		}
		data = b
	} else {
		data = []byte(payload)
	}
	sum := sha1.Sum(data)
	key := hex.EncodeToString(sum[:])
	if p, ok := w.media[key]; ok {
		return p, nil
	}
	ext := "png"
	switch strings.ToLower(m[1]) {
	case "image/jpeg", "image/jpg":
		ext = "jpeg"
	case "image/gif":
		ext = "gif"
	case "image/bmp":
		ext = "bmp"
	case "image/svg+xml":
		ext = "svg"
	case "image/tiff":
		ext = "tiff"
	}
	n := 1
	var part string
	for {
		part = fmt.Sprintf("ppt/media/image%d.%s", n, ext)
		if _, ok := w.out.findFile(part); !ok {
			break
		}
		n++
	}
	w.out.put(part, data)
	w.media[key] = part
	return part, nil
}

// --- notes ---------------------------------------------------------------

func (w *writer) writeNotes(part, slidePart string, s *Slide, so *slideOrig) error {
	nm, err := w.ensureNotesMaster()
	if err != nil {
		return err
	}
	rels := []Rel{{ID: "rId1", Type: rtNotesMaster, Target: nm}, {ID: "rId2", Type: rtSlide, Target: slidePart}}
	if so != nil && so.notesPart != "" && so.notes == s.Notes {
		if data, ok := w.rd.pkg.Files[so.notesPart]; ok {
			orels := w.rd.pkg.Rels(so.notesPart)
			var out []Rel
			for _, id := range sortedRelIDs(orels) {
				r := orels[id]
				switch {
				case sameRelKind(r.Type, rtSlide):
					r.Target = slidePart
				case r.External:
				default:
					if _, ok := w.out.Files[r.Target]; !ok {
						continue
					}
				}
				out = append(out, r)
			}
			w.out.put(part, data)
			w.out.put(relsPartName(part), writeRels(out, part))
			w.newCT[part] = ctNotesSlide
			return nil
		}
	}
	var sb strings.Builder
	sb.WriteString(xmlDecl)
	sb.WriteString(`<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`)
	sb.WriteString(`<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`)
	sb.WriteString(`<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>`)
	for _, line := range strings.Split(strings.ReplaceAll(s.Notes, "\r\n", "\n"), "\n") {
		if line == "" {
			sb.WriteString(`<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>`)
		} else {
			sb.WriteString(`<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>` + escText(line) + `</a:t></a:r></a:p>`)
		}
	}
	sb.WriteString(`</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`)
	w.out.put(part, []byte(sb.String()))
	w.out.put(relsPartName(part), writeRels(rels, part))
	w.newCT[part] = ctNotesSlide
	return nil
}

func (w *writer) presRels() map[string]Rel {
	root := w.out.XML(relsPartName(w.rd.presPart))
	out := map[string]Rel{}
	if root == nil {
		return out
	}
	base := path.Dir(w.rd.presPart)
	for _, e := range root.Elems() {
		r := Rel{ID: e.A("Id"), Type: e.A("Type"), Target: e.A("Target")}
		if strings.EqualFold(e.A("TargetMode"), "External") {
			r.External = true
		} else {
			r.Target = resolveTarget(base, r.Target)
		}
		out[r.ID] = r
	}
	return out
}

func (w *writer) ensureNotesMaster() (string, error) {
	for _, r := range w.presRels() {
		if sameRelKind(r.Type, rtNotesMaster) {
			if _, ok := w.out.Files[r.Target]; ok {
				return r.Target, nil
			}
		}
	}
	// add the template notes master and its theme
	n := 1
	for {
		if _, ok := w.out.Files[fmt.Sprintf("ppt/theme/theme%d.xml", n)]; !ok {
			break
		}
		n++
	}
	theme := fmt.Sprintf("ppt/theme/theme%d.xml", n)
	nm := "ppt/notesMasters/notesMaster1.xml"
	w.out.put(theme, []byte(tplTheme))
	w.out.put(nm, []byte(tplNotesMaster))
	w.out.put(relsPartName(nm), writeRels([]Rel{{ID: "rId1", Type: rtTheme, Target: theme}}, nm))
	w.newCT[theme] = ctTheme
	w.newCT[nm] = ctNotesMaster
	w.addNotesMaster = nm
	return nm, nil
}

// --- presentation.xml ------------------------------------------------------

func prefixFor(root *Node, uri string) string {
	for k, v := range root.inScopeNS() {
		if v == uri && k != "" {
			return k
		}
	}
	return ""
}

func (w *writer) writePresentation() error {
	presPart := w.rd.presPart
	root, err := parseXML(w.out.Files[presPart])
	if err != nil {
		return err
	}
	pp := root.Prefix
	rp := prefixFor(root, nsR)
	if rp == "" {
		rp = "r"
		root.NS = append(root.NS, [2]string{"r", nsR})
	}
	el := func(local string, attrs ...Attr) *Node {
		n := &Node{Prefix: pp, Local: local, URI: nsP, Attrs: attrs}
		return n
	}
	// relationships: keep everything except slides
	rels := w.presRels()
	var keep []Rel
	maxID := 0
	for _, id := range sortedRelIDs(rels) {
		r := rels[id]
		if v, err := strconv.Atoi(strings.TrimPrefix(r.ID, "rId")); err == nil && v > maxID {
			maxID = v
		}
		if sameRelKind(r.Type, rtSlide) {
			continue
		}
		keep = append(keep, r)
	}
	next := func() string {
		maxID++
		return "rId" + strconv.Itoa(maxID)
	}
	if w.addNotesMaster != "" {
		id := next()
		keep = append(keep, Rel{ID: id, Type: rtNotesMaster, Target: w.addNotesMaster})
		lst := el("notesMasterIdLst")
		lst.Children = []*Node{el("notesMasterId", Attr{Prefix: rp, Local: "id", URI: nsR, Value: id})}
		insertAfter(root, lst, "sldMasterIdLst")
	}
	sldLst := root.C("p:sldIdLst")
	if sldLst == nil {
		sldLst = el("sldIdLst")
		insertBefore(root, sldLst, "sldSz", "notesSz")
	}
	sldLst.Children = nil
	for i, part := range w.slideRefs {
		id := next()
		keep = append(keep, Rel{ID: id, Type: rtSlide, Target: part})
		sid := el("sldId", Attr{Local: "id", Value: strconv.Itoa(256 + i)}, Attr{Prefix: rp, Local: "id", URI: nsR, Value: id})
		sid.Parent = sldLst
		sldLst.Children = append(sldLst.Children, sid)
	}
	if len(w.slideRefs) == 0 {
		root.RemoveChild(sldLst)
	}
	if sz := root.C("p:sldSz"); sz != nil && w.model.Size.CX > 0 && w.model.Size.CY > 0 {
		if atoi64(sz.A("cx")) != w.model.Size.CX || atoi64(sz.A("cy")) != w.model.Size.CY {
			sz.SetA("cx", strconv.FormatInt(w.model.Size.CX, 10))
			sz.SetA("cy", strconv.FormatInt(w.model.Size.CY, 10))
			sz.DelA("type")
		}
	}
	if c := root.C("p:custShowLst"); c != nil {
		root.RemoveChild(c)
	}
	if ext := root.C("p:extLst"); ext != nil {
		for _, e := range ext.All("p:ext") {
			if e.Find("p:sldIdLst") != nil || strings.Contains(e.Serialize(), "sectionLst") {
				ext.RemoveChild(e)
			}
		}
		if len(ext.Elems()) == 0 {
			root.RemoveChild(ext)
		}
	}
	w.out.put(presPart, root.Document())
	w.out.put(relsPartName(presPart), writeRels(keep, presPart))
	return nil
}

func insertAfter(root, n *Node, after string) {
	n.Parent = root
	for i, c := range root.Children {
		if c.Local == after && c.URI == nsP {
			root.Children = append(root.Children[:i+1], append([]*Node{n}, root.Children[i+1:]...)...)
			return
		}
	}
	root.Children = append([]*Node{n}, root.Children...)
}

func insertBefore(root, n *Node, before ...string) {
	n.Parent = root
	for i, c := range root.Children {
		for _, b := range before {
			if c.Local == b && c.URI == nsP {
				root.Children = append(root.Children[:i], append([]*Node{n}, root.Children[i:]...)...)
				return
			}
		}
	}
	root.Children = append(root.Children, n)
}

// --- [Content_Types].xml ---------------------------------------------------

func (w *writer) writeContentTypes(removed map[string]bool) {
	root, err := parseXML(w.out.Files["[Content_Types].xml"])
	if err != nil || root == nil {
		root, _ = parseXML([]byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`))
	}
	defaults := map[string]bool{}
	var kids []*Node
	for _, c := range root.Elems() {
		switch c.Local {
		case "Default":
			defaults[strings.ToLower(c.A("Extension"))] = true
			kids = append(kids, c)
		case "Override":
			pn := strings.TrimPrefix(c.A("PartName"), "/")
			if removed[pn] {
				if _, still := w.out.Files[pn]; !still {
					continue
				}
			}
			if _, ok := w.newCT[pn]; ok {
				continue
			}
			if _, ok := w.out.Files[pn]; !ok {
				continue
			}
			kids = append(kids, c)
		}
	}
	mk := func(local string, attrs ...[2]string) *Node {
		n := &Node{Local: local, URI: nsCT, Parent: root}
		for _, a := range attrs {
			n.Attrs = append(n.Attrs, Attr{Local: a[0], Value: a[1]})
		}
		return n
	}
	exts := map[string]bool{}
	for name := range w.out.Files {
		if strings.HasPrefix(name, "ppt/media/") {
			exts[strings.ToLower(strings.TrimPrefix(path.Ext(name), "."))] = true
		}
	}
	var extList []string
	for e := range exts {
		extList = append(extList, e)
	}
	sort.Strings(extList)
	var newDefaults []*Node
	for _, e := range extList {
		if !defaults[e] && mimeForExt(e) != "" {
			newDefaults = append(newDefaults, mk("Default", [2]string{"Extension", e}, [2]string{"ContentType", mimeForExt(e)}))
		}
	}
	var parts []string
	for p := range w.newCT {
		parts = append(parts, p)
	}
	sort.Slice(parts, func(i, j int) bool { return natLess(parts[i], parts[j]) })
	var overrides []*Node
	for _, p := range parts {
		overrides = append(overrides, mk("Override", [2]string{"PartName", "/" + p}, [2]string{"ContentType", w.newCT[p]}))
	}
	// Defaults must come before Overrides is not required, but keep it tidy.
	var ds, os []*Node
	for _, k := range kids {
		if k.Local == "Default" {
			ds = append(ds, k)
		} else {
			os = append(os, k)
		}
	}
	root.Children = append(append(append(append([]*Node{}, ds...), newDefaults...), os...), overrides...)
	w.out.put("[Content_Types].xml", root.Document())
}

var digitsRe = regexp.MustCompile(`\d+`)

func natLess(a, b string) bool {
	pa := digitsRe.ReplaceAllStringFunc(a, func(s string) string { return fmt.Sprintf("%08s", s) })
	pb := digitsRe.ReplaceAllStringFunc(b, func(s string) string { return fmt.Sprintf("%08s", s) })
	return pa < pb
}

var (
	appSlidesRe = regexp.MustCompile(`<Slides>\d*</Slides>`)
	appNotesRe  = regexp.MustCompile(`<Notes>\d*</Notes>`)
	appHiddenRe = regexp.MustCompile(`<HiddenSlides>\d*</HiddenSlides>`)
)

func (w *writer) writeAppProps() {
	name := "docProps/app.xml"
	data, ok := w.out.Files[name]
	if !ok {
		return
	}
	notes, hidden := 0, 0
	for _, s := range w.model.Slides {
		if strings.TrimSpace(s.Notes) != "" {
			notes++
		}
		if s.Hidden {
			hidden++
		}
	}
	s := string(data)
	s = appSlidesRe.ReplaceAllString(s, fmt.Sprintf("<Slides>%d</Slides>", len(w.model.Slides)))
	s = appNotesRe.ReplaceAllString(s, fmt.Sprintf("<Notes>%d</Notes>", notes))
	s = appHiddenRe.ReplaceAllString(s, fmt.Sprintf("<HiddenSlides>%d</HiddenSlides>", hidden))
	w.out.put(name, []byte(s))
}
