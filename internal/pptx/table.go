package pptx

import (
	"encoding/json"
	"fmt"
	"strings"
)

// readLine overlays an a:ln-like element onto base. Returns nil for no line.
func readLine(ln *Node, cc *colorCtx, phClr string, base *Line) *Line {
	if ln == nil {
		return base
	}
	var out *Line
	if base != nil {
		c := *base
		out = &c
	}
	col, kind, _ := cc.fillOf(ln, phClr)
	switch kind {
	case "none":
		return nil
	case "solid":
		if out == nil {
			out = &Line{Width: 9525, Dash: "solid"}
		}
		out.Color = col
	}
	if out == nil {
		return nil
	}
	if w, ok := ln.AttrOK("w"); ok {
		out.Width = atoi64(w)
	}
	if d := ln.C("a:prstDash"); d != nil && d.A("val") != "" {
		out.Dash = d.A("val")
	}
	if h := ln.C("a:headEnd"); h != nil {
		out.HeadEnd = noneEmpty(h.A("type"))
	}
	if t := ln.C("a:tailEnd"); t != nil {
		out.TailEnd = noneEmpty(t.A("type"))
	}
	return out
}

func noneEmpty(s string) string {
	if s == "none" {
		return ""
	}
	return s
}

// lineFromRef resolves an a:lnRef element against the theme.
func lineFromRef(ref *Node, cc *colorCtx) *Line {
	if ref == nil {
		return nil
	}
	idx := int(atoi64(ref.A("idx")))
	if idx <= 0 {
		return nil
	}
	col := cc.resolveColor(firstColor(ref), "")
	l := &Line{Color: col, Width: 9525, Dash: "solid"}
	if idx-1 < len(cc.theme.lnStyle) {
		th := cc.theme.lnStyle[idx-1]
		if w, ok := th.AttrOK("w"); ok {
			l.Width = atoi64(w)
		}
		if _, kind, _ := cc.fillOf(th, col); kind == "none" {
			return nil
		}
		if d := th.C("a:prstDash"); d != nil && d.A("val") != "" {
			l.Dash = d.A("val")
		}
	}
	return l
}

// fillFromRef resolves an a:fillRef / a:bgRef color (ignores gradients).
func fillFromRef(ref *Node, cc *colorCtx) (*string, bool) {
	if ref == nil {
		return nil, false
	}
	idx := int(atoi64(ref.A("idx")))
	if idx == 0 {
		return nil, true
	}
	col := cc.resolveColor(firstColor(ref), "")
	var entry *Node
	if idx >= 1001 && idx-1001 < len(cc.theme.bgFill) {
		entry = cc.theme.bgFill[idx-1001]
	} else if idx >= 1 && idx-1 < len(cc.theme.fillStyle) {
		entry = cc.theme.fillStyle[idx-1]
	}
	if entry != nil {
		c, kind, _ := cc.fillOf(&Node{Children: []*Node{entry}}, col)
		switch kind {
		case "none":
			return nil, true
		case "solid":
			return &c, true
		}
	}
	return &col, true
}

type cellStyle struct {
	fill    *string
	borders Borders
	text    rProps
}

func linePtrEq(a, b *Line) bool {
	if a == nil || b == nil {
		return a == nil && b == nil
	}
	return *a == *b
}

func strPtrEq(a, b *string) bool {
	if a == nil || b == nil {
		return a == nil && b == nil
	}
	return strings.EqualFold(*a, *b)
}

// tableCellStyle resolves the table style for a cell position.
func tableCellStyle(style *Node, ts *TableStyle, row, col, nRows, nCols int, cc *colorCtx) cellStyle {
	var cs cellStyle
	if style == nil || ts == nil {
		return cs
	}
	lastRow, lastCol := nRows-1, nCols-1
	type region struct {
		name string
		kind int // 0 whole, 1 row region, 2 column region
	}
	var regs []region
	regs = append(regs, region{"wholeTbl", 0})
	if ts.BandRow {
		ri := row
		if ts.FirstRow {
			ri--
		}
		inHeader := ts.FirstRow && row == 0
		inFooter := ts.LastRow && row == lastRow
		if !inHeader && !inFooter && ri >= 0 {
			if ri%2 == 0 {
				regs = append(regs, region{"band1H", 1})
			} else {
				regs = append(regs, region{"band2H", 1})
			}
		}
	}
	if ts.BandCol {
		ci := col
		if ts.FirstCol {
			ci--
		}
		inFirst := ts.FirstCol && col == 0
		inLast := ts.LastCol && col == lastCol
		if !inFirst && !inLast && ci >= 0 {
			if ci%2 == 0 {
				regs = append(regs, region{"band1V", 2})
			} else {
				regs = append(regs, region{"band2V", 2})
			}
		}
	}
	if ts.LastCol && col == lastCol {
		regs = append(regs, region{"lastCol", 2})
	}
	if ts.FirstCol && col == 0 {
		regs = append(regs, region{"firstCol", 2})
	}
	if ts.LastRow && row == lastRow {
		regs = append(regs, region{"lastRow", 1})
	}
	if ts.FirstRow && row == 0 {
		regs = append(regs, region{"firstRow", 1})
	}
	borderOf := func(e *Node) (*Line, bool) {
		if e == nil {
			return nil, false
		}
		if ln := e.C("a:ln"); ln != nil {
			return readLine(ln, cc, "", nil), true
		}
		if ref := e.C("a:lnRef"); ref != nil {
			return lineFromRef(ref, cc), true
		}
		return nil, false
	}
	for _, rg := range regs {
		part := style.C("a:" + rg.name)
		if part == nil {
			continue
		}
		if tx := part.C("a:tcTxStyle"); tx != nil {
			if v, ok := tx.AttrOK("b"); ok {
				cs.text.B = boolp(v == "on")
			}
			if v, ok := tx.AttrOK("i"); ok {
				cs.text.I = boolp(v == "on")
			}
			if c := firstColor(tx); c != nil {
				cs.text.Color = strp(cc.resolveColor(c, ""))
			}
		}
		st := part.C("a:tcStyle")
		if st == nil {
			continue
		}
		if f := st.C("a:fill"); f != nil {
			c, kind, _ := cc.fillOf(f, "")
			switch kind {
			case "none":
				cs.fill = nil
			case "solid":
				cs.fill = strp(c)
			}
		} else if ref := st.C("a:fillRef"); ref != nil {
			if f, ok := fillFromRef(ref, cc); ok {
				cs.fill = f
			}
		}
		bd := st.C("a:tcBdr")
		if bd == nil {
			continue
		}
		pick := func(edge bool, outer, inner string) (*Line, bool) {
			if edge {
				return borderOf(bd.C("a:" + outer))
			}
			return borderOf(bd.C("a:" + inner))
		}
		var l, r, t, b *Line
		var okL, okR, okT, okB bool
		switch rg.kind {
		case 0:
			l, okL = pick(col == 0, "left", "insideV")
			r, okR = pick(col == lastCol, "right", "insideV")
			t, okT = pick(row == 0, "top", "insideH")
			b, okB = pick(row == lastRow, "bottom", "insideH")
		case 1:
			l, okL = pick(col == 0, "left", "insideV")
			r, okR = pick(col == lastCol, "right", "insideV")
			t, okT = borderOf(bd.C("a:top"))
			b, okB = borderOf(bd.C("a:bottom"))
		case 2:
			l, okL = borderOf(bd.C("a:left"))
			r, okR = borderOf(bd.C("a:right"))
			t, okT = pick(row == 0, "top", "insideH")
			b, okB = pick(row == lastRow, "bottom", "insideH")
		}
		if okL {
			cs.borders.L = l
		}
		if okR {
			cs.borders.R = r
		}
		if okT {
			cs.borders.T = t
		}
		if okB {
			cs.borders.B = b
		}
	}
	return cs
}

func (rd *reader) tableStyleNode(id string) *Node {
	if id == "" {
		return nil
	}
	return rd.tblStyles[strings.ToUpper(id)]
}

// parseTable converts a:tbl into a table shape (geometry already set).
func (rd *reader) parseTable(sh *Shape, tbl *Node, sc *sctx) {
	sh.Kind = "table"
	tp := tbl.C("a:tblPr")
	ts := &TableStyle{
		FirstRow: parseBool(tp.A("firstRow")), FirstCol: parseBool(tp.A("firstCol")),
		LastRow: parseBool(tp.A("lastRow")), LastCol: parseBool(tp.A("lastCol")),
		BandRow: parseBool(tp.A("bandRow")), BandCol: parseBool(tp.A("bandCol")),
	}
	if id := tp.C("a:tableStyleId"); id != nil {
		ts.ID = strings.TrimSpace(id.TextContent())
	} else if tp != nil && tp.C("a:tableStyle") == nil {
		ts.ID = ""
	}
	sh.TableStyle = ts
	style := rd.tableStyleNode(ts.ID)
	for _, gc := range tbl.C("a:tblGrid").All("a:gridCol") {
		sh.ColW = append(sh.ColW, atoi64(gc.A("w")))
	}
	trs := tbl.All("a:tr")
	nRows, nCols := len(trs), len(sh.ColW)
	base := newChain(sc.cc, layer{list: rd.defText})
	for ri, tr := range trs {
		sh.RowH = append(sh.RowH, atoi64(tr.A("h")))
		var row []*Cell
		for ci, tc := range tr.All("a:tc") {
			cs := tableCellStyle(style, ts, ri, ci, nRows, nCols, sc.cc)
			cell := rd.readCell(tc, cs, base, sc)
			row = append(row, cell)
		}
		for len(row) < nCols {
			row = append(row, &Cell{Paras: []*Para{}, Span: 1, RowSpan: 1, HMerge: true})
		}
		sh.Rows = append(sh.Rows, row)
	}
}

func (rd *reader) readCell(tc *Node, cs cellStyle, base *textChain, sc *sctx) *Cell {
	cell := &Cell{Span: 1, RowSpan: 1, Anchor: "t", Inset: []int64{91440, 45720, 91440, 45720}}
	if v := tc.A("gridSpan"); v != "" {
		cell.Span = int(atoi64(v))
	}
	if v := tc.A("rowSpan"); v != "" {
		cell.RowSpan = int(atoi64(v))
	}
	cell.HMerge = parseBool(tc.A("hMerge"))
	cell.VMerge = parseBool(tc.A("vMerge"))
	cell.Fill = cs.fill
	b := cs.borders
	pr := tc.C("a:tcPr")
	if pr != nil {
		ins := []string{"marL", "marT", "marR", "marB"}
		for i, k := range ins {
			if v, ok := pr.AttrOK(k); ok {
				cell.Inset[i] = atoi64(v)
			}
		}
		if v := pr.A("anchor"); v != "" {
			cell.Anchor = v
		}
		b.L = readLine(pr.C("a:lnL"), sc.cc, "", b.L)
		b.R = readLine(pr.C("a:lnR"), sc.cc, "", b.R)
		b.T = readLine(pr.C("a:lnT"), sc.cc, "", b.T)
		b.B = readLine(pr.C("a:lnB"), sc.cc, "", b.B)
		if c, kind, _ := sc.cc.fillOf(pr, ""); kind == "solid" {
			cell.Fill = strp(c)
		} else if kind == "none" {
			cell.Fill = nil
		}
	}
	if b.L != nil || b.R != nil || b.T != nil || b.B != nil {
		bb := b
		cell.Borders = &bb
	}
	tr := cs.text
	chain := base.with(layer{r: &tr})
	cell.Paras = readParas(tc.C("a:txBody"), chain, sc.rels)
	if cell.Paras == nil {
		cell.Paras = []*Para{}
	}
	return cell
}

// writeTable writes a:graphicFrame for a table shape.
func (w *slideWriter) writeTable(sb *strings.Builder, sh *Shape, id int) {
	fmt.Fprintf(sb, `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="%s"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr>%s</p:nvPr></p:nvGraphicFramePr>`, id, esc(nameOr(sh.Name, fmt.Sprintf("Table %d", id-1))), w.phXML(sh, w.phParentsOf(sh)))
	fmt.Fprintf(sb, `<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>`, sh.X, sh.Y, sh.W, sh.H)
	sb.WriteString(`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>`)
	ts := sh.TableStyle
	if ts == nil {
		ts = &TableStyle{}
	}
	attr := func(name string, v bool) string {
		if v {
			return " " + name + `="1"`
		}
		return ""
	}
	fmt.Fprintf(sb, `<a:tblPr%s%s%s%s%s%s>`, attr("firstRow", ts.FirstRow), attr("firstCol", ts.FirstCol), attr("lastRow", ts.LastRow), attr("lastCol", ts.LastCol), attr("bandRow", ts.BandRow), attr("bandCol", ts.BandCol))
	if ts.ID != "" {
		sb.WriteString(`<a:tableStyleId>` + escText(ts.ID) + `</a:tableStyleId>`)
	}
	sb.WriteString(`</a:tblPr><a:tblGrid>`)
	nCols := len(sh.ColW)
	for _, r := range sh.Rows {
		if len(r) > nCols {
			nCols = len(r)
		}
	}
	for c := 0; c < nCols; c++ {
		cw := int64(914400)
		if c < len(sh.ColW) && sh.ColW[c] > 0 {
			cw = sh.ColW[c]
		}
		fmt.Fprintf(sb, `<a:gridCol w="%d"/>`, cw)
	}
	sb.WriteString(`</a:tblGrid>`)
	style := w.rd.tableStyleNode(ts.ID)
	base := newChain(w.cc, layer{list: w.rd.defText})
	nRows := len(sh.Rows)
	for ri, row := range sh.Rows {
		h := int64(370840)
		if ri < len(sh.RowH) && sh.RowH[ri] >= 0 {
			h = sh.RowH[ri]
		}
		fmt.Fprintf(sb, `<a:tr h="%d">`, h)
		for ci := 0; ci < nCols; ci++ {
			var cell *Cell
			if ci < len(row) {
				cell = row[ci]
			}
			if cell == nil {
				cell = &Cell{Span: 1, RowSpan: 1}
			}
			cs := tableCellStyle(style, ts, ri, ci, nRows, nCols, w.cc)
			sb.WriteString(`<a:tc`)
			if cell.Span > 1 {
				fmt.Fprintf(sb, ` gridSpan="%d"`, cell.Span)
			}
			if cell.RowSpan > 1 {
				fmt.Fprintf(sb, ` rowSpan="%d"`, cell.RowSpan)
			}
			if cell.HMerge {
				sb.WriteString(` hMerge="1"`)
			}
			if cell.VMerge {
				sb.WriteString(` vMerge="1"`)
			}
			sb.WriteString(`><a:txBody><a:bodyPr/><a:lstStyle/>`)
			tr := cs.text
			writeParas(sb, cell.Paras, base.with(layer{r: &tr}), w)
			sb.WriteString(`</a:txBody>`)
			w.writeTcPr(sb, cell, cs)
			sb.WriteString(`</a:tc>`)
		}
		sb.WriteString(`</a:tr>`)
	}
	sb.WriteString(`</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`)
}

func (w *slideWriter) writeTcPr(sb *strings.Builder, cell *Cell, cs cellStyle) {
	var attrs, kids strings.Builder
	def := []int64{91440, 45720, 91440, 45720}
	names := []string{"marL", "marT", "marR", "marB"}
	for i := range def {
		if i < len(cell.Inset) && cell.Inset[i] != def[i] {
			fmt.Fprintf(&attrs, ` %s="%d"`, names[i], cell.Inset[i])
		}
	}
	if cell.Anchor != "" && cell.Anchor != "t" {
		fmt.Fprintf(&attrs, ` anchor="%s"`, esc(cell.Anchor))
	}
	var mb Borders
	if cell.Borders != nil {
		mb = *cell.Borders
	}
	edge := func(tag string, m, s *Line) {
		if linePtrEq(m, s) {
			return
		}
		if m == nil {
			fmt.Fprintf(&kids, `<a:%s><a:noFill/></a:%s>`, tag, tag)
			return
		}
		kids.WriteString(lineXML("a:"+tag, m))
	}
	edge("lnL", mb.L, cs.borders.L)
	edge("lnR", mb.R, cs.borders.R)
	edge("lnT", mb.T, cs.borders.T)
	edge("lnB", mb.B, cs.borders.B)
	if !strPtrEq(cell.Fill, cs.fill) {
		if cell.Fill == nil {
			kids.WriteString(`<a:noFill/>`)
		} else {
			fmt.Fprintf(&kids, `<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, hexVal(*cell.Fill))
		}
	}
	sb.WriteString(`<a:tcPr` + attrs.String())
	if kids.Len() > 0 {
		sb.WriteString(">" + kids.String() + "</a:tcPr>")
	} else {
		sb.WriteString("/>")
	}
}

func lineXML(tag string, l *Line) string {
	if l == nil {
		return "<" + tag + "><a:noFill/></" + tag + ">"
	}
	var sb strings.Builder
	w := l.Width
	if w <= 0 {
		w = 9525
	}
	fmt.Fprintf(&sb, `<%s w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, tag, w, hexVal(l.Color))
	if l.Dash != "" && l.Dash != "solid" {
		fmt.Fprintf(&sb, `<a:prstDash val="%s"/>`, esc(l.Dash))
	}
	if l.HeadEnd != "" {
		fmt.Fprintf(&sb, `<a:headEnd type="%s"/>`, esc(l.HeadEnd))
	}
	if l.TailEnd != "" {
		fmt.Fprintf(&sb, `<a:tailEnd type="%s"/>`, esc(l.TailEnd))
	}
	sb.WriteString("</" + tag + ">")
	return sb.String()
}

func jsonEq(a, b any) bool {
	x, _ := json.Marshal(a)
	y, _ := json.Marshal(b)
	return string(x) == string(y)
}
