package doc

import "fmt"

// Character and paragraph property state and sprm application.

type chp struct {
	b, i, strike, caps, smallCaps, vanish, u, del bool
	hps                                           int // half-points
	ftc                                           int
	color, highlight, shade                       string // "#rrggbb" or ""
	iss                                           int    // 1 super, 2 sub
	spacing                                       int    // twips
	fSpec, fData, fOle2                           bool
	picLoc                                        int
	symFtc                                        int
	symChar                                       uint16
	istd                                          int
}

type brc struct {
	set   bool // explicitly specified (false = inherit)
	none  bool
	width float64 // points
	color string
}

type tc struct {
	hMerge, vMerge int // h: 0 none, 1 first, 2 merged; v: 0 none, 1 restart, 2 continue
	vAlign         int
	borders        [4]brc // top, left, bottom, right
	shade          string
}

type tdef struct {
	centers []int
	cells   []tc
	gapHalf int
	tblBrc  [6]brc // top, left, bottom, right, insideH, insideV
}

type pap struct {
	istd                int
	jc                  int
	left, right, left1  int
	before, after       int
	line                int
	lineMult            bool
	hasLine             bool
	inTable, ttp        bool
	itap                int
	innerCell, innerTtp bool
	ilfo, ilvl          int
	outLvl              int
	pageBreakBefore     bool
	shade               string
	tdef                *tdef
	hasLeft, hasLeft1   bool
}

var icoColors = []string{"", "#000000", "#0000ff", "#00ffff", "#00ff00", "#ff00ff", "#ff0000", "#ffff00",
	"#ffffff", "#000080", "#008080", "#008000", "#800080", "#800000", "#808000", "#808080", "#c0c0c0"}

func ico(i int) string {
	if i > 0 && i < len(icoColors) {
		return icoColors[i]
	}
	return ""
}

// colorRef decodes a COLORREF; "" for automatic.
func colorRef(a []byte) string {
	if len(a) < 4 || a[3] == 0xFF {
		return ""
	}
	return fmt.Sprintf("#%02x%02x%02x", a[0], a[1], a[2])
}

func toggle(v byte, styleVal bool) bool {
	switch v {
	case 0:
		return false
	case 1:
		return true
	case 0x80:
		return styleVal
	case 0x81:
		return !styleVal
	}
	return v != 0
}

// applyChpx applies a character grpprl; st holds the style values that the
// 0x80/0x81 toggle operands refer to.
func (d *wordDoc) applyChpx(c *chp, g []byte, st chp) {
	eachSprm(g, func(op uint16, a []byte) {
		v := byte(argU(a))
		switch op {
		case 0x0835:
			c.b = toggle(v, st.b)
		case 0x0836:
			c.i = toggle(v, st.i)
		case 0x0837, 0x2A53:
			c.strike = toggle(v, st.strike)
		case 0x083A:
			c.smallCaps = toggle(v, st.smallCaps)
		case 0x083B:
			c.caps = toggle(v, st.caps)
		case 0x083C:
			c.vanish = toggle(v, st.vanish)
		case 0x2A3E:
			c.u = v != 0
		case 0x4A43:
			if n := int(argU(a)); n > 0 {
				c.hps = n
			}
		case 0x2A42:
			c.color = ico(int(v))
		case 0x6870:
			c.color = colorRef(a)
		case 0x2A0C:
			c.highlight = ico(int(v))
		case 0x4866:
			s := argU(a)
			c.shade = ico(int((s >> 5) & 0x1F))
		case 0xCA71:
			if len(a) >= 10 {
				c.shade = colorRef(a[4:8])
			}
		case 0x2A48:
			c.iss = int(v)
		case 0x4A4F:
			c.ftc = int(argU(a))
		case 0x8840:
			c.spacing = argI(a)
		case 0x0855:
			c.fSpec = toggle(v, false)
		case 0x0806:
			c.fData = toggle(v, false)
		case 0x080A:
			c.fOle2 = toggle(v, false)
		case 0x0800:
			c.del = toggle(v, false)
		case 0x6A03:
			c.picLoc = int(argU(a))
		case 0x6A09:
			if len(a) >= 4 {
				c.symFtc = int(le.Uint16(a))
				c.symChar = le.Uint16(a[2:])
			}
		case 0x2A33: // sprmCPlain: back to the style's properties
			keep := *c
			*c = st
			c.fSpec, c.picLoc, c.symChar = keep.fSpec, keep.picLoc, keep.symChar
		}
	})
}

func readBrc80(a []byte) brc {
	if len(a) < 4 {
		return brc{}
	}
	switch le.Uint32(a) {
	case 0:
		return brc{} // unspecified: inherit the table border
	case 0xFFFFFFFF:
		return brc{set: true, none: true} // explicitly nil
	}
	b := brc{set: true}
	if a[1] == 0 || a[1] == 0xFF {
		b.none = true
		return b
	}
	b.width = float64(a[0]) / 8
	b.color = ico(int(a[2]))
	if b.color == "" {
		b.color = "#000000"
	}
	return b
}

func readBrc(a []byte) brc {
	if len(a) < 8 {
		return brc{}
	}
	if le.Uint32(a[4:]) == 0xFFFFFFFF && le.Uint32(a) == 0xFFFFFFFF {
		return brc{}
	}
	b := brc{set: true}
	if a[5] == 0 || a[5] == 0xFF {
		b.none = true
		return b
	}
	b.width = float64(a[4]) / 8
	b.color = colorRef(a[:4])
	if b.color == "" {
		b.color = "#000000"
	}
	return b
}

func (d *wordDoc) applyPapx(p *pap, g []byte) {
	eachSprm(g, func(op uint16, a []byte) {
		switch op {
		case 0x2403, 0x2461:
			p.jc = int(argU(a))
		case 0x840F, 0x845E:
			p.left, p.hasLeft = argI(a), true
		case 0x840E, 0x845D:
			p.right = argI(a)
		case 0x8411, 0x8460:
			p.left1, p.hasLeft1 = argI(a), true
		case 0xA413:
			p.before = int(argU(a))
		case 0xA414:
			p.after = int(argU(a))
		case 0x6412:
			if len(a) >= 4 {
				p.line = int(int16(le.Uint16(a)))
				p.lineMult = le.Uint16(a[2:]) != 0
				p.hasLine = true
			}
		case 0x2416:
			p.inTable = argU(a) != 0
			if p.inTable && p.itap == 0 {
				p.itap = 1
			}
		case 0x2417:
			p.ttp = argU(a) != 0
		case 0x6649:
			p.itap = int(argU(a))
		case 0x664A:
			p.itap += argI(a)
		case 0x244B:
			p.innerCell = argU(a) != 0
		case 0x244C:
			p.innerTtp = argU(a) != 0
		case 0x260A:
			p.ilvl = int(argU(a))
		case 0x460B:
			p.ilfo = int(argU(a))
		case 0x2640:
			p.outLvl = int(argU(a))
		case 0x2407:
			p.pageBreakBefore = argU(a) != 0
		case 0x442D:
			s := argU(a)
			p.shade = ico(int((s >> 5) & 0x1F))
		case 0xC64D:
			if len(a) >= 10 {
				p.shade = colorRef(a[4:8])
			}
		case 0xD608:
			p.tdef = d.readTDef(a, p.tdef)
		default:
			if p.tdef != nil || op>>10&7 == 5 {
				d.applyTableSprm(p, op, a)
			}
		}
	})
}

func (d *wordDoc) readTDef(a []byte, old *tdef) *tdef {
	t := &tdef{}
	if old != nil {
		t.gapHalf, t.tblBrc = old.gapHalf, old.tblBrc
	}
	if len(a) < 1 {
		return t
	}
	n := int(a[0])
	for i := 0; i <= n; i++ {
		t.centers = append(t.centers, i16(a, 1+2*i))
	}
	base := 1 + 2*(n+1)
	for i := 0; i < n; i++ {
		c := tc{}
		o := base + 20*i
		if o+20 <= len(a) {
			g := u16(a, o)
			switch g & 3 {
			case 1:
				c.hMerge = 1
			case 2, 3:
				c.hMerge = 2
			}
			switch (g >> 5) & 3 {
			case 3:
				c.vMerge = 1
			case 1, 2:
				c.vMerge = 2
			}
			c.vAlign = int((g >> 7) & 3)
			for k := 0; k < 4; k++ {
				c.borders[k] = readBrc80(a[o+4+4*k:])
			}
		}
		t.cells = append(t.cells, c)
	}
	return t
}

func (d *wordDoc) applyTableSprm(p *pap, op uint16, a []byte) {
	if p.tdef == nil {
		p.tdef = &tdef{}
	}
	t := p.tdef
	cellRange := func(first, lim int, fn func(c *tc)) {
		for i := first; i < lim && i < len(t.cells); i++ {
			fn(&t.cells[i])
		}
	}
	switch op {
	case 0x9602:
		t.gapHalf = argI(a)
	case 0xD605:
		for k := 0; k < 6 && 4*k+4 <= len(a); k++ {
			t.tblBrc[k] = readBrc80(a[4*k:])
		}
	case 0xD613:
		for k := 0; k < 6 && 8*k+8 <= len(a); k++ {
			t.tblBrc[k] = readBrc(a[8*k:])
		}
	case 0xD609:
		for i := 0; 2*i+2 <= len(a) && i < len(t.cells); i++ {
			s := u16(a, 2*i)
			t.cells[i].shade = ico(int((s >> 5) & 0x1F))
		}
	case 0xD612:
		for i := 0; 10*i+10 <= len(a) && i < len(t.cells); i++ {
			t.cells[i].shade = colorRef(a[10*i+4:])
		}
	case 0xD62C:
		if len(a) >= 3 {
			cellRange(int(a[0]), int(a[1]), func(c *tc) { c.vAlign = int(a[2]) })
		}
	case 0xD620, 0xD62F:
		if len(a) >= 3 {
			var b brc
			if op == 0xD620 {
				b = readBrc80(a[3:])
			} else {
				b = readBrc(a[3:])
			}
			m := a[2]
			cellRange(int(a[0]), int(a[1]), func(c *tc) {
				for k := 0; k < 4; k++ {
					if m&(1<<k) != 0 {
						c.borders[k] = b
					}
				}
			})
		}
	}
}
