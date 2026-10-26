package doc

import (
	"encoding/binary"
	"errors"
	"sort"
	"unicode/utf16"
)

// Parsing of the binary Word 97-2003 format (MS-DOC): FIB, piece table,
// formatted disk pages (character/paragraph properties), stylesheet, fonts,
// lists, sections, footnotes and text boxes.

var le = binary.LittleEndian

// FibRgFcLcb97 indexes used here.
const (
	iStshf       = 1
	iPlcffndRef  = 2
	iPlcffndTxt  = 3
	iPlcfSed     = 6
	iPlcfHdd     = 11
	iPlcfBteChpx = 12
	iPlcfBtePapx = 13
	iSttbfFfn    = 15
	iDop         = 31
	iClx         = 33
	iPlcSpaMom   = 40
	iPlcfendRef  = 46
	iPlcfendTxt  = 47
	iDggInfo     = 50
	iPlcftxbxTxt = 56
	iPlfLst      = 73
	iPlfLfo      = 74
)

type piece struct {
	cpStart, cpEnd int
	fc             uint32 // byte offset of the first character in WordDocument
	compressed     bool
	prm            uint16
}

type fkpRun struct {
	fcStart, fcEnd uint32
	istd           int // paragraph runs only
	grpprl         []byte
}

type style struct {
	name     string
	sti      int
	stk      int // 1 para, 2 char, 3 table, 4 list
	base     int
	papx     []byte
	chpx     []byte
	resolved bool
	busy     bool
	pap      pap
	chp      chp
}

type lvl struct {
	start int
	nfc   int
	text  []uint16 // number text; values < 9 are level placeholders
	papx  []byte
}

type lst struct {
	lsid   int32
	simple bool
	lvls   []lvl
}

type section struct {
	cpEnd int
	sep   sep
}

type sep struct {
	bkc                      int
	pageW, pageH             int
	left, right, top, bottom int
	hdr, ftr                 int
	landscape                bool
}

type wordDoc struct {
	wd, tbl, data []byte
	flags         uint16
	nFib          uint16
	ccp           [7]int // text, ftn, hdd, (mcr), atn, edn, txbx  (index 3 unused)
	ccpHdrTxbx    int
	fcLcb         [][2]uint32

	pieces  []piece
	prcs    [][]byte
	chpRuns []fkpRun
	papRuns []fkpRun

	fonts    []string
	styles   []*style
	defFtc   int
	lists    []*lst
	lfos     []int32 // lsid per 1-based ilfo-1
	sections []section
	dxaTab   int

	// drawing objects
	spaCP   []int
	spa     []fspa
	shapes  map[uint32]*shapeInfo
	bse     []bseInfo
	txbxCP  []int
	txbxLid []int32
}

func (d *wordDoc) fc(i int) (uint32, uint32) {
	if i < len(d.fcLcb) {
		return d.fcLcb[i][0], d.fcLcb[i][1]
	}
	return 0, 0
}

// tblBytes returns the table-stream slice for FibRgFcLcb entry i.
func (d *wordDoc) tblBytes(i int) []byte {
	fc, lcb := d.fc(i)
	if lcb == 0 || uint64(fc)+uint64(lcb) > uint64(len(d.tbl)) {
		return nil
	}
	return d.tbl[fc : fc+lcb]
}

func u16(b []byte, off int) uint16 {
	if off < 0 || off+2 > len(b) {
		return 0
	}
	return le.Uint16(b[off:])
}

func u32(b []byte, off int) uint32 {
	if off < 0 || off+4 > len(b) {
		return 0
	}
	return le.Uint32(b[off:])
}

func i16(b []byte, off int) int { return int(int16(u16(b, off))) }
func i32(b []byte, off int) int { return int(int32(u32(b, off))) }

var errEncrypted = errors.New("this .doc file is password-protected (encrypted); remove the password in Word and try again")

func parseWord(cf *cfbFile) (*wordDoc, error) {
	wd := cf.stream("WordDocument")
	if wd == nil {
		return nil, errors.New("not a Word document")
	}
	if len(wd) < 0x44 || u16(wd, 0) != 0xA5EC {
		if len(wd) >= 2 && u16(wd, 0) == 0xA5DC {
			return nil, errors.New("Word 6.0/95 documents are not supported; open the file in Word and save it as .docx")
		}
		return nil, errors.New("corrupt Word document: bad FIB signature")
	}
	d := &wordDoc{wd: wd, shapes: map[uint32]*shapeInfo{}}
	d.nFib = u16(wd, 2)
	d.flags = u16(wd, 0x0A)
	if d.nFib < 0xC0 {
		// nFib 101..105 are Word 6.0/95 files, which use a different layout.
		return nil, errors.New("Word 6.0/95 documents are not supported; open the file in Word and save it as .docx")
	}
	if d.flags&0x0100 != 0 || d.flags&0x8000 != 0 {
		return nil, errEncrypted
	}
	if cf.has("EncryptionInfo") {
		return nil, errEncrypted
	}
	tblName := "0Table"
	if d.flags&0x0200 != 0 {
		tblName = "1Table"
	}
	d.tbl = cf.stream(tblName)
	if d.tbl == nil {
		return nil, errors.New("corrupt Word document: missing " + tblName + " stream")
	}
	d.data = cf.stream("Data")

	// FIB variable part.
	pos := 32
	csw := int(u16(wd, pos))
	pos += 2 + 2*csw
	cslw := int(u16(wd, pos))
	lw := pos + 2
	lwAt := func(i int) int {
		if i >= cslw {
			return 0
		}
		return i32(wd, lw+4*i)
	}
	d.ccp[0] = lwAt(3)
	d.ccp[1] = lwAt(4)
	d.ccp[2] = lwAt(5)
	d.ccp[3] = 0
	d.ccp[4] = lwAt(7)
	d.ccp[5] = lwAt(8)
	d.ccp[6] = lwAt(9)
	d.ccpHdrTxbx = lwAt(10)
	pos = lw + 4*cslw
	cbRg := int(u16(wd, pos))
	pos += 2
	for i := 0; i < cbRg && pos+8 <= len(wd); i++ {
		d.fcLcb = append(d.fcLcb, [2]uint32{u32(wd, pos), u32(wd, pos+4)})
		pos += 8
	}
	for i := range d.ccp {
		if d.ccp[i] < 0 {
			d.ccp[i] = 0
		}
	}

	if err := d.readClx(); err != nil {
		return nil, err
	}
	d.chpRuns = d.readBte(iPlcfBteChpx, false)
	d.papRuns = d.readBte(iPlcfBtePapx, true)
	d.readFonts()
	d.readStyles()
	d.readLists()
	d.readSections()
	d.readDop()
	d.readDrawings()
	return d, nil
}

func (d *wordDoc) readClx() error {
	clx := d.tblBytes(iClx)
	if clx == nil {
		return errors.New("corrupt Word document: missing piece table")
	}
	p := 0
	for p < len(clx) {
		switch clx[p] {
		case 1:
			cb := int(u16(clx, p+1))
			if p+3+cb > len(clx) {
				return errors.New("corrupt Word document: bad piece table")
			}
			d.prcs = append(d.prcs, clx[p+3:p+3+cb])
			p += 3 + cb
		case 2:
			lcb := int(u32(clx, p+1))
			plc := clx[p+5:]
			if lcb > len(plc) {
				lcb = len(plc)
			}
			plc = plc[:lcb]
			n := (lcb - 4) / 12
			if n <= 0 {
				return errors.New("corrupt Word document: empty piece table")
			}
			for i := 0; i < n; i++ {
				cs, ce := i32(plc, 4*i), i32(plc, 4*(i+1))
				pcd := plc[4*(n+1)+8*i:]
				fcv := u32(pcd, 2)
				pc := piece{cpStart: cs, cpEnd: ce, prm: u16(pcd, 6)}
				if fcv&0x40000000 != 0 {
					pc.compressed = true
					pc.fc = (fcv & 0x3FFFFFFF) / 2
				} else {
					pc.fc = fcv & 0x3FFFFFFF
				}
				if ce > cs {
					d.pieces = append(d.pieces, pc)
				}
			}
			return nil
		default:
			return errors.New("corrupt Word document: bad piece table")
		}
	}
	return errors.New("corrupt Word document: no piece table")
}

func (d *wordDoc) readBte(idx int, para bool) []fkpRun {
	plc := d.tblBytes(idx)
	if len(plc) < 12 {
		return nil
	}
	n := (len(plc) - 4) / 8
	var runs []fkpRun
	seen := map[uint32]bool{}
	for i := 0; i < n; i++ {
		pn := u32(plc, 4*(n+1)+4*i) & 0x3FFFFF
		if seen[pn] {
			continue
		}
		seen[pn] = true
		off := int(pn) * 512
		if off+512 > len(d.wd) {
			continue
		}
		page := d.wd[off : off+512]
		cnt := int(page[511])
		if 4*(cnt+1) > 511 {
			continue
		}
		for j := 0; j < cnt; j++ {
			r := fkpRun{fcStart: u32(page, 4*j), fcEnd: u32(page, 4*(j+1))}
			if para {
				bx := 4*(cnt+1) + 13*j
				if bx >= 511 {
					break
				}
				o := int(page[bx]) * 2
				if o > 0 && o < 511 {
					cb := int(page[o])
					var g []byte
					if cb != 0 {
						g = sliceN(page, o+1, 2*cb-1)
					} else if o+1 < 511 {
						g = sliceN(page, o+2, 2*int(page[o+1]))
					}
					if len(g) >= 2 {
						r.istd = int(u16(g, 0))
						r.grpprl = g[2:]
					}
					r.grpprl = d.hugePapx(r.grpprl)
				}
			} else {
				o := int(page[4*(cnt+1)+j]) * 2
				if o > 0 && o < 511 {
					r.grpprl = sliceN(page, o+1, int(page[o]))
				}
			}
			runs = append(runs, r)
		}
	}
	sort.SliceStable(runs, func(a, b int) bool { return runs[a].fcStart < runs[b].fcStart })
	return runs
}

// hugePapx replaces a sprmPHugePapx reference by the grpprl in the Data stream.
func (d *wordDoc) hugePapx(g []byte) []byte {
	if len(g) >= 6 && u16(g, 0) == 0x6646 {
		off := int(u32(g, 2))
		cb := int(u16(d.data, off))
		if b := sliceN(d.data, off+2, cb); b != nil {
			return b
		}
	}
	return g
}

func sliceN(b []byte, off, n int) []byte {
	if off < 0 || n < 0 || off > len(b) {
		return nil
	}
	if off+n > len(b) {
		n = len(b) - off
	}
	return b[off : off+n]
}

func findRun(runs []fkpRun, fc uint32) int {
	i := sort.Search(len(runs), func(i int) bool { return runs[i].fcEnd > fc })
	if i < len(runs) && runs[i].fcStart <= fc {
		return i
	}
	return -1
}

func utf16z(b []byte, off, maxChars int) (string, int) {
	var u []uint16
	p := off
	for i := 0; i < maxChars && p+2 <= len(b); i++ {
		c := le.Uint16(b[p:])
		p += 2
		if c == 0 {
			break
		}
		u = append(u, c)
	}
	return string(utf16.Decode(u)), p
}

func (d *wordDoc) readFonts() {
	b := d.tblBytes(iSttbfFfn)
	if len(b) < 4 {
		return
	}
	p := 0
	cData := int(u16(b, 0))
	extended := false
	if cData == 0xFFFF {
		extended = true
		cData = int(u16(b, 2))
		p = 2
	}
	p += 4 // cData + cbExtra
	for i := 0; i < cData && p < len(b); i++ {
		var cb int
		if extended {
			cb = int(u16(b, p)) * 2
			p += 2
		} else {
			cb = int(b[p])
			p++
		}
		ffn := sliceN(b, p, cb)
		p += cb
		name := ""
		if len(ffn) > 39 {
			name, _ = utf16z(ffn, 39, 65)
		}
		d.fonts = append(d.fonts, name)
	}
}

func (d *wordDoc) fontName(ftc int) string {
	if ftc >= 0 && ftc < len(d.fonts) {
		return d.fonts[ftc]
	}
	return ""
}

func (d *wordDoc) readStyles() {
	b := d.tblBytes(iStshf)
	if len(b) < 4 {
		return
	}
	cbStshi := int(u16(b, 0))
	shi := sliceN(b, 2, cbStshi)
	cstd := int(u16(shi, 0))
	cbBase := int(u16(shi, 2))
	if cbBase == 0 {
		cbBase = 10
	}
	d.defFtc = int(u16(shi, 12))
	p := 2 + cbStshi
	for i := 0; i < cstd && p+2 <= len(b); i++ {
		cb := int(u16(b, p))
		p += 2
		std := sliceN(b, p, cb)
		p += cb
		if cb == 0 || len(std) < 10 {
			d.styles = append(d.styles, nil)
			continue
		}
		s := &style{
			sti:  int(u16(std, 0) & 0x0FFF),
			stk:  int(u16(std, 2) & 0x000F),
			base: int(u16(std, 2) >> 4),
		}
		cupx := int(u16(std, 4) & 0x000F)
		q := cbBase
		cch := int(u16(std, q))
		s.name, _ = utf16z(std, q+2, cch)
		q += 2 + 2*cch + 2
		for k := 0; k < cupx && q+2 <= len(std); k++ {
			cbUpx := int(u16(std, q))
			q += 2
			upx := sliceN(std, q, cbUpx)
			q += cbUpx
			if cbUpx%2 == 1 {
				q++
			}
			switch {
			case s.stk == 1 && k == 0, s.stk == 4 && k == 0:
				if len(upx) >= 2 {
					s.papx = upx[2:]
				}
			case s.stk == 1 && k == 1, s.stk == 2 && k == 0:
				s.chpx = upx
			case s.stk == 3 && k == 1:
				if len(upx) >= 2 {
					s.papx = upx[2:]
				}
			case s.stk == 3 && k == 2:
				s.chpx = upx
			}
		}
		d.styles = append(d.styles, s)
	}
}

func (d *wordDoc) style(istd int) *style {
	if istd >= 0 && istd < len(d.styles) {
		return d.styles[istd]
	}
	return nil
}

func (d *wordDoc) defaultChp() chp {
	return chp{hps: 20, ftc: d.defFtc, picLoc: -1}
}

// resolveStyle computes the effective paragraph/character props of a style,
// following the istdBase chain.
func (d *wordDoc) resolveStyle(istd int) (pap, chp) {
	s := d.style(istd)
	if s == nil {
		if istd != 0 {
			if s0 := d.style(0); s0 != nil {
				return d.resolveStyle(0)
			}
		}
		return pap{istd: istd}, d.defaultChp()
	}
	if s.resolved {
		return s.pap, s.chp
	}
	if s.busy {
		return pap{istd: istd}, d.defaultChp()
	}
	s.busy = true
	bp, bc := pap{}, d.defaultChp()
	if s.base != 0x0FFF && s.base != istd {
		bp, bc = d.resolveStyle(s.base)
	}
	bp.istd = istd
	d.applyPapx(&bp, s.papx)
	d.applyChpx(&bc, s.chpx, bc)
	bc.fSpec, bc.picLoc = false, -1
	s.pap, s.chp = bp, bc
	s.resolved, s.busy = true, false
	return bp, bc
}

func (d *wordDoc) readLists() {
	if lb := d.tblBytes(iPlfLst); len(lb) >= 2 {
		// The LVL records follow the PlfLst directly (outside lcbPlfLst).
		fc, _ := d.fc(iPlfLst)
		b := d.tbl[fc:]
		n := i16(b, 0)
		p := 2
		for i := 0; i < n && p+28 <= len(b); i++ {
			l := &lst{lsid: int32(i32(b, p)), simple: b[p+26]&1 != 0}
			d.lists = append(d.lists, l)
			p += 28
		}
		for _, l := range d.lists {
			cnt := 9
			if l.simple {
				cnt = 1
			}
			for k := 0; k < cnt && p+28 <= len(b); k++ {
				v := lvl{start: i32(b, p), nfc: int(b[p+4])}
				cbChpx, cbPapx := int(b[p+24]), int(b[p+25])
				p += 28
				v.papx = sliceN(b, p, cbPapx)
				p += cbPapx + cbChpx
				cch := int(u16(b, p))
				p += 2
				for c := 0; c < cch && p+2 <= len(b); c++ {
					v.text = append(v.text, le.Uint16(b[p:]))
					p += 2
				}
				l.lvls = append(l.lvls, v)
			}
		}
	}
	f := d.tblBytes(iPlfLfo)
	if len(f) >= 4 {
		n := i32(f, 0)
		for i := 0; i < n && 4+16*(i+1) <= len(f); i++ {
			d.lfos = append(d.lfos, int32(i32(f, 4+16*i)))
		}
	}
}

func (d *wordDoc) listFor(ilfo int) *lst {
	if ilfo <= 0 || ilfo > len(d.lfos) {
		return nil
	}
	id := d.lfos[ilfo-1]
	for _, l := range d.lists {
		if l.lsid == id {
			return l
		}
	}
	return nil
}

func defaultSep() sep {
	return sep{bkc: 2, pageW: 12240, pageH: 15840, left: 1800, right: 1800, top: 1440, bottom: 1440, hdr: 720, ftr: 720}
}

func (d *wordDoc) readSections() {
	b := d.tblBytes(iPlcfSed)
	if len(b) < 16 {
		return
	}
	n := (len(b) - 4) / 16
	for i := 0; i < n; i++ {
		s := section{cpEnd: i32(b, 4*(i+1)), sep: defaultSep()}
		fcSepx := u32(b, 4*(n+1)+12*i+2)
		if fcSepx != 0xFFFFFFFF {
			cb := int(u16(d.wd, int(fcSepx)))
			g := sliceN(d.wd, int(fcSepx)+2, cb)
			eachSprm(g, func(op uint16, a []byte) {
				switch op {
				case 0x3009:
					s.sep.bkc = int(argU(a))
				case 0xB01F:
					s.sep.pageW = int(argU(a))
				case 0xB020:
					s.sep.pageH = int(argU(a))
				case 0xB021:
					s.sep.left = int(argU(a))
				case 0xB022:
					s.sep.right = int(argU(a))
				case 0x9023:
					s.sep.top = argI(a)
				case 0x9024:
					s.sep.bottom = argI(a)
				case 0xB017:
					s.sep.hdr = int(argU(a))
				case 0xB018:
					s.sep.ftr = int(argU(a))
				case 0x301D:
					s.sep.landscape = argU(a) == 2
				}
			})
		}
		d.sections = append(d.sections, s)
	}
}

func (d *wordDoc) readDop() {
	d.dxaTab = 720
	b := d.tblBytes(iDop)
	if v := int(u16(b, 10)); v > 0 && v < 10000 {
		d.dxaTab = v
	}
}

// sectionBreakAt reports whether a 0x0C at cp ends a section.
func (d *wordDoc) sectionBreakAt(cp int) (bool, int) {
	for i, s := range d.sections {
		if s.cpEnd == cp+1 {
			bkc := 2
			if i+1 < len(d.sections) {
				bkc = d.sections[i+1].sep.bkc
			}
			return true, bkc
		}
	}
	return false, 0
}

// ---- text ----

type dchar struct {
	r     uint16
	fc    uint32
	piece int
}

var cp1252 = [32]uint16{
	0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D, 0x8F,
	0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D, 0x017E, 0x0178,
}

// chars returns the characters in [cpStart, cpEnd).
func (d *wordDoc) chars(cpStart, cpEnd int) []dchar {
	var out []dchar
	for pi, p := range d.pieces {
		if p.cpEnd <= cpStart || p.cpStart >= cpEnd {
			continue
		}
		s, e := p.cpStart, p.cpEnd
		if s < cpStart {
			s = cpStart
		}
		if e > cpEnd {
			e = cpEnd
		}
		for cp := s; cp < e; cp++ {
			k := cp - p.cpStart
			if p.compressed {
				fc := p.fc + uint32(k)
				if int(fc) >= len(d.wd) {
					break
				}
				c := uint16(d.wd[fc])
				if c >= 0x80 && c < 0xA0 {
					c = cp1252[c-0x80]
				}
				out = append(out, dchar{r: c, fc: fc, piece: pi})
			} else {
				fc := p.fc + uint32(2*k)
				if int(fc)+2 > len(d.wd) {
					break
				}
				out = append(out, dchar{r: le.Uint16(d.wd[fc:]), fc: fc, piece: pi})
			}
		}
	}
	return out
}

// pieceGrpprl returns the sprms attached to a piece through its Prm.
func (d *wordDoc) pieceGrpprl(pi int) []byte {
	if pi < 0 || pi >= len(d.pieces) {
		return nil
	}
	prm := d.pieces[pi].prm
	if prm&1 != 0 {
		i := int(prm >> 1)
		if i < len(d.prcs) {
			return d.prcs[i]
		}
		return nil
	}
	return nil // Prm0 single-sprm modifiers are rare and ignored
}

// ---- sprm iteration ----

func sprmSize(op uint16, g []byte, p int) (argOff, argLen int) {
	switch op >> 13 {
	case 0, 1:
		return 0, 1
	case 2, 4, 5:
		return 0, 2
	case 3:
		return 0, 4
	case 7:
		return 0, 3
	}
	// variable
	if op == 0xD608 || op == 0xD606 {
		cb := int(u16(g, p))
		return 2, cb - 1
	}
	if p >= len(g) {
		return 0, 0
	}
	cb := int(g[p])
	if op == 0xC615 && cb == 255 {
		// sprmPChgTabs with a large operand.
		if p+1 < len(g) {
			del := int(g[p+1])
			add := 0
			if q := p + 2 + 4*del; q < len(g) {
				add = int(g[q])
			}
			return 1, 1 + 4*del + 1 + 3*add
		}
	}
	return 1, cb
}

func eachSprm(g []byte, fn func(op uint16, arg []byte)) {
	p := 0
	for p+2 <= len(g) {
		op := le.Uint16(g[p:])
		p += 2
		ao, al := sprmSize(op, g, p)
		if al < 0 {
			return
		}
		arg := sliceN(g, p+ao, al)
		fn(op, arg)
		p += ao + al
	}
}

func argU(a []byte) uint32 {
	switch len(a) {
	case 0:
		return 0
	case 1:
		return uint32(a[0])
	case 2, 3:
		return uint32(le.Uint16(a))
	default:
		return le.Uint32(a)
	}
}

func argI(a []byte) int {
	switch len(a) {
	case 1:
		return int(int8(a[0]))
	case 2, 3:
		return int(int16(le.Uint16(a)))
	case 4:
		return int(int32(le.Uint32(a)))
	}
	return 0
}
