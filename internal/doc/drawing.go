package doc

import (
	"bytes"
	"encoding/base64"
	"image"
	"image/color"
	"image/png"
)

// Pictures: inline PICF structures in the Data stream and floating OfficeArt
// shapes anchored with FSPA records.

type fspa struct {
	spid                     uint32
	left, top, right, bottom int
}

type shapeInfo struct {
	typ    int
	pib    int
	txid   int32
	noLine bool
}

type bseInfo struct {
	blip  []byte // record bytes of the blip (header included), may be nil
	delay uint32
	size  uint32
}

type rec struct {
	ver, inst int
	typ       uint16
	body      []byte
	all       []byte
}

func readRecs(b []byte) []rec {
	var out []rec
	p := 0
	for p+8 <= len(b) {
		vi := le.Uint16(b[p:])
		t := le.Uint16(b[p+2:])
		n := int(le.Uint32(b[p+4:]))
		if t < 0xF000 || n < 0 {
			break
		}
		end := p + 8 + n
		if end > len(b) {
			end = len(b)
		}
		out = append(out, rec{ver: int(vi & 0xF), inst: int(vi >> 4), typ: t, body: b[p+8 : end], all: b[p:end]})
		p = end
	}
	return out
}

func (d *wordDoc) readDrawings() {
	if b := d.tblBytes(iPlcSpaMom); len(b) >= 30 {
		n := (len(b) - 4) / 30
		for i := 0; i <= n; i++ {
			d.spaCP = append(d.spaCP, i32(b, 4*i))
		}
		for i := 0; i < n; i++ {
			o := 4*(n+1) + 26*i
			d.spa = append(d.spa, fspa{spid: u32(b, o), left: i32(b, o+4), top: i32(b, o+8), right: i32(b, o+12), bottom: i32(b, o+16)})
		}
	}
	if b := d.tblBytes(iPlcftxbxTxt); len(b) >= 26 {
		n := (len(b) - 4) / 26
		for i := 0; i <= n; i++ {
			d.txbxCP = append(d.txbxCP, i32(b, 4*i))
		}
		for i := 0; i < n; i++ {
			o := 4*(n+1) + 22*i
			d.txbxLid = append(d.txbxLid, int32(i32(b, o+14)))
		}
	}
	dgg := d.tblBytes(iDggInfo)
	if dgg == nil {
		return
	}
	var walk func(rs []rec)
	walk = func(rs []rec) {
		for _, r := range rs {
			switch r.typ {
			case 0xF000, 0xF002, 0xF003:
				walk(readRecs(r.body))
			case 0xF001:
				for _, b := range readRecs(r.body) {
					if b.typ != 0xF007 {
						d.bse = append(d.bse, bseInfo{})
						continue
					}
					d.bse = append(d.bse, parseBSE(b.body))
				}
			case 0xF004:
				d.readSpContainer(r.body)
			}
		}
	}
	// OfficeArtContent: DggContainer, then (dgglbl byte, DgContainer)*.
	rs := readRecs(dgg)
	if len(rs) > 0 {
		walk(rs[:1])
		p := len(rs[0].all)
		for p < len(dgg) {
			p++ // dgglbl
			sub := readRecs(dgg[p:])
			if len(sub) == 0 {
				break
			}
			walk(sub[:1])
			p += len(sub[0].all)
		}
	}
}

func parseBSE(b []byte) bseInfo {
	if len(b) < 36 {
		return bseInfo{}
	}
	bi := bseInfo{size: u32(b, 20), delay: u32(b, 28)}
	cbName := int(b[33])
	if 36+cbName < len(b) {
		bi.blip = b[36+cbName:]
	}
	return bi
}

func (d *wordDoc) readSpContainer(b []byte) {
	var info shapeInfo
	var spid uint32
	have := false
	for _, r := range readRecs(b) {
		switch r.typ {
		case 0xF00A:
			spid = u32(r.body, 0)
			info.typ = r.inst
			have = true
		case 0xF00B:
			for i := 0; i < r.inst && 6*i+6 <= len(r.body); i++ {
				id := le.Uint16(r.body[6*i:]) & 0x3FFF
				v := le.Uint32(r.body[6*i+2:])
				switch id {
				case 0x0104:
					info.pib = int(v)
				case 0x0186:
					if info.pib == 0 {
						info.pib = int(v)
					}
				case 0x0080:
					info.txid = int32(v)
				case 0x01FF:
					// fLine is bit 3, fUsefLine bit 19
					if v&(1<<19) != 0 && v&(1<<3) == 0 {
						info.noLine = true
					}
				}
			}
		case 0xF003, 0xF004:
			d.readSpContainer(r.body)
		}
	}
	if have {
		d.shapes[spid] = &info
	}
}

// blipFromBSE returns the blip record for 1-based index pib.
func (d *wordDoc) blipFromBSE(pib int) []byte {
	if pib <= 0 || pib > len(d.bse) {
		return nil
	}
	b := d.bse[pib-1]
	if len(b.blip) >= 8 {
		return b.blip
	}
	if b.size > 0 && int64(b.delay)+int64(b.size) <= int64(len(d.wd)) {
		return d.wd[b.delay : b.delay+b.size]
	}
	return nil
}

// blipToDataURL converts an OfficeArt blip record into a browser-friendly
// data URL. ok=false for formats that cannot be shown (EMF/WMF/PICT/TIFF).
func blipToDataURL(r []byte) (url string, kind string, ok bool) {
	rs := readRecs(r)
	if len(rs) == 0 {
		return "", "", false
	}
	b := rs[0]
	twoUID := func(single int) int {
		if b.inst == single+1 {
			return 32
		}
		return 16
	}
	switch b.typ {
	case 0xF01D, 0xF02A: // JPEG
		off := 16
		if b.inst == 0x46B || b.inst == 0x6E3 {
			off = 32
		}
		data := sliceN(b.body, off+1, len(b.body))
		return dataURL("image/jpeg", data), "jpeg", len(data) > 0
	case 0xF01E: // PNG
		off := twoUID(0x6E0)
		data := sliceN(b.body, off+1, len(b.body))
		return dataURL("image/png", data), "png", len(data) > 0
	case 0xF01F: // DIB
		off := twoUID(0x7A8)
		data := sliceN(b.body, off+1, len(b.body))
		if p := dibToPNG(data); p != nil {
			return dataURL("image/png", p), "png", true
		}
		return "", "dib", false
	case 0xF01A, 0xF01B, 0xF01C:
		kinds := map[uint16]string{0xF01A: "EMF", 0xF01B: "WMF", 0xF01C: "PICT"}
		return "", kinds[b.typ], false
	case 0xF029:
		return "", "TIFF", false
	}
	return "", "", false
}

func dataURL(mime string, b []byte) string {
	return "data:" + mime + ";base64," + base64.StdEncoding.EncodeToString(b)
}

// findBlip searches a record list (recursively, through BSE records) for the
// first blip record.
func findBlip(b []byte) []byte {
	for _, r := range readRecs(b) {
		switch {
		case r.typ >= 0xF018 && r.typ <= 0xF117:
			return r.all
		case r.typ == 0xF007:
			bi := parseBSE(r.body)
			if bi.blip != nil {
				if x := findBlip(bi.blip); x != nil {
					return x
				}
			}
		case r.ver == 0xF:
			if x := findBlip(r.body); x != nil {
				return x
			}
		}
	}
	return nil
}

type picture struct {
	url  string
	kind string
	w, h int // px
	ok   bool
}

// inlinePicture decodes the PICF at offset off in the Data stream.
func (d *wordDoc) inlinePicture(off int) picture {
	b := d.data
	lcb := i32(b, off)
	cbHeader := int(u16(b, off+4))
	if lcb <= 0 || off+lcb > len(b) || cbHeader < 0x44 || cbHeader > lcb {
		return picture{}
	}
	pic := b[off : off+lcb]
	mm := i16(pic, 6)
	dxaGoal, dyaGoal := i16(pic, 28), i16(pic, 30)
	mx, my := int(u16(pic, 32)), int(u16(pic, 34))
	if mx == 0 {
		mx = 1000
	}
	if my == 0 {
		my = 1000
	}
	p := picture{w: dxaGoal * mx / 1000 / 15, h: dyaGoal * my / 1000 / 15}
	rest := pic[cbHeader:]
	if mm == 0x66 && len(rest) > 0 {
		rest = rest[1+int(rest[0]):]
	}
	if mm != 0x64 && mm != 0x66 {
		p.kind = "WMF"
		return p
	}
	blip := findBlip(rest)
	if blip == nil {
		// Some writers store a BSE whose blip lives in the WordDocument stream.
		for _, r := range readRecs(rest) {
			if r.typ == 0xF004 {
				for _, s := range readRecs(r.body) {
					if s.typ == 0xF00B {
						for i := 0; i < s.inst && 6*i+6 <= len(s.body); i++ {
							if le.Uint16(s.body[6*i:])&0x3FFF == 0x0104 {
								blip = d.blipFromBSE(int(le.Uint32(s.body[6*i+2:])))
							}
						}
					}
				}
			}
		}
	}
	if blip == nil {
		return p
	}
	p.url, p.kind, p.ok = blipToDataURL(blip)
	return p
}

// floatingPicture returns the picture of the shape anchored at FSPA index i.
func (d *wordDoc) floatingPicture(i int) picture {
	f := d.spa[i]
	p := picture{w: (f.right - f.left) / 15, h: (f.bottom - f.top) / 15}
	s := d.shapes[f.spid]
	if s == nil || s.pib == 0 {
		return p
	}
	if blip := d.blipFromBSE(s.pib); blip != nil {
		p.url, p.kind, p.ok = blipToDataURL(blip)
	}
	return p
}

// dibToPNG converts an uncompressed BITMAPINFOHEADER DIB to PNG.
func dibToPNG(b []byte) []byte {
	if len(b) < 40 {
		return nil
	}
	hdr := int(u32(b, 0))
	w, h := i32(b, 4), i32(b, 8)
	bpp := int(u16(b, 14))
	comp := u32(b, 16)
	clrUsed := int(u32(b, 32))
	if hdr < 40 || w <= 0 || w > 20000 || h == 0 || h > 20000 || h < -20000 || (comp != 0 && comp != 3) {
		return nil
	}
	topDown := h < 0
	if topDown {
		h = -h
	}
	var pal []color.RGBA
	if bpp <= 8 {
		n := clrUsed
		if n == 0 {
			n = 1 << bpp
		}
		for i := 0; i < n; i++ {
			o := hdr + 4*i
			if o+4 > len(b) {
				break
			}
			pal = append(pal, color.RGBA{b[o+2], b[o+1], b[o], 255})
		}
	}
	po := hdr + 4*len(pal)
	if comp == 3 && hdr == 40 {
		po += 12
	}
	stride := ((w*bpp + 31) / 32) * 4
	if po+stride*h > len(b) {
		return nil
	}
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	for y := 0; y < h; y++ {
		row := b[po+stride*y:]
		dy := h - 1 - y
		if topDown {
			dy = y
		}
		for x := 0; x < w; x++ {
			var c color.RGBA
			switch bpp {
			case 1, 4, 8:
				bit := x * bpp
				v := int(row[bit/8]>>(8-bpp-bit%8)) & (1<<bpp - 1)
				if v < len(pal) {
					c = pal[v]
				}
			case 24:
				c = color.RGBA{row[3*x+2], row[3*x+1], row[3*x], 255}
			case 32:
				c = color.RGBA{row[4*x+2], row[4*x+1], row[4*x], 255}
			default:
				return nil
			}
			img.SetRGBA(x, dy, c)
		}
	}
	var out bytes.Buffer
	if png.Encode(&out, img) != nil {
		return nil
	}
	return out.Bytes()
}
