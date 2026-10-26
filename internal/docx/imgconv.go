package docx

import (
	"bytes"
	"encoding/binary"
	"image"
	"image/color"
	"image/png"
	"strings"
)

// WebImage converts embedded picture data to something a browser (and
// FromHTML) can display. PNG, JPEG and GIF pass through; BMP/DIB is decoded to
// PNG; EMF and WMF metafiles are searched for their largest embedded bitmap
// (most metafiles in documents are wrapped scans/screenshots), which is
// converted to PNG. ok is false when nothing displayable was found.
func WebImage(data []byte, ext string) (mime string, out []byte, ok bool) {
	ext = strings.ToLower(strings.TrimPrefix(ext, "."))
	switch {
	case len(data) >= 8 && bytes.HasPrefix(data, []byte("\x89PNG")):
		return "image/png", data, true
	case len(data) >= 3 && data[0] == 0xFF && data[1] == 0xD8:
		return "image/jpeg", data, true
	case len(data) >= 6 && bytes.HasPrefix(data, []byte("GIF8")):
		return "image/gif", data, true
	case len(data) >= 2 && data[0] == 'B' && data[1] == 'M':
		if len(data) > 14 {
			if b, ok := dibToPNG(data[14:]); ok {
				return "image/png", b, true
			}
		}
		return "", nil, false
	case len(data) >= 44 && binary.LittleEndian.Uint32(data) == 1 && string(data[40:44]) == " EMF":
		return metaBitmap(emfBitmaps(data))
	case len(data) >= 4 && (binary.LittleEndian.Uint32(data) == 0x9AC6CDD7 || ext == "wmf"):
		return metaBitmap(wmfBitmaps(data))
	case ext == "svg" || bytes.Contains(data[:min(len(data), 256)], []byte("<svg")):
		return "image/svg+xml", data, true
	case ext == "webp":
		return "image/webp", data, true
	}
	return "", nil, false
}

func metaBitmap(dibs [][]byte) (string, []byte, bool) {
	var best []byte
	for _, d := range dibs {
		if len(d) > len(best) {
			best = d
		}
	}
	if best == nil {
		return "", nil, false
	}
	if b, ok := dibToPNG(best); ok {
		return "image/png", b, true
	}
	return "", nil, false
}

func u32(b []byte, off int) int {
	if off < 0 || off+4 > len(b) {
		return -1
	}
	return int(binary.LittleEndian.Uint32(b[off:]))
}

// emfBitmaps returns the DIBs (BITMAPINFO + bits) found in an EMF.
func emfBitmaps(data []byte) [][]byte {
	var out [][]byte
	pos := 0
	for pos+8 <= len(data) {
		typ, size := u32(data, pos), u32(data, pos+4)
		if size < 8 || pos+size > len(data) {
			break
		}
		rec := data[pos : pos+size]
		boff := -1
		switch typ {
		case 81, 80: // EMR_STRETCHDIBITS, EMR_SETDIBITSTODEVICE
			boff = 48
		case 76, 77: // EMR_BITBLT, EMR_STRETCHBLT
			boff = 84
		}
		if boff > 0 {
			offBmi, cbBmi, offBits, cbBits := u32(rec, boff), u32(rec, boff+4), u32(rec, boff+8), u32(rec, boff+12)
			if offBmi > 0 && cbBmi > 0 && offBits > 0 && cbBits > 0 && offBmi+cbBmi <= len(rec) && offBits+cbBits <= len(rec) {
				d := append(append([]byte{}, rec[offBmi:offBmi+cbBmi]...), rec[offBits:offBits+cbBits]...)
				out = append(out, d)
			}
		}
		if typ == 14 { // EMR_EOF
			break
		}
		pos += size
	}
	return out
}

// wmfBitmaps returns DIBs found in a WMF (optionally with a placeable header).
func wmfBitmaps(data []byte) [][]byte {
	pos := 0
	if u32(data, 0) == 0x9AC6CDD7 {
		pos = 22
	}
	if pos+18 > len(data) {
		return nil
	}
	hdrWords := int(binary.LittleEndian.Uint16(data[pos+2:]))
	pos += hdrWords * 2
	var out [][]byte
	for pos+6 <= len(data) {
		size := u32(data, pos) * 2
		if size < 6 || pos+size > len(data) {
			break
		}
		fn := binary.LittleEndian.Uint16(data[pos+4:])
		params := data[pos+6 : pos+size]
		skip := -1
		switch fn {
		case 0x0F43: // META_STRETCHDIB
			skip = 22
		case 0x0B41: // META_DIBSTRETCHBLT
			skip = 20
		case 0x0940: // META_DIBBITBLT
			skip = 16
		case 0x0000:
			return out
		}
		if skip > 0 && len(params) > skip+40 {
			out = append(out, params[skip:])
		}
		pos += size
	}
	return out
}

// dibToPNG decodes a packed DIB (BITMAPINFOHEADER, optional palette/masks,
// pixel bits) into PNG bytes. Supports 1/4/8/16/24/32 bpp, BI_RGB and
// BI_BITFIELDS, and embedded JPEG/PNG (BI_JPEG/BI_PNG).
func dibToPNG(d []byte) ([]byte, bool) {
	hs := u32(d, 0)
	if hs < 12 || hs > len(d) {
		return nil, false
	}
	var w, h, bpp, comp, clrUsed int
	if hs == 12 { // BITMAPCOREHEADER
		w = int(binary.LittleEndian.Uint16(d[4:]))
		h = int(int16(binary.LittleEndian.Uint16(d[6:])))
		bpp = int(binary.LittleEndian.Uint16(d[10:]))
	} else {
		if hs < 40 || len(d) < 40 {
			return nil, false
		}
		w = int(int32(binary.LittleEndian.Uint32(d[4:])))
		h = int(int32(binary.LittleEndian.Uint32(d[8:])))
		bpp = int(binary.LittleEndian.Uint16(d[14:]))
		comp = u32(d, 16)
		clrUsed = u32(d, 32)
	}
	if comp == 4 || comp == 5 { // BI_JPEG / BI_PNG
		bits := d[hs:]
		if comp == 5 {
			return bits, true
		}
		if img, _, err := image.Decode(bytes.NewReader(bits)); err == nil {
			var buf bytes.Buffer
			if png.Encode(&buf, img) == nil {
				return buf.Bytes(), true
			}
		}
		return nil, false
	}
	topDown := h < 0
	if topDown {
		h = -h
	}
	if w <= 0 || h <= 0 || w > 20000 || h > 20000 || int64(w)*int64(h) > 80e6 {
		return nil, false
	}
	pos := hs
	masks := [3]uint32{0x7C00, 0x03E0, 0x001F}
	if bpp == 32 {
		masks = [3]uint32{0xFF0000, 0xFF00, 0xFF}
	}
	if comp == 3 { // BI_BITFIELDS
		if hs == 40 {
			if len(d) < pos+12 {
				return nil, false
			}
			for i := range masks {
				masks[i] = binary.LittleEndian.Uint32(d[pos+4*i:])
			}
			pos += 12
		} else if hs >= 52 {
			for i := range masks {
				masks[i] = binary.LittleEndian.Uint32(d[40+4*i:])
			}
		}
	} else if comp != 0 {
		return nil, false // RLE not supported
	}
	var pal []color.RGBA
	if bpp <= 8 {
		n := clrUsed
		if n == 0 || n > 1<<bpp {
			n = 1 << bpp
		}
		entry := 4
		if hs == 12 {
			entry = 3
		}
		if len(d) < pos+n*entry {
			return nil, false
		}
		for i := 0; i < n; i++ {
			p := d[pos+i*entry:]
			pal = append(pal, color.RGBA{p[2], p[1], p[0], 255})
		}
		pos += n * entry
	}
	stride := ((w*bpp + 31) / 32) * 4
	if len(d) < pos+stride*h {
		if len(d) < pos+stride*(h-1)+(w*bpp+7)/8 {
			return nil, false
		}
	}
	bits := d[pos:]
	img := image.NewNRGBA(image.Rect(0, 0, w, h))
	alphaSeen := false
	shift := func(m uint32) (uint, uint32) {
		if m == 0 {
			return 0, 0
		}
		s := uint(0)
		for m&1 == 0 {
			m >>= 1
			s++
		}
		return s, m
	}
	var sh [3]uint
	var mx [3]uint32
	for i := range masks {
		sh[i], mx[i] = shift(masks[i])
	}
	scale := func(v, m uint32) uint8 {
		if m == 0 {
			return 0
		}
		return uint8(v * 255 / m)
	}
	for y := 0; y < h; y++ {
		row := y
		if !topDown {
			row = h - 1 - y
		}
		rb := bits[min(row*stride, len(bits)):]
		for x := 0; x < w; x++ {
			var c color.NRGBA
			switch bpp {
			case 1, 2, 4, 8:
				bit := x * bpp
				if bit/8 >= len(rb) {
					continue
				}
				idx := int(rb[bit/8]>>(8-bpp-bit%8)) & (1<<bpp - 1)
				if idx < len(pal) {
					p := pal[idx]
					c = color.NRGBA{p.R, p.G, p.B, 255}
				}
			case 16:
				if 2*x+2 > len(rb) {
					continue
				}
				v := uint32(binary.LittleEndian.Uint16(rb[2*x:]))
				c = color.NRGBA{scale(v>>sh[0]&mx[0], mx[0]), scale(v>>sh[1]&mx[1], mx[1]), scale(v>>sh[2]&mx[2], mx[2]), 255}
			case 24:
				if 3*x+3 > len(rb) {
					continue
				}
				c = color.NRGBA{rb[3*x+2], rb[3*x+1], rb[3*x], 255}
			case 32:
				if 4*x+4 > len(rb) {
					continue
				}
				v := binary.LittleEndian.Uint32(rb[4*x:])
				a := uint8(v >> 24)
				if a != 0 {
					alphaSeen = true
				}
				c = color.NRGBA{scale(v>>sh[0]&mx[0], mx[0]), scale(v>>sh[1]&mx[1], mx[1]), scale(v>>sh[2]&mx[2], mx[2]), a}
			default:
				return nil, false
			}
			img.SetNRGBA(x, y, c)
		}
	}
	if bpp == 32 && !alphaSeen { // alpha channel unused: opaque
		for i := 3; i < len(img.Pix); i += 4 {
			img.Pix[i] = 255
		}
	}
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		return nil, false
	}
	return buf.Bytes(), true
}
