package doc

import (
	"encoding/binary"
	"errors"
	"strings"
	"unicode/utf16"
)

// Minimal reader for the Compound File Binary format (MS-CFB), enough to pull
// the streams of a Word 97-2003 document out of the container.

var cfbMagic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

const (
	secFree       = 0xFFFFFFFF
	secEndOfChain = 0xFFFFFFFE
	secFAT        = 0xFFFFFFFD
	secDIFAT      = 0xFFFFFFFC
	noStream      = 0xFFFFFFFF
)

type cfbEntry struct {
	name               string
	typ                byte // 1 storage, 2 stream, 5 root
	left, right, child uint32
	start              uint32
	size               uint64
}

type cfbFile struct {
	data       []byte
	secSize    int
	miniSize   int
	miniCutoff uint64
	fat        []uint32
	miniFAT    []uint32
	entries    []cfbEntry
	ministream []byte
}

func isCFB(b []byte) bool {
	return len(b) >= 512 && string(b[:8]) == string(cfbMagic)
}

func openCFB(b []byte) (*cfbFile, error) {
	if !isCFB(b) {
		return nil, errors.New("not a compound file")
	}
	le := binary.LittleEndian
	f := &cfbFile{data: b}
	secShift := le.Uint16(b[0x1E:])
	miniShift := le.Uint16(b[0x20:])
	if secShift != 9 && secShift != 12 {
		return nil, errors.New("corrupt compound file: bad sector size")
	}
	if miniShift == 0 || miniShift > 12 {
		miniShift = 6
	}
	f.secSize = 1 << secShift
	f.miniSize = 1 << miniShift
	f.miniCutoff = uint64(le.Uint32(b[0x38:]))
	if f.miniCutoff == 0 {
		f.miniCutoff = 4096
	}
	numFAT := le.Uint32(b[0x2C:])
	firstDir := le.Uint32(b[0x30:])
	firstMiniFAT := le.Uint32(b[0x3C:])
	firstDIFAT := le.Uint32(b[0x44:])
	numDIFAT := le.Uint32(b[0x48:])

	// Collect FAT sector numbers from the header DIFAT and DIFAT chain.
	var fatSecs []uint32
	for i := 0; i < 109; i++ {
		s := le.Uint32(b[0x4C+4*i:])
		if s == secFree || s == secEndOfChain {
			continue
		}
		fatSecs = append(fatSecs, s)
	}
	d := firstDIFAT
	for n := uint32(0); n < numDIFAT && d != secEndOfChain && d != secFree; n++ {
		sec := f.sector(d)
		if sec == nil {
			break
		}
		per := f.secSize/4 - 1
		for i := 0; i < per; i++ {
			s := le.Uint32(sec[4*i:])
			if s != secFree && s != secEndOfChain {
				fatSecs = append(fatSecs, s)
			}
		}
		d = le.Uint32(sec[4*per:])
	}
	if uint32(len(fatSecs)) > numFAT && numFAT > 0 {
		fatSecs = fatSecs[:numFAT]
	}
	for _, s := range fatSecs {
		sec := f.sector(s)
		if sec == nil {
			return nil, errors.New("corrupt compound file: FAT sector out of range")
		}
		for i := 0; i+4 <= len(sec); i += 4 {
			f.fat = append(f.fat, le.Uint32(sec[i:]))
		}
	}

	dir := f.chain(firstDir, -1)
	for i := 0; i+128 <= len(dir); i += 128 {
		e := dir[i : i+128]
		nl := int(le.Uint16(e[64:]))
		if nl > 64 {
			nl = 64
		}
		var u []uint16
		for j := 0; j+1 < nl; j += 2 {
			c := le.Uint16(e[j:])
			if c == 0 {
				break
			}
			u = append(u, c)
		}
		ent := cfbEntry{
			name:  string(utf16.Decode(u)),
			typ:   e[66],
			left:  le.Uint32(e[68:]),
			right: le.Uint32(e[72:]),
			child: le.Uint32(e[76:]),
			start: le.Uint32(e[116:]),
			size:  le.Uint64(e[120:]),
		}
		if f.secSize == 512 {
			ent.size &= 0xFFFFFFFF
		}
		f.entries = append(f.entries, ent)
	}
	if len(f.entries) == 0 || f.entries[0].typ != 5 {
		return nil, errors.New("corrupt compound file: missing root entry")
	}
	root := f.entries[0]
	f.ministream = f.chain(root.start, int64(root.size))
	mf := f.chain(firstMiniFAT, -1)
	for i := 0; i+4 <= len(mf); i += 4 {
		f.miniFAT = append(f.miniFAT, le.Uint32(mf[i:]))
	}
	return f, nil
}

func (f *cfbFile) sector(n uint32) []byte {
	off := (int64(n) + 1) * int64(f.secSize)
	if n >= secDIFAT || off < 0 || off >= int64(len(f.data)) {
		return nil
	}
	end := off + int64(f.secSize)
	if end > int64(len(f.data)) {
		end = int64(len(f.data))
	}
	return f.data[off:end]
}

// chain reads a regular-sector chain; size < 0 reads it all.
func (f *cfbFile) chain(start uint32, size int64) []byte {
	var out []byte
	seen := map[uint32]bool{}
	for s := start; s < secDIFAT && !seen[s]; {
		seen[s] = true
		sec := f.sector(s)
		if sec == nil {
			break
		}
		out = append(out, sec...)
		if size >= 0 && int64(len(out)) >= size {
			break
		}
		if int(s) >= len(f.fat) {
			break
		}
		s = f.fat[s]
	}
	if size >= 0 && int64(len(out)) > size {
		out = out[:size]
	}
	return out
}

func (f *cfbFile) miniChain(start uint32, size int64) []byte {
	var out []byte
	seen := map[uint32]bool{}
	for s := start; s < secDIFAT && !seen[s]; {
		seen[s] = true
		off := int64(s) * int64(f.miniSize)
		if off >= int64(len(f.ministream)) {
			break
		}
		end := off + int64(f.miniSize)
		if end > int64(len(f.ministream)) {
			end = int64(len(f.ministream))
		}
		out = append(out, f.ministream[off:end]...)
		if int64(len(out)) >= size || int(s) >= len(f.miniFAT) {
			break
		}
		s = f.miniFAT[s]
	}
	if int64(len(out)) > size {
		out = out[:size]
	}
	return out
}

// rootChildren lists the entries directly inside the root storage.
func (f *cfbFile) rootChildren() []cfbEntry {
	var out []cfbEntry
	seen := map[uint32]bool{}
	var walk func(i uint32)
	walk = func(i uint32) {
		if i == noStream || int(i) >= len(f.entries) || seen[i] {
			return
		}
		seen[i] = true
		e := f.entries[i]
		walk(e.left)
		out = append(out, e)
		walk(e.right)
	}
	walk(f.entries[0].child)
	return out
}

// stream returns the named top-level stream (case-insensitive), or nil.
func (f *cfbFile) stream(name string) []byte {
	for _, e := range f.rootChildren() {
		if e.typ == 2 && strings.EqualFold(e.name, name) {
			return f.read(e)
		}
	}
	return nil
}

func (f *cfbFile) has(name string) bool {
	for _, e := range f.rootChildren() {
		if strings.EqualFold(e.name, name) {
			return true
		}
	}
	return false
}

func (f *cfbFile) read(e cfbEntry) []byte {
	if e.size == 0 {
		return []byte{}
	}
	if e.size < f.miniCutoff {
		return f.miniChain(e.start, int64(e.size))
	}
	return f.chain(e.start, int64(e.size))
}
