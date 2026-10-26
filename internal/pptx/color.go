package pptx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// themeInfo holds the parts of a theme the resolver needs.
type themeInfo struct {
	colors    map[string]string // dk1.. -> "#RRGGBB"
	major     string
	minor     string
	fillStyle []*Node // fmtScheme/fillStyleLst children
	lnStyle   []*Node
	bgFill    []*Node
	spDef     *Node // objectDefaults/spDef
	txDef     *Node
}

func parseTheme(root *Node) *themeInfo {
	t := &themeInfo{colors: map[string]string{}, major: "Calibri Light", minor: "Calibri"}
	if root == nil {
		for k, v := range defaultThemeColors {
			t.colors[k] = v
		}
		return t
	}
	te := root.C("a:themeElements")
	if cs := te.C("a:clrScheme"); cs != nil {
		for _, c := range cs.Elems() {
			for _, v := range c.Elems() {
				switch v.Local {
				case "srgbClr":
					t.colors[c.Local] = "#" + strings.ToUpper(v.A("val"))
				case "sysClr":
					lc := v.A("lastClr")
					if lc == "" {
						lc = sysColors[v.A("val")]
					}
					if lc == "" {
						lc = "000000"
					}
					t.colors[c.Local] = "#" + strings.ToUpper(lc)
				}
			}
		}
	}
	for k, v := range defaultThemeColors {
		if _, ok := t.colors[k]; !ok {
			t.colors[k] = v
		}
	}
	if fs := te.C("a:fontScheme"); fs != nil {
		if f := fs.Path("a:majorFont", "a:latin").A("typeface"); f != "" {
			t.major = f
		}
		if f := fs.Path("a:minorFont", "a:latin").A("typeface"); f != "" {
			t.minor = f
		}
	}
	if fm := te.C("a:fmtScheme"); fm != nil {
		t.fillStyle = fm.C("a:fillStyleLst").Elems()
		t.lnStyle = fm.C("a:lnStyleLst").Elems()
		t.bgFill = fm.C("a:bgFillStyleLst").Elems()
	}
	if od := root.C("a:objectDefaults"); od != nil {
		t.spDef = od.C("a:spDef")
		t.txDef = od.C("a:txDef")
	}
	return t
}

var defaultThemeColors = map[string]string{
	"dk1": "#000000", "lt1": "#FFFFFF", "dk2": "#44546A", "lt2": "#E7E6E6",
	"accent1": "#4472C4", "accent2": "#ED7D31", "accent3": "#A5A5A5", "accent4": "#FFC000",
	"accent5": "#5B9BD5", "accent6": "#70AD47", "hlink": "#0563C1", "folHlink": "#954F72",
}

var sysColors = map[string]string{
	"windowText": "000000", "window": "FFFFFF", "btnFace": "F0F0F0", "btnText": "000000",
	"highlight": "0078D7", "highlightText": "FFFFFF", "grayText": "6D6D6D", "menu": "F0F0F0",
	"menuText": "000000", "captionText": "000000", "3dDkShadow": "696969", "3dLight": "E3E3E3",
}

var presetColors = map[string]string{
	"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
	"yellow": "FFFF00", "cyan": "00FFFF", "magenta": "FF00FF", "gray": "808080", "grey": "808080",
	"darkGray": "A9A9A9", "lightGray": "D3D3D3", "orange": "FFA500", "purple": "800080",
	"navy": "000080", "maroon": "800000", "olive": "808000", "teal": "008080", "silver": "C0C0C0",
	"lime": "00FF00", "brown": "A52A2A", "pink": "FFC0CB", "gold": "FFD700", "darkBlue": "00008B",
	"darkRed": "8B0000", "darkGreen": "006400", "ltGray": "D3D3D3", "dkGray": "A9A9A9",
}

// colorCtx resolves scheme colors for one slide/layout/master.
type colorCtx struct {
	theme  *themeInfo
	clrMap map[string]string // bg1 -> lt1 ...
}

func defaultClrMap() map[string]string {
	return map[string]string{"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
		"accent1": "accent1", "accent2": "accent2", "accent3": "accent3", "accent4": "accent4",
		"accent5": "accent5", "accent6": "accent6", "hlink": "hlink", "folHlink": "folHlink"}
}

func readClrMap(n *Node, base map[string]string) map[string]string {
	m := map[string]string{}
	for k, v := range base {
		m[k] = v
	}
	if n != nil {
		for _, a := range n.Attrs {
			m[a.Local] = a.Value
		}
	}
	return m
}

func (c *colorCtx) scheme(name string) string {
	if mapped, ok := c.clrMap[name]; ok {
		name = mapped
	}
	if v, ok := c.theme.colors[name]; ok {
		return v
	}
	switch name {
	case "tx1", "dk1", "tx2":
		return "#000000"
	}
	return "#FFFFFF"
}

func isColorElem(n *Node) bool {
	if n == nil || n.URI != nsA {
		return false
	}
	switch n.Local {
	case "srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr":
		return true
	}
	return false
}

// firstColor finds the first color element child of n.
func firstColor(n *Node) *Node {
	for _, c := range n.Elems() {
		if isColorElem(c) {
			return c
		}
	}
	return nil
}

// resolveColor turns a DrawingML color element into "#RRGGBB". phClr is the
// placeholder color substituted for schemeClr val="phClr".
func (c *colorCtx) resolveColor(e *Node, phClr string) string {
	if e == nil {
		return ""
	}
	var r, g, b float64
	hex := ""
	switch e.Local {
	case "srgbClr":
		hex = e.A("val")
	case "schemeClr":
		v := e.A("val")
		if v == "phClr" {
			hex = strings.TrimPrefix(phClr, "#")
			if hex == "" {
				hex = "000000"
			}
		} else {
			hex = strings.TrimPrefix(c.scheme(v), "#")
		}
	case "sysClr":
		hex = e.A("lastClr")
		if hex == "" {
			hex = sysColors[e.A("val")]
		}
	case "prstClr":
		hex = presetColors[e.A("val")]
	case "scrgbClr":
		r = toSRGB(pct(e.A("r")))
		g = toSRGB(pct(e.A("g")))
		b = toSRGB(pct(e.A("b")))
		hex = fmt.Sprintf("%02X%02X%02X", clamp255(r*255), clamp255(g*255), clamp255(b*255))
	case "hslClr":
		h := atof(e.A("hue")) / 60000 / 360
		s := pct(e.A("sat"))
		l := pct(e.A("lum"))
		r, g, b = hslToRGB(h, s, l)
		hex = fmt.Sprintf("%02X%02X%02X", clamp255(r*255), clamp255(g*255), clamp255(b*255))
	}
	if len(hex) != 6 {
		hex = "000000"
	}
	r, g, b = parseHex(hex)
	for _, t := range e.Elems() {
		v := pct(t.A("val"))
		switch t.Local {
		case "lumMod", "lumOff", "satMod", "satOff", "hueOff", "hueMod":
			h, s, l := rgbToHSL(r, g, b)
			switch t.Local {
			case "lumMod":
				l *= v
			case "lumOff":
				l += v
			case "satMod":
				s *= v
			case "satOff":
				s += v
			case "hueMod":
				h *= v
			case "hueOff":
				h += atof(t.A("val")) / 60000 / 360
			}
			r, g, b = hslToRGB(h, clamp01(s), clamp01(l))
		case "tint":
			r = toSRGB(toLinear(r)*v + (1 - v))
			g = toSRGB(toLinear(g)*v + (1 - v))
			b = toSRGB(toLinear(b)*v + (1 - v))
		case "shade":
			r = toSRGB(toLinear(r) * v)
			g = toSRGB(toLinear(g) * v)
			b = toSRGB(toLinear(b) * v)
		case "comp":
			h, s, l := rgbToHSL(r, g, b)
			r, g, b = hslToRGB(h+0.5, s, l)
		case "inv":
			r, g, b = 1-r, 1-g, 1-b
		case "gray":
			y := 0.299*r + 0.587*g + 0.114*b
			r, g, b = y, y, y
		}
	}
	return fmt.Sprintf("#%02X%02X%02X", clamp255(r*255), clamp255(g*255), clamp255(b*255))
}

// fillColor resolves a fill-bearing element's fill: returns (color, kind)
// where kind is "solid", "none", "image", "" (no fill info found).
func (c *colorCtx) fillOf(parent *Node, phClr string) (string, string, *Node) {
	for _, f := range parent.Elems() {
		switch f.Local {
		case "solidFill":
			return c.resolveColor(firstColor(f), phClr), "solid", f
		case "noFill":
			return "", "none", f
		case "gradFill":
			stops := f.C("a:gsLst").All("a:gs")
			if len(stops) > 0 {
				return c.resolveColor(firstColor(stops[0]), phClr), "solid", f
			}
			return "", "none", f
		case "pattFill":
			if fg := f.C("a:fgClr"); fg != nil {
				return c.resolveColor(firstColor(fg), phClr), "solid", f
			}
			return "", "none", f
		case "blipFill":
			return "", "image", f
		case "grpFill":
			return "", "group", f
		}
	}
	return "", "", nil
}

func pct(s string) float64 { return atof(s) / 100000 }

func atof(s string) float64 {
	v, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil {
		return 0
	}
	return v
}

func atoi64(s string) int64 {
	v, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil {
		return 0
	}
	return int64(math.Round(v))
}

func clamp01(v float64) float64 { return math.Max(0, math.Min(1, v)) }

func clamp255(v float64) int { return int(math.Round(math.Max(0, math.Min(255, v)))) }

func parseHex(h string) (float64, float64, float64) {
	h = strings.TrimPrefix(h, "#")
	if len(h) != 6 {
		return 0, 0, 0
	}
	v, err := strconv.ParseUint(h, 16, 32)
	if err != nil {
		return 0, 0, 0
	}
	return float64(v>>16&255) / 255, float64(v>>8&255) / 255, float64(v&255) / 255
}

func toLinear(c float64) float64 {
	if c <= 0.04045 {
		return c / 12.92
	}
	return math.Pow((c+0.055)/1.055, 2.4)
}

func toSRGB(c float64) float64 {
	c = clamp01(c)
	if c <= 0.0031308 {
		return c * 12.92
	}
	return 1.055*math.Pow(c, 1/2.4) - 0.055
}

func rgbToHSL(r, g, b float64) (float64, float64, float64) {
	mx := math.Max(r, math.Max(g, b))
	mn := math.Min(r, math.Min(g, b))
	l := (mx + mn) / 2
	if mx == mn {
		return 0, 0, l
	}
	d := mx - mn
	s := d / (1 - math.Abs(2*l-1))
	var h float64
	switch mx {
	case r:
		h = math.Mod((g-b)/d, 6)
	case g:
		h = (b-r)/d + 2
	default:
		h = (r-g)/d + 4
	}
	h /= 6
	if h < 0 {
		h++
	}
	return h, s, l
}

func hslToRGB(h, s, l float64) (float64, float64, float64) {
	h = h - math.Floor(h)
	c := (1 - math.Abs(2*l-1)) * s
	x := c * (1 - math.Abs(math.Mod(h*6, 2)-1))
	m := l - c/2
	var r, g, b float64
	switch int(h * 6) {
	case 0:
		r, g, b = c, x, 0
	case 1:
		r, g, b = x, c, 0
	case 2:
		r, g, b = 0, c, x
	case 3:
		r, g, b = 0, x, c
	case 4:
		r, g, b = x, 0, c
	default:
		r, g, b = c, 0, x
	}
	return r + m, g + m, b + m
}
