package pptx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// custGeomPath converts a:custGeom into an SVG path "d" string in
// shape-local EMU coordinates (0..w, 0..h). Returns "" when the geometry
// uses formulas we do not evaluate.
func custGeomPath(cg *Node, w, h int64) string {
	pl := cg.C("a:pathLst")
	if pl == nil {
		return ""
	}
	var sb strings.Builder
	num := func(s string) (float64, bool) {
		v, err := strconv.ParseFloat(s, 64)
		return v, err == nil
	}
	f := func(v float64) string { return strconv.FormatFloat(math.Round(v), 'f', -1, 64) }
	for _, p := range pl.All("a:path") {
		pw, ph := atof(p.A("w")), atof(p.A("h"))
		sx, sy := 1.0, 1.0
		if pw > 0 {
			sx = float64(w) / pw
		}
		if ph > 0 {
			sy = float64(h) / ph
		}
		var cx, cy float64
		pt := func(n *Node) (float64, float64, bool) {
			x, ok1 := num(n.A("x"))
			y, ok2 := num(n.A("y"))
			return x, y, ok1 && ok2
		}
		for _, c := range p.Elems() {
			switch c.Local {
			case "moveTo", "lnTo":
				x, y, ok := pt(c.C("a:pt"))
				if !ok {
					return ""
				}
				cmd := "M"
				if c.Local == "lnTo" {
					cmd = "L"
				}
				fmt.Fprintf(&sb, "%s%s %s ", cmd, f(x*sx), f(y*sy))
				cx, cy = x, y
			case "cubicBezTo", "quadBezTo":
				pts := c.All("a:pt")
				cmd := "C"
				if c.Local == "quadBezTo" {
					cmd = "Q"
				}
				sb.WriteString(cmd)
				for _, q := range pts {
					x, y, ok := pt(q)
					if !ok {
						return ""
					}
					fmt.Fprintf(&sb, "%s %s ", f(x*sx), f(y*sy))
					cx, cy = x, y
				}
			case "arcTo":
				wR, ok1 := num(c.A("wR"))
				hR, ok2 := num(c.A("hR"))
				st, ok3 := num(c.A("stAng"))
				sw, ok4 := num(c.A("swAng"))
				if !(ok1 && ok2 && ok3 && ok4) {
					return ""
				}
				a1 := st / 60000 * math.Pi / 180
				a2 := (st + sw) / 60000 * math.Pi / 180
				ox := cx - wR*math.Cos(a1)
				oy := cy - hR*math.Sin(a1)
				ex := ox + wR*math.Cos(a2)
				ey := oy + hR*math.Sin(a2)
				large := 0
				if math.Abs(sw) > 180*60000 {
					large = 1
				}
				sweep := 0
				if sw > 0 {
					sweep = 1
				}
				fmt.Fprintf(&sb, "A%s %s 0 %d %d %s %s ", f(wR*sx), f(hR*sy), large, sweep, f(ex*sx), f(ey*sy))
				cx, cy = ex, ey
			case "close":
				sb.WriteString("Z ")
			}
		}
	}
	return strings.TrimSpace(sb.String())
}
