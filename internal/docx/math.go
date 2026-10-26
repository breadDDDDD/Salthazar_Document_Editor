package docx

import (
	"html"
	"strings"
	"unicode/utf8"
)

// Office Math (OMML) is rendered as linear text with <sup>/<sub>, which keeps
// every character visible and editable (e.g. "x² + (a)/(b)" or "√(x)").

func mathAttr(n *xnode, prop, child string) (string, bool) {
	pr := n.child(prop)
	if pr == nil {
		return "", false
	}
	c := pr.child(child)
	if c == nil {
		return "", false
	}
	return c.attr("val"), true
}

// mathText returns the plain text length of a converted fragment (tags stripped).
func mathLen(h string) int { return utf8.RuneCountInString(html.UnescapeString(stripTags(h))) }

func paren(h string) string {
	if mathLen(h) <= 1 {
		return h
	}
	return "(" + h + ")"
}

func (ic *inlineConv) mathHTML(n *xnode) string {
	var sb strings.Builder
	for _, k := range n.kids {
		sb.WriteString(ic.mathNode(k))
	}
	return sb.String()
}

func (ic *inlineConv) mathNode(k *xnode) string {
	kid := func(name string) string { return ic.mathHTML(k.child(name)) }
	switch k.name {
	case "r":
		var sb strings.Builder
		for _, t := range k.kids {
			if t.name == "t" {
				sb.WriteString(html.EscapeString(t.text))
			}
		}
		return sb.String()
	case "f":
		num, den := kid("num"), kid("den")
		if t, _ := mathAttr(k, "fPr", "type"); t == "noBar" {
			return "<sup>" + num + "</sup><sub>" + den + "</sub>"
		}
		return paren(num) + "/" + paren(den)
	case "sSup":
		return kid("e") + "<sup>" + kid("sup") + "</sup>"
	case "sSub":
		return kid("e") + "<sub>" + kid("sub") + "</sub>"
	case "sSubSup":
		return kid("e") + "<sub>" + kid("sub") + "</sub><sup>" + kid("sup") + "</sup>"
	case "sPre":
		return "<sub>" + kid("sub") + "</sub><sup>" + kid("sup") + "</sup>" + kid("e")
	case "rad":
		deg := kid("deg")
		pre := ""
		if mathLen(deg) > 0 {
			pre = "<sup>" + deg + "</sup>"
		}
		return pre + "√" + paren(kid("e"))
	case "d":
		beg, ok := mathAttr(k, "dPr", "begChr")
		if !ok {
			beg = "("
		}
		end, ok := mathAttr(k, "dPr", "endChr")
		if !ok {
			end = ")"
		}
		sep, ok := mathAttr(k, "dPr", "sepChr")
		if !ok {
			sep = "|"
		}
		var parts []string
		for _, e := range k.kids {
			if e.name == "e" {
				parts = append(parts, ic.mathHTML(e))
			}
		}
		return html.EscapeString(beg) + strings.Join(parts, html.EscapeString(sep)) + html.EscapeString(end)
	case "nary":
		chr, ok := mathAttr(k, "naryPr", "chr")
		if !ok {
			chr = "∫"
		}
		out := html.EscapeString(chr)
		if s := kid("sub"); mathLen(s) > 0 {
			out += "<sub>" + s + "</sub>"
		}
		if s := kid("sup"); mathLen(s) > 0 {
			out += "<sup>" + s + "</sup>"
		}
		return out + kid("e")
	case "func":
		return kid("fName") + " " + kid("e")
	case "limLow":
		return kid("e") + "<sub>" + kid("lim") + "</sub>"
	case "limUpp":
		return kid("e") + "<sup>" + kid("lim") + "</sup>"
	case "groupChr", "bar", "acc", "box", "borderBox", "phant":
		e := kid("e")
		if k.name == "acc" {
			if c, ok := mathAttr(k, "accPr", "chr"); ok && c != "" {
				return e + html.EscapeString(c)
			}
			return e + "̂"
		}
		return e
	case "eqArr":
		var parts []string
		for _, e := range k.kids {
			if e.name == "e" {
				// "#" separates an equation from its number: show it as a wide gap
				parts = append(parts, strings.ReplaceAll(ic.mathHTML(e), "#", "   "))
			}
		}
		return strings.Join(parts, "<br>")
	case "m": // matrix
		var rows []string
		for _, mr := range k.kids {
			if mr.name != "mr" {
				continue
			}
			var cells []string
			for _, e := range mr.kids {
				if e.name == "e" {
					cells = append(cells, ic.mathHTML(e))
				}
			}
			rows = append(rows, strings.Join(cells, "  "))
		}
		return "[" + strings.Join(rows, "; ") + "]"
	case "rPr", "ctrlPr", "fPr", "dPr", "naryPr", "radPr", "sSupPr", "sSubPr", "sSubSupPr", "funcPr",
		"accPr", "barPr", "groupChrPr", "eqArrPr", "mPr", "oMathParaPr", "boxPr", "limLowPr", "limUppPr", "argPr", "sPrePr", "phantPr", "borderBoxPr":
		return ""
	}
	return ic.mathHTML(k)
}

// math converts an m:oMath or m:oMathPara element and adds it as raw HTML.
func (ic *inlineConv) math(n *xnode, base rpr) {
	h := ic.mathHTML(n)
	if h == "" {
		return
	}
	style := "font-family:'Cambria Math', 'Cambria', serif"
	if base.sz > 0 {
		style += ";" + fmtPt("font-size", float64(base.sz)/2)
	}
	ic.addRaw(`<span style="` + html.EscapeString(style) + `">` + h + `</span>`)
}
