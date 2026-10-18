package docx

import (
	"html"
	"strings"
)

// hnode is a node of a tolerant, minimal HTML tree. tag == "" means text.
type hnode struct {
	tag   string
	attrs map[string]string
	kids  []*hnode
	text  string
}

func (n *hnode) attr(k string) string { return n.attrs[k] }

var voidTags = map[string]bool{
	"br": true, "img": true, "hr": true, "input": true, "meta": true, "link": true, "col": true,
	"wbr": true, "source": true, "area": true, "base": true, "embed": true, "param": true, "track": true,
}

var inlineTags = map[string]bool{
	"a": true, "b": true, "strong": true, "i": true, "em": true, "u": true, "ins": true, "s": true,
	"strike": true, "del": true, "sup": true, "sub": true, "span": true, "font": true, "code": true,
	"mark": true, "small": true, "big": true, "abbr": true, "cite": true, "q": true, "kbd": true,
	"samp": true, "var": true, "label": true, "tt": true,
}

// parseHTML builds a tree from an HTML fragment or document, tolerating
// unclosed and misnested tags. It also returns the <title> text, if any.
func parseHTML(s string) (*hnode, string) {
	root := &hnode{tag: "#root"}
	stack := []*hnode{root}
	title := ""
	top := func() *hnode { return stack[len(stack)-1] }
	popTo := func(i int) { stack = stack[:i] }
	// find returns index of nearest open tag in set, stopping at any boundary tag.
	find := func(set, boundary map[string]bool) int {
		for i := len(stack) - 1; i > 0; i-- {
			t := stack[i].tag
			if set[t] {
				return i
			}
			if boundary[t] {
				return -1
			}
		}
		return -1
	}
	i := 0
	for i < len(s) {
		if s[i] != '<' {
			j := strings.IndexByte(s[i:], '<')
			if j < 0 {
				j = len(s) - i
			}
			top().kids = append(top().kids, &hnode{text: html.UnescapeString(s[i : i+j])})
			i += j
			continue
		}
		rest := s[i:]
		switch {
		case strings.HasPrefix(rest, "<!--"):
			j := strings.Index(rest[4:], "-->")
			if j < 0 {
				i = len(s)
			} else {
				i += 4 + j + 3
			}
			continue
		case strings.HasPrefix(rest, "<!") || strings.HasPrefix(rest, "<?"):
			j := strings.IndexByte(rest, '>')
			if j < 0 {
				i = len(s)
			} else {
				i += j + 1
			}
			continue
		case strings.HasPrefix(rest, "</"):
			j := strings.IndexByte(rest, '>')
			if j < 0 {
				i = len(s)
				continue
			}
			name := strings.ToLower(strings.TrimSpace(rest[2:j]))
			if k := strings.IndexAny(name, " \t\r\n/"); k >= 0 {
				name = name[:k]
			}
			i += j + 1
			for k := len(stack) - 1; k > 0; k-- {
				if stack[k].tag == name {
					popTo(k)
					break
				}
			}
			continue
		}
		if len(rest) < 2 || !isAlpha(rest[1]) {
			top().kids = append(top().kids, &hnode{text: "<"})
			i++
			continue
		}
		// start tag
		j := 1
		for j < len(rest) && isNameChar(rest[j]) {
			j++
		}
		n := &hnode{tag: strings.ToLower(rest[1:j]), attrs: map[string]string{}}
		selfClose := false
		for j < len(rest) {
			for j < len(rest) && isSpace(rest[j]) {
				j++
			}
			if j >= len(rest) {
				break
			}
			if rest[j] == '>' {
				j++
				break
			}
			if rest[j] == '/' {
				j++
				if j < len(rest) && rest[j] == '>' {
					selfClose = true
					j++
					break
				}
				continue
			}
			k := j
			for k < len(rest) && !isSpace(rest[k]) && rest[k] != '=' && rest[k] != '>' && rest[k] != '/' {
				k++
			}
			if k == j {
				j++
				continue
			}
			an := strings.ToLower(rest[j:k])
			j = k
			for j < len(rest) && isSpace(rest[j]) {
				j++
			}
			val := ""
			if j < len(rest) && rest[j] == '=' {
				j++
				for j < len(rest) && isSpace(rest[j]) {
					j++
				}
				if j < len(rest) && (rest[j] == '"' || rest[j] == '\'') {
					q := rest[j]
					e := strings.IndexByte(rest[j+1:], q)
					if e < 0 {
						val = rest[j+1:]
						j = len(rest)
					} else {
						val = rest[j+1 : j+1+e]
						j += e + 2
					}
				} else {
					k = j
					for k < len(rest) && !isSpace(rest[k]) && rest[k] != '>' {
						k++
					}
					val = rest[j:k]
					j = k
				}
			}
			if _, dup := n.attrs[an]; !dup {
				n.attrs[an] = html.UnescapeString(val)
			}
		}
		i += j
		// raw text elements
		if n.tag == "script" || n.tag == "style" || n.tag == "title" || n.tag == "textarea" {
			end := strings.Index(strings.ToLower(s[i:]), "</"+n.tag)
			raw := ""
			if end < 0 {
				raw, i = s[i:], len(s)
			} else {
				raw = s[i : i+end]
				i += end
				if k := strings.IndexByte(s[i:], '>'); k >= 0 {
					i += k + 1
				} else {
					i = len(s)
				}
			}
			if n.tag == "title" && title == "" {
				title = strings.TrimSpace(html.UnescapeString(raw))
			}
			continue
		}
		// implied end tags
		switch n.tag {
		case "li":
			if k := find(map[string]bool{"li": true}, map[string]bool{"ul": true, "ol": true, "table": true}); k > 0 {
				popTo(k)
			}
		case "td", "th":
			if k := find(map[string]bool{"td": true, "th": true}, map[string]bool{"tr": true, "table": true}); k > 0 {
				popTo(k)
			}
		case "tr":
			if k := find(map[string]bool{"tr": true}, map[string]bool{"table": true}); k > 0 {
				popTo(k)
			}
		case "thead", "tbody", "tfoot":
			if k := find(map[string]bool{"thead": true, "tbody": true, "tfoot": true}, map[string]bool{"table": true}); k > 0 {
				popTo(k)
			}
		}
		if !inlineTags[n.tag] && n.tag != "img" && n.tag != "br" && n.tag != "wbr" {
			// a block start closes an open <p> (through inline ancestors)
			for k := len(stack) - 1; k > 0; k-- {
				t := stack[k].tag
				if t == "p" {
					popTo(k)
					break
				}
				if !inlineTags[t] {
					break
				}
			}
		}
		top().kids = append(top().kids, n)
		if !selfClose && !voidTags[n.tag] {
			stack = append(stack, n)
		}
	}
	return root, title
}

func isAlpha(c byte) bool { return c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' }
func isNameChar(c byte) bool {
	return isAlpha(c) || c >= '0' && c <= '9' || c == '-' || c == ':' || c == '_'
}
func isSpace(c byte) bool { return c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '\f' }

// parseCSS parses an inline style attribute into lowercase property -> value.
func parseCSS(s string) map[string]string {
	m := map[string]string{}
	for _, decl := range strings.Split(s, ";") {
		k, v, ok := strings.Cut(decl, ":")
		if !ok {
			continue
		}
		k = strings.ToLower(strings.TrimSpace(k))
		v = strings.TrimSpace(strings.TrimSuffix(strings.TrimSpace(v), "!important"))
		if k != "" && v != "" {
			m[k] = strings.TrimSpace(v)
		}
	}
	return m
}
