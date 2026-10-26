package pptx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"
)

// Namespace URIs used throughout the package.
const (
	nsA    = "http://schemas.openxmlformats.org/drawingml/2006/main"
	nsP    = "http://schemas.openxmlformats.org/presentationml/2006/main"
	nsR    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsMC   = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	nsRel  = "http://schemas.openxmlformats.org/package/2006/relationships"
	nsCT   = "http://schemas.openxmlformats.org/package/2006/content-types"
	nsXML  = "http://www.w3.org/XML/1998/namespace"
	relDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

var shortNS = map[string]string{"a": nsA, "p": nsP, "r": nsR, "mc": nsMC}

// Node is a minimal XML DOM element (or text node when Local == "").
// Original prefixes and namespace declarations are kept so that elements
// can be re-serialized verbatim.
type Node struct {
	Prefix, Local, URI string
	Attrs              []Attr
	NS                 [][2]string // namespace declarations on this element: {prefix, uri}; prefix "" = default
	Children           []*Node
	Text               string
	Parent             *Node
}

// Attr is an attribute with its original prefix and resolved namespace.
type Attr struct {
	Prefix, Local, URI, Value string
}

func (n *Node) isText() bool { return n.Local == "" }

// parseXML parses a whole document and returns its root element.
func parseXML(b []byte) (*Node, error) {
	d := xml.NewDecoder(bytes.NewReader(b))
	d.Strict = false
	d.CharsetReader = func(label string, r io.Reader) (io.Reader, error) { return r, nil }
	var root *Node
	var stack []*Node
	for {
		tok, err := d.RawToken()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			n := &Node{Prefix: t.Name.Space, Local: t.Name.Local}
			for _, a := range t.Attr {
				if a.Name.Space == "xmlns" {
					n.NS = append(n.NS, [2]string{a.Name.Local, a.Value})
				} else if a.Name.Space == "" && a.Name.Local == "xmlns" {
					n.NS = append(n.NS, [2]string{"", a.Value})
				} else {
					n.Attrs = append(n.Attrs, Attr{Prefix: a.Name.Space, Local: a.Name.Local, Value: a.Value})
				}
			}
			if len(stack) > 0 {
				p := stack[len(stack)-1]
				n.Parent = p
				p.Children = append(p.Children, n)
			} else if root == nil {
				root = n
			}
			n.URI = n.lookupNS(n.Prefix)
			for i := range n.Attrs {
				if p := n.Attrs[i].Prefix; p != "" {
					if p == "xml" {
						n.Attrs[i].URI = nsXML
					} else {
						n.Attrs[i].URI = n.lookupNS(p)
					}
				}
			}
			stack = append(stack, n)
		case xml.EndElement:
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			if len(stack) > 0 {
				p := stack[len(stack)-1]
				p.Children = append(p.Children, &Node{Text: string(t), Parent: p})
			}
		}
	}
	if root == nil {
		return nil, fmt.Errorf("empty XML document")
	}
	return root, nil
}

func (n *Node) lookupNS(prefix string) string {
	for e := n; e != nil; e = e.Parent {
		for _, d := range e.NS {
			if d[0] == prefix {
				return d[1]
			}
		}
	}
	if prefix == "xml" {
		return nsXML
	}
	return ""
}

func splitName(name string) (string, string) {
	if i := strings.IndexByte(name, ':'); i >= 0 {
		return shortNS[name[:i]], name[i+1:]
	}
	return "", name
}

// Is reports whether n is the element "a:off" style short name.
func (n *Node) Is(name string) bool {
	if n == nil || n.isText() {
		return false
	}
	u, l := splitName(name)
	return n.Local == l && n.URI == u
}

// C returns the first child element with the short name.
func (n *Node) C(name string) *Node {
	if n == nil {
		return nil
	}
	u, l := splitName(name)
	for _, c := range n.Children {
		if c.Local == l && c.URI == u {
			return c
		}
	}
	return nil
}

// Path follows a chain of child names.
func (n *Node) Path(names ...string) *Node {
	for _, nm := range names {
		n = n.C(nm)
		if n == nil {
			return nil
		}
	}
	return n
}

// All returns all child elements with the short name.
func (n *Node) All(name string) []*Node {
	if n == nil {
		return nil
	}
	u, l := splitName(name)
	var out []*Node
	for _, c := range n.Children {
		if c.Local == l && c.URI == u {
			out = append(out, c)
		}
	}
	return out
}

// Elems returns all child elements.
func (n *Node) Elems() []*Node {
	if n == nil {
		return nil
	}
	var out []*Node
	for _, c := range n.Children {
		if !c.isText() {
			out = append(out, c)
		}
	}
	return out
}

// Find returns the first descendant (depth first, excluding n) with the name.
func (n *Node) Find(name string) *Node {
	if n == nil {
		return nil
	}
	for _, c := range n.Children {
		if c.isText() {
			continue
		}
		if c.Is(name) {
			return c
		}
		if f := c.Find(name); f != nil {
			return f
		}
	}
	return nil
}

// Walk visits n and all descendant elements.
func (n *Node) Walk(f func(*Node)) {
	if n == nil || n.isText() {
		return
	}
	f(n)
	for _, c := range n.Children {
		c.Walk(f)
	}
}

// A returns the value of an attribute. name may be "r:id" style or plain.
func (n *Node) A(name string) string {
	v, _ := n.AttrOK(name)
	return v
}

// AttrOK returns an attribute value and whether it exists.
func (n *Node) AttrOK(name string) (string, bool) {
	if n == nil {
		return "", false
	}
	u, l := splitName(name)
	for _, a := range n.Attrs {
		if a.Local == l && a.URI == u {
			return a.Value, true
		}
	}
	return "", false
}

// SetA sets or adds a non-namespaced attribute.
func (n *Node) SetA(name, val string) {
	for i, a := range n.Attrs {
		if a.Local == name && a.URI == "" {
			n.Attrs[i].Value = val
			return
		}
	}
	n.Attrs = append(n.Attrs, Attr{Local: name, Value: val})
}

// DelA removes a non-namespaced attribute.
func (n *Node) DelA(name string) {
	out := n.Attrs[:0]
	for _, a := range n.Attrs {
		if !(a.Local == name && a.URI == "") {
			out = append(out, a)
		}
	}
	n.Attrs = out
}

// TextContent concatenates all descendant text.
func (n *Node) TextContent() string {
	if n == nil {
		return ""
	}
	if n.isText() {
		return n.Text
	}
	var sb strings.Builder
	for _, c := range n.Children {
		sb.WriteString(c.TextContent())
	}
	return sb.String()
}

// RemoveChild removes c from n's children.
func (n *Node) RemoveChild(c *Node) {
	for i, x := range n.Children {
		if x == c {
			n.Children = append(n.Children[:i], n.Children[i+1:]...)
			return
		}
	}
}

// inScopeNS returns the namespace bindings visible at n (nearest wins).
func (n *Node) inScopeNS() map[string]string {
	m := map[string]string{}
	var chain []*Node
	for e := n; e != nil; e = e.Parent {
		chain = append(chain, e)
	}
	for i := len(chain) - 1; i >= 0; i-- {
		for _, d := range chain[i].NS {
			m[d[0]] = d[1]
		}
	}
	return m
}

// Serialize writes n as a self-contained fragment: every namespace binding in
// scope at n is declared on the fragment root.
func (n *Node) Serialize() string {
	var sb strings.Builder
	scope := n.inScopeNS()
	own := map[string]bool{}
	for _, d := range n.NS {
		own[d[0]] = true
	}
	var extra [][2]string
	keys := make([]string, 0, len(scope))
	for k := range scope {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k := range keys {
		if !own[k] {
			extra = append(extra, [2]string{k, scope[k]})
		}
	}
	n.write(&sb, extra)
	return sb.String()
}

func (n *Node) write(sb *strings.Builder, extraNS [][2]string) {
	if n.isText() {
		xmlEscText(sb, n.Text)
		return
	}
	name := n.Local
	if n.Prefix != "" {
		name = n.Prefix + ":" + n.Local
	}
	sb.WriteByte('<')
	sb.WriteString(name)
	for _, d := range append(append([][2]string{}, n.NS...), extraNS...) {
		if d[0] == "" {
			sb.WriteString(` xmlns="`)
		} else {
			sb.WriteString(` xmlns:` + d[0] + `="`)
		}
		xmlEscAttr(sb, d[1])
		sb.WriteByte('"')
	}
	for _, a := range n.Attrs {
		sb.WriteByte(' ')
		if a.Prefix != "" {
			sb.WriteString(a.Prefix + ":")
		}
		sb.WriteString(a.Local + `="`)
		xmlEscAttr(sb, a.Value)
		sb.WriteByte('"')
	}
	if len(n.Children) == 0 {
		sb.WriteString("/>")
		return
	}
	sb.WriteByte('>')
	for _, c := range n.Children {
		c.write(sb, nil)
	}
	sb.WriteString("</" + name + ">")
}

// Document serializes a whole document with the XML declaration.
func (n *Node) Document() []byte {
	var sb strings.Builder
	sb.WriteString(xmlDecl)
	n.write(&sb, nil)
	return []byte(sb.String())
}

func xmlEscText(sb *strings.Builder, s string) {
	for _, r := range s {
		switch r {
		case '&':
			sb.WriteString("&amp;")
		case '<':
			sb.WriteString("&lt;")
		case '>':
			sb.WriteString("&gt;")
		default:
			if r < 0x20 && r != '\t' && r != '\n' && r != '\r' {
				continue
			}
			if r == 0xFFFE || r == 0xFFFF {
				continue
			}
			sb.WriteRune(r)
		}
	}
}

func xmlEscAttr(sb *strings.Builder, s string) {
	for _, r := range s {
		switch r {
		case '&':
			sb.WriteString("&amp;")
		case '<':
			sb.WriteString("&lt;")
		case '>':
			sb.WriteString("&gt;")
		case '"':
			sb.WriteString("&quot;")
		case '\n':
			sb.WriteString("&#xA;")
		case '\r':
			sb.WriteString("&#xD;")
		case '\t':
			sb.WriteString("&#x9;")
		default:
			if r < 0x20 || r == 0xFFFE || r == 0xFFFF {
				continue
			}
			sb.WriteRune(r)
		}
	}
}

// esc escapes a string for use in attribute or text content.
func esc(s string) string {
	var sb strings.Builder
	xmlEscAttr(&sb, s)
	return sb.String()
}

func escText(s string) string {
	var sb strings.Builder
	xmlEscText(&sb, s)
	return sb.String()
}
