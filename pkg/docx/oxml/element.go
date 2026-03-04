package oxml

import (
	"bytes"
	"strings"

	"github.com/beevik/etree"
)

// Element is a base wrapper over *etree.Element that provides namespace-aware
// navigation, insertion, and attribute operations. All CT_* structures embed Element.
type Element struct {
	e *etree.Element
}

// RawElement returns the underlying *etree.Element.
//
// This accessor is provided for interoperability with code that needs direct
// access to the etree node (e.g. cross-package construction of CT_* wrappers).
// Prefer using Element methods (FindChild, GetAttr, etc.) when possible.
func (el *Element) RawElement() *etree.Element { return el.e }

// WrapElement creates an Element value wrapping the given *etree.Element.
// This is the cross-package equivalent of Element{e: e} for constructing
// CT_* types that embed Element.
func WrapElement(e *etree.Element) Element {
	return Element{e: e}
}

// NewElement wraps an existing *etree.Element (returns a pointer).
func NewElement(e *etree.Element) *Element {
	if e == nil {
		return nil
	}
	return &Element{e: e}
}

// --- Tag ---

// Tag returns the Clark-name of the element's tag, e.g. "{http://...}p".
func (el *Element) Tag() string {
	return clarkFromEtree(el.e)
}

// --- Child navigation ---

// FindChild finds the first child element matching a namespace-prefixed tag like "w:pPr".
// Returns nil if not found.
func (el *Element) FindChild(nspTag string) *etree.Element {
	space, local := resolveNspTag(nspTag)
	for _, child := range el.e.ChildElements() {
		if child.Space == space && child.Tag == local {
			return child
		}
	}
	return nil
}

// FindAllChildren returns all child elements matching the given namespace-prefixed tag.
func (el *Element) FindAllChildren(nspTag string) []*etree.Element {
	space, local := resolveNspTag(nspTag)
	var result []*etree.Element
	for _, child := range el.e.ChildElements() {
		if child.Space == space && child.Tag == local {
			result = append(result, child)
		}
	}
	return result
}

// FirstChildIn returns the first child element whose tag matches one of the given
// namespace-prefixed tags. Returns nil if none found.
func (el *Element) FirstChildIn(tags ...string) *etree.Element {
	for _, tag := range tags {
		if child := el.FindChild(tag); child != nil {
			return child
		}
	}
	return nil
}

// --- Insertion and removal ---

// InsertElementBefore inserts child before the first element matching one of the
// successor tags. If no successor is found, child is appended to the end.
// This is the core mechanism for maintaining correct element ordering in OOXML.
func (el *Element) InsertElementBefore(child *etree.Element, successors ...string) {
	for _, succ := range successors {
		if found := el.FindChild(succ); found != nil {
			insertBefore(el.e, child, found)
			return
		}
	}
	el.e.AddChild(child)
}

// RemoveAll removes all child elements with the specified namespace-prefixed tags.
func (el *Element) RemoveAll(tags ...string) {
	for _, tag := range tags {
		matches := el.FindAllChildren(tag)
		for _, m := range matches {
			el.e.RemoveChild(m)
		}
	}
}

// Remove removes a specific child element.
func (el *Element) Remove(child *etree.Element) {
	el.e.RemoveChild(child)
}

// --- Attributes ---

// GetAttr returns the value of an attribute. The name can be a simple name like "val"
// or a namespace-prefixed name like "w:val" or a Clark-name like "{http://...}val".
func (el *Element) GetAttr(name string) (string, bool) {
	space, local := resolveAttrName(name)
	attr := el.e.SelectAttr(local)
	if attr == nil {
		return "", false
	}
	// If a namespace was requested, verify it matches.
	if space != "" && attr.Space != space {
		// Try to find among all attrs
		for _, a := range el.e.Attr {
			if a.Key == local && a.Space == space {
				return a.Value, true
			}
		}
		return "", false
	}
	return attr.Value, true
}

// SetAttr sets an attribute value. The name can be simple or namespace-prefixed.
func (el *Element) SetAttr(name, value string) {
	space, local := resolveAttrName(name)
	if space != "" {
		el.e.CreateAttr(space+":"+local, value)
	} else {
		el.e.CreateAttr(local, value)
	}
}

// RemoveAttr removes an attribute by name.
func (el *Element) RemoveAttr(name string) {
	space, local := resolveAttrName(name)
	el.e.RemoveAttr(local)
	if space != "" {
		// Also try removing with qualified name
		el.e.RemoveAttr(space + ":" + local)
	}
}

// --- Text ---

// Text returns the direct text content of the element.
func (el *Element) Text() string {
	return el.e.Text()
}

// SetText sets the text content of the element.
func (el *Element) SetText(text string) {
	el.e.SetText(text)
}

// --- XPath ---

// XPath executes an XPath query using standard OOXML namespace mappings.
// Note: etree XPath uses local paths; for namespace-aware lookups, prefer
// FindChild/FindAllChildren.
func (el *Element) XPath(expr string) []*etree.Element {
	return el.e.FindElements(expr)
}

// --- Utilities ---

// AddSubElement creates a new child element with the given namespace-prefixed tag
// and appends it.
func (el *Element) AddSubElement(nspTag string) *etree.Element {
	nspt := NewNSPTag(nspTag)
	child := el.e.CreateElement(nspt.LocalPart())
	child.Space = nspt.Prefix()
	return child
}

// Xml returns an XML string representation for debugging.
func (el *Element) Xml() string {
	doc := etree.NewDocument()
	doc.SetRoot(el.e.Copy())
	doc.Indent(2)
	var buf bytes.Buffer
	_, _ = doc.WriteTo(&buf)
	return buf.String()
}

// --- Internal helpers ---

// resolveNspTag splits a namespace-prefixed tag (e.g. "w:p") into the prefix and local part.
// Returns (prefix, localPart). etree uses prefix as Space field.
func resolveNspTag(nspTag string) (prefix, local string) {
	if p, l, ok := strings.Cut(nspTag, ":"); ok {
		return p, l
	}
	return "", nspTag
}

// resolveAttrName resolves an attribute name which can be:
//   - simple: "val" → ("", "val")
//   - prefixed: "w:val" → ("w", "val")
//   - clark: "{http://...}val" → resolves to prefix form
func resolveAttrName(name string) (space, local string) {
	if strings.HasPrefix(name, "{") {
		// Clark notation
		closeBrace := strings.Index(name, "}")
		if closeBrace > 0 {
			uri := name[1:closeBrace]
			local = name[closeBrace+1:]
			if pfx, ok := pfxmap[uri]; ok {
				return pfx, local
			}
			return "", local
		}
	}
	if p, l, ok := strings.Cut(name, ":"); ok {
		return p, l
	}
	return "", name
}

// clarkFromEtree builds a Clark-notation tag from an etree.Element.
func clarkFromEtree(e *etree.Element) string {
	if e.Space == "" {
		return e.Tag
	}
	// etree stores the prefix in Space, resolve to URI
	if uri, ok := nsmap[e.Space]; ok {
		return "{" + uri + "}" + e.Tag
	}
	// If Space is already a URI (shouldn't normally happen with our usage)
	return "{" + e.Space + "}" + e.Tag
}

// insertBefore inserts newChild before refChild within parent.
func insertBefore(parent, newChild, refChild *etree.Element) {
	// Find index of refChild among parent's child tokens
	tokens := parent.Child
	refIndex := -1
	for i, t := range tokens {
		if elem, ok := t.(*etree.Element); ok && elem == refChild {
			refIndex = i
			break
		}
	}
	if refIndex < 0 {
		parent.AddChild(newChild)
		return
	}

	// Remove newChild from any current parent
	if p := newChild.Parent(); p != nil {
		p.RemoveChild(newChild)
	}

	// Insert at the found index
	// We need to manipulate the Child slice directly
	parent.InsertChildAt(refIndex, newChild)
}
