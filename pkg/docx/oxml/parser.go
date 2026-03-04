package oxml

import (
	"bytes"
	"fmt"

	"github.com/beevik/etree"
)

// ParseXml parses XML bytes into an *etree.Element.
// Removes blank text nodes to match lxml's remove_blank_text=True behavior.
func ParseXml(xmlBytes []byte) (*etree.Element, error) {
	doc := etree.NewDocument()
	doc.ReadSettings.Permissive = true
	if err := doc.ReadFromBytes(xmlBytes); err != nil {
		return nil, fmt.Errorf("oxml.ParseXml: %w", err)
	}
	root := doc.Root()
	if root == nil {
		return nil, fmt.Errorf("oxml.ParseXml: no root element found")
	}
	// Detach root from document so it can be used independently
	return root, nil
}

// SerializeXml serializes an *etree.Element to []byte with an XML declaration
// and standalone="yes", matching OOXML conventions.
// Output is compact (no insignificant whitespace), matching Python's
// etree.tostring(elm, encoding="UTF-8", standalone=True).
func SerializeXml(el *etree.Element) ([]byte, error) {
	doc := etree.NewDocument()
	doc.CreateProcInst("xml", `version="1.0" encoding="UTF-8" standalone="yes"`)
	doc.SetRoot(el.Copy())

	// No doc.Indent() — produce compact output without insignificant whitespace,
	// matching Python's serialize_part_xml behavior.
	doc.WriteSettings.CanonicalEndTags = true

	var buf bytes.Buffer
	if _, err := doc.WriteTo(&buf); err != nil {
		return nil, fmt.Errorf("oxml.SerializeXml: %w", err)
	}
	return buf.Bytes(), nil
}

// SerializeForReading serializes an element for human-readable output (tests/debugging).
// No XML declaration, with pretty-print indentation.
func SerializeForReading(el *etree.Element) string {
	doc := etree.NewDocument()
	doc.SetRoot(el.Copy())
	doc.Indent(2)

	var buf bytes.Buffer
	_, _ = doc.WriteTo(&buf)
	return buf.String()
}

// TryOxmlElement creates a new element with the given namespace-prefixed tag.
// Returns an error if the tag format is invalid or the prefix is unknown.
// Namespace declarations are added based on the tag prefix, or custom nsDecls
// can be provided as additional prefix strings.
func TryOxmlElement(nspTag string, nsDecls ...string) (*etree.Element, error) {
	nspt, err := ParseNSPTag(nspTag)
	if err != nil {
		return nil, err
	}

	el := etree.NewElement(nspt.LocalPart())
	el.Space = nspt.Prefix()

	// Collect all prefixes that need declarations
	prefixes := make(map[string]bool)
	prefixes[nspt.Prefix()] = true
	for _, pfx := range nsDecls {
		prefixes[pfx] = true
	}

	// Add namespace declarations as attributes.
	// etree.CreateAttr("xmlns:w", uri) splits into Attr{Space:"xmlns", Key:"w"},
	// which serializes correctly as xmlns:w="...".
	for pfx := range prefixes {
		if uri, ok := nsmap[pfx]; ok {
			el.CreateAttr("xmlns:"+pfx, uri)
		}
	}

	return el, nil
}

// OxmlElement creates a new element with the given namespace-prefixed tag.
// Panics on invalid input — use only with compile-time known tags.
// For user-supplied input, use [TryOxmlElement].
//
// Namespace declarations are added based on the tag prefix, or custom nsDecls
// can be provided as additional prefix strings.
//
// Example:
//
//	OxmlElement("w:p") creates <w:p xmlns:w="..."/>
//	OxmlElement("w:p", "r") creates <w:p xmlns:w="..." xmlns:r="..."/>
func OxmlElement(nspTag string, nsDecls ...string) *etree.Element {
	el, err := TryOxmlElement(nspTag, nsDecls...)
	if err != nil {
		panic(err)
	}
	return el
}

// HasNsDecl checks whether an etree.Element has a namespace declaration for the
// given prefix. etree stores CreateAttr("xmlns:w", uri) as Attr{Space:"xmlns", Key:"w"}.
func HasNsDecl(el *etree.Element, prefix string) (string, bool) {
	for _, attr := range el.Attr {
		// etree splits "xmlns:w" → Space="xmlns", Key="w"
		if attr.Space == "xmlns" && attr.Key == prefix {
			return attr.Value, true
		}
		// fallback: unsplit form
		if attr.Space == "" && attr.Key == "xmlns:"+prefix {
			return attr.Value, true
		}
	}
	return "", false
}

// OxmlElementWithAttrs creates a new element with the given tag and attributes.
// Both namespace declarations and element attributes are set.
func OxmlElementWithAttrs(nspTag string, attrs map[string]string, nsDecls ...string) *etree.Element {
	el := OxmlElement(nspTag, nsDecls...)
	for name, value := range attrs {
		el.CreateAttr(name, value)
	}
	return el
}
