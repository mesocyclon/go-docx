// Package oxml provides low-level XML element manipulation for Office Open XML documents.
package oxml

import (
	"fmt"
	"strings"
)

// --------------------------------------------------------------------------
// DrawingML / WML namespace URIs
// --------------------------------------------------------------------------

const (
	NsDml     = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NsChart   = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	NsDiagram = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
	NsPicture = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	NsWp      = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	NsWml     = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	NsRel     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// nsmap maps namespace prefixes to their URIs.
var nsmap = map[string]string{
	"a":        NsDml,
	"c":        NsChart,
	"cp":       "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
	"dc":       "http://purl.org/dc/elements/1.1/",
	"dcmitype": "http://purl.org/dc/dcmitype/",
	"dcterms":  "http://purl.org/dc/terms/",
	"dgm":      NsDiagram,
	"m":        "http://schemas.openxmlformats.org/officeDocument/2006/math",
	"pic":      NsPicture,
	"r":        NsRel,
	"sl":       "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
	"w":        NsWml,
	"w14":      "http://schemas.microsoft.com/office/word/2010/wordml",
	"wp":       NsWp,
	"xml":      "http://www.w3.org/XML/1998/namespace",
	"xsi":      "http://www.w3.org/2001/XMLSchema-instance",
}

// pfxmap is the reverse mapping of URI → prefix.
var pfxmap map[string]string

func init() {
	pfxmap = make(map[string]string, len(nsmap))
	for pfx, uri := range nsmap {
		pfxmap[uri] = pfx
	}
}

// LookupNsURI returns the namespace URI for the given prefix and true,
// or ("", false) if the prefix is not registered.
func LookupNsURI(prefix string) (string, bool) {
	uri, ok := nsmap[prefix]
	return uri, ok
}

// LookupPrefix returns the namespace prefix for the given URI and true,
// or ("", false) if the URI is not registered.
func LookupPrefix(uri string) (string, bool) {
	pfx, ok := pfxmap[uri]
	return pfx, ok
}

// TryQn converts a namespace-prefixed tag to Clark notation.
// Returns an error if the prefix is not in nsmap.
// For example, TryQn("w:p") returns "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p".
func TryQn(tag string) (string, error) {
	prefix, local, ok := strings.Cut(tag, ":")
	if !ok {
		return tag, nil
	}
	uri, exists := nsmap[prefix]
	if !exists {
		return "", fmt.Errorf("oxml: unknown namespace prefix %q in tag %q", prefix, tag)
	}
	return "{" + uri + "}" + local, nil
}

// Qn converts a namespace-prefixed tag to Clark notation.
// Panics on unknown prefix — use only with compile-time known tags.
// For user-supplied input, use [TryQn].
func Qn(tag string) string {
	s, err := TryQn(tag)
	if err != nil {
		panic(err)
	}
	return s
}

// NsPfxMap returns a subset of nsmap for the specified prefixes.
func NsPfxMap(prefixes ...string) map[string]string {
	result := make(map[string]string, len(prefixes))
	for _, pfx := range prefixes {
		if uri, ok := nsmap[pfx]; ok {
			result[pfx] = uri
		}
	}
	return result
}

// NamespacePrefixedTag is a value object that knows the semantics of an XML tag
// with a namespace prefix, such as "w:p".
type NamespacePrefixedTag struct {
	prefix    string
	localPart string
	nsURI     string
}

// ParseNSPTag parses a prefixed tag string like "w:p" into a NamespacePrefixedTag.
// Returns an error if the tag format is invalid or the prefix is unknown.
func ParseNSPTag(nstag string) (NamespacePrefixedTag, error) {
	prefix, local, ok := strings.Cut(nstag, ":")
	if !ok {
		return NamespacePrefixedTag{}, fmt.Errorf("oxml: invalid namespace-prefixed tag %q", nstag)
	}
	uri, exists := nsmap[prefix]
	if !exists {
		return NamespacePrefixedTag{}, fmt.Errorf("oxml: unknown namespace prefix %q in tag %q", prefix, nstag)
	}
	return NamespacePrefixedTag{
		prefix:    prefix,
		localPart: local,
		nsURI:     uri,
	}, nil
}

// NewNSPTag creates a NamespacePrefixedTag from a prefixed tag string like "w:p".
// Panics on invalid input — use only with compile-time known tags.
// For user-supplied input, use [ParseNSPTag].
func NewNSPTag(nstag string) NamespacePrefixedTag {
	t, err := ParseNSPTag(nstag)
	if err != nil {
		panic(err)
	}
	return t
}

// ParseNSPTagFromClark parses Clark notation like "{http://...}p" into a NamespacePrefixedTag.
// Returns an error if the format is invalid or the namespace URI is unknown.
func ParseNSPTagFromClark(clark string) (NamespacePrefixedTag, error) {
	if len(clark) == 0 || clark[0] != '{' {
		return NamespacePrefixedTag{}, fmt.Errorf("oxml: invalid Clark notation %q", clark)
	}
	closeBrace := strings.Index(clark, "}")
	if closeBrace < 0 {
		return NamespacePrefixedTag{}, fmt.Errorf("oxml: invalid Clark notation %q", clark)
	}
	nsURI := clark[1:closeBrace]
	local := clark[closeBrace+1:]

	pfx, ok := pfxmap[nsURI]
	if !ok {
		return NamespacePrefixedTag{}, fmt.Errorf("oxml: unknown namespace URI %q", nsURI)
	}
	return NamespacePrefixedTag{
		prefix:    pfx,
		localPart: local,
		nsURI:     nsURI,
	}, nil
}

// NSPTagFromClark creates a NamespacePrefixedTag from Clark notation like "{http://...}p".
// Panics on invalid input — use only with compile-time known tags.
// For user-supplied input, use [ParseNSPTagFromClark].
func NSPTagFromClark(clark string) NamespacePrefixedTag {
	t, err := ParseNSPTagFromClark(clark)
	if err != nil {
		panic(err)
	}
	return t
}

// ClarkName returns the Clark notation for this tag, e.g. "{http://...}p".
func (t NamespacePrefixedTag) ClarkName() string {
	return "{" + t.nsURI + "}" + t.localPart
}

// LocalPart returns the local part of the tag, e.g. "p" for "w:p".
func (t NamespacePrefixedTag) LocalPart() string {
	return t.localPart
}

// Prefix returns the namespace prefix, e.g. "w" for "w:p".
func (t NamespacePrefixedTag) Prefix() string {
	return t.prefix
}

// NsURI returns the namespace URI.
func (t NamespacePrefixedTag) NsURI() string {
	return t.nsURI
}

// String returns the prefixed tag string, e.g. "w:p".
func (t NamespacePrefixedTag) String() string {
	return t.prefix + ":" + t.localPart
}

// NsMap returns a single-member map of this tag's prefix to its namespace URI.
func (t NamespacePrefixedTag) NsMap() map[string]string {
	return map[string]string{t.prefix: t.nsURI}
}
