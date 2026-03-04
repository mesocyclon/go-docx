package opc

import (
	"fmt"
	"path"
	"strings"
)

// PackURI represents a URI for a part within an OPC package (e.g. "/word/document.xml").
// It is always absolute, beginning with "/".
type PackURI string

// Well-known PackURIs.
const (
	PackageURI      PackURI = "/"
	ContentTypesURI PackURI = "/[Content_Types].xml"
)

// NewPackURI creates a PackURI, ensuring it starts with "/".
func NewPackURI(uri string) PackURI {
	if uri == "" || uri[0] != '/' {
		uri = "/" + uri
	}
	return PackURI(uri)
}

// FromRelRef resolves a relative reference against a base URI to produce an absolute PackURI.
// For example, FromRelRef("/word", "media/image1.png") returns "/word/media/image1.png".
func FromRelRef(baseURI, relativeRef string) PackURI {
	joined := path.Join(baseURI, relativeRef)
	abs := path.Clean(joined)
	if abs == "." {
		abs = "/"
	}
	if !strings.HasPrefix(abs, "/") {
		abs = "/" + abs
	}
	return PackURI(abs)
}

// BaseURI returns the directory portion, e.g. "/word" for "/word/document.xml".
func (u PackURI) BaseURI() string {
	dir, _ := path.Split(string(u))
	// Remove trailing slash unless root
	if len(dir) > 1 && dir[len(dir)-1] == '/' {
		dir = dir[:len(dir)-1]
	}
	return dir
}

// Filename returns the filename portion, e.g. "document.xml" for "/word/document.xml".
func (u PackURI) Filename() string {
	_, file := path.Split(string(u))
	return file
}

// Ext returns the extension without the leading dot, e.g. "xml" for "/word/document.xml".
func (u PackURI) Ext() string {
	ext := path.Ext(string(u))
	if strings.HasPrefix(ext, ".") {
		return ext[1:]
	}
	return ext
}

// Membername returns the pack URI without the leading slash — the form used as the
// ZIP member name. Returns "" for the package pseudo-partname "/".
func (u PackURI) Membername() string {
	s := string(u)
	if len(s) > 1 {
		return s[1:]
	}
	return ""
}

// RelsURI returns the pack URI of the .rels part corresponding to this pack URI.
// E.g. "/word/document.xml" → "/word/_rels/document.xml.rels".
func (u PackURI) RelsURI() PackURI {
	filename := u.Filename()
	baseDir := u.BaseURI()
	relsFilename := filename + ".rels"
	relsURI := path.Join(baseDir, "_rels", relsFilename)
	return PackURI(relsURI)
}

// RelativeRef returns a relative reference from baseURI to this pack URI.
func (u PackURI) RelativeRef(baseURI string) string {
	if baseURI == "/" {
		return string(u)[1:] // strip leading "/"
	}
	return relativePath(baseURI, string(u))
}

// String returns the pack URI as a string.
func (u PackURI) String() string {
	return string(u)
}

// relativePath computes a relative path from base to target using POSIX semantics.
func relativePath(base, target string) string {
	// Split into components
	baseParts := splitPath(base)
	targetParts := splitPath(target)

	// Find common prefix length
	common := 0
	for common < len(baseParts) && common < len(targetParts) && baseParts[common] == targetParts[common] {
		common++
	}

	// Number of ".." needed
	ups := len(baseParts) - common
	var parts []string
	for i := 0; i < ups; i++ {
		parts = append(parts, "..")
	}
	parts = append(parts, targetParts[common:]...)

	result := strings.Join(parts, "/")
	if result == "" {
		return "."
	}
	return result
}

func splitPath(p string) []string {
	p = strings.TrimPrefix(p, "/")
	if p == "" {
		return nil
	}
	return strings.Split(p, "/")
}

// Idx extracts the numeric index from a tuple partname.
// For example, "/word/media/image21.png" returns 21, true.
// Singleton partnames like "/word/document.xml" return 0, false.
// The number must start with a non-zero digit, matching Python's
// [1-9][0-9]* regex pattern.
func (u PackURI) Idx() (int, bool) {
	filename := u.Filename()
	if filename == "" {
		return 0, false
	}
	// Remove extension
	dot := strings.LastIndex(filename, ".")
	namePart := filename
	if dot >= 0 {
		namePart = filename[:dot]
	}
	// Extract trailing digits
	i := len(namePart)
	for i > 0 && namePart[i-1] >= '0' && namePart[i-1] <= '9' {
		i--
	}
	if i == len(namePart) || i == 0 {
		return 0, false
	}
	digits := namePart[i:]
	// Must start with [1-9] to match Python's regex pattern
	if digits[0] == '0' {
		return 0, false
	}
	// Parse the numeric suffix
	n := 0
	for _, c := range digits {
		n = n*10 + int(c-'0')
	}
	return n, true
}

// Validate checks if the PackURI is valid (starts with "/").
func (u PackURI) Validate() error {
	s := string(u)
	if s == "" || s[0] != '/' {
		return fmt.Errorf("PackURI must begin with slash, got %q", s)
	}
	return nil
}
