// Package xmlutil provides low-level helpers shared across the docx module
// for working with XML attribute values and element trees.
package xmlutil

// IsDigits reports whether s is non-empty and consists only of ASCII digits.
//
// This is used to identify bare numeric "id" attributes on OOXML drawing
// elements (wp:docPr, pic:cNvPr) that must be unique within a story part.
func IsDigits(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}
