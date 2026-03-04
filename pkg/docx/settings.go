package docx

import "github.com/vortex/go-docx/pkg/docx/oxml"

// Settings provides access to document-level settings.
//
// Mirrors Python Settings(ElementProxy).
type Settings struct {
	settings *oxml.CT_Settings
}

// newSettings creates a new Settings proxy wrapping the given CT_Settings element.
func newSettings(elm *oxml.CT_Settings) *Settings {
	return &Settings{settings: elm}
}

// OddAndEvenPagesHeaderFooter returns true if this document has distinct
// odd and even page headers and footers.
//
// Mirrors Python Settings.odd_and_even_pages_header_footer (getter).
func (s *Settings) OddAndEvenPagesHeaderFooter() bool {
	return s.settings.EvenAndOddHeadersVal()
}

// SetOddAndEvenPagesHeaderFooter enables or disables distinct odd/even page
// headers and footers.
//
// Mirrors Python Settings.odd_and_even_pages_header_footer (setter).
func (s *Settings) SetOddAndEvenPagesHeaderFooter(v bool) error {
	return s.settings.SetEvenAndOddHeadersVal(&v)
}
