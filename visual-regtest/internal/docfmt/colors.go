// Package docfmt provides shared document formatting helpers for
// visual-regtest programs that build "before/after" test documents.
package docfmt

import "github.com/vortex/go-docx/pkg/docx"

// Pre-defined colors used across visual regtest documents.
var (
	ColorRed       = docx.NewRGBColor(0xFF, 0x00, 0x00)
	ColorBlue      = docx.NewRGBColor(0x00, 0x00, 0xFF)
	ColorDarkRed   = docx.NewRGBColor(0xCC, 0x00, 0x00)
	ColorDarkGreen = docx.NewRGBColor(0x00, 0x66, 0x00)
	ColorGray      = docx.NewRGBColor(0x88, 0x88, 0x88)
	ColorLightGray = docx.NewRGBColor(0x99, 0x99, 0x99)
)
