package docx

import (
	"fmt"
	"math"
	"strconv"
)

// EMU conversion constants.
const (
	EmusPerInch = 914400
	EmusPerCm   = 360000
	EmusPerMm   = 36000
	EmusPerPt   = 12700
	EmusPerTwip = 635
)

// Length represents a length value stored internally as English Metric Units (EMU).
// There are 914,400 EMUs per inch and 36,000 per millimeter.
type Length int64

// Inches returns the length in inches.
func (l Length) Inches() float64 { return float64(l) / float64(EmusPerInch) }

// Cm returns the length in centimeters.
func (l Length) Cm() float64 { return float64(l) / float64(EmusPerCm) }

// Mm returns the length in millimeters.
func (l Length) Mm() float64 { return float64(l) / float64(EmusPerMm) }

// Pt returns the length in points.
func (l Length) Pt() float64 { return float64(l) / float64(EmusPerPt) }

// Twips returns the length in twips (twentieth of a point).
func (l Length) Twips() int { return int(math.Round(float64(l) / float64(EmusPerTwip))) }

// Emu returns the raw EMU value.
func (l Length) Emu() int64 { return int64(l) }

// String returns a human-readable representation of the length.
func (l Length) String() string {
	return fmt.Sprintf("%d EMU", int64(l))
}

// Inches creates a Length from a value in inches.
func Inches(v float64) Length { return Length(int64(v * float64(EmusPerInch))) }

// Cm creates a Length from a value in centimeters.
func Cm(v float64) Length { return Length(int64(v * float64(EmusPerCm))) }

// Mm creates a Length from a value in millimeters.
func Mm(v float64) Length { return Length(int64(v * float64(EmusPerMm))) }

// Pt creates a Length from a value in points.
func Pt(v float64) Length { return Length(int64(v * float64(EmusPerPt))) }

// Twips creates a Length from a value in twips.
func Twips(v float64) Length { return Length(int64(v * float64(EmusPerTwip))) }

// Emu creates a Length from a raw EMU value.
func Emu(v int64) Length { return Length(v) }

// RGBColor represents an RGB color as three bytes (red, green, blue).
type RGBColor [3]byte

// NewRGBColor creates a new RGBColor from individual red, green, and blue components.
func NewRGBColor(r, g, b byte) RGBColor {
	return RGBColor{r, g, b}
}

// RGBColorFromString parses a six-character hex string (e.g. "3C2F80") into an RGBColor.
func RGBColorFromString(hex string) (RGBColor, error) {
	if len(hex) != 6 {
		return RGBColor{}, fmt.Errorf("RGBColor hex string must be 6 characters, got %q", hex)
	}

	var c RGBColor
	for i := 0; i < 3; i++ {
		val, err := parseHexByte(hex[i*2 : i*2+2])
		if err != nil {
			return RGBColor{}, fmt.Errorf("invalid hex in RGBColor string %q: %w", hex, err)
		}
		c[i] = val
	}
	return c, nil
}

// String returns the six-character uppercase hex representation (e.g. "3C2F80").
func (c RGBColor) String() string {
	return fmt.Sprintf("%02X%02X%02X", c[0], c[1], c[2])
}

// R returns the red component.
func (c RGBColor) R() byte { return c[0] }

// G returns the green component.
func (c RGBColor) G() byte { return c[1] }

// B returns the blue component.
func (c RGBColor) B() byte { return c[2] }

// parseHexByte parses a two-character hex string into a byte.
func parseHexByte(s string) (byte, error) {
	v, err := strconv.ParseUint(s, 16, 8)
	if err != nil {
		return 0, fmt.Errorf("invalid hex byte %q: %w", s, err)
	}
	return byte(v), nil
}
