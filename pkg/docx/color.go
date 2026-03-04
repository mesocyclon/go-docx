package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ColorFormat provides access to color settings such as RGB and theme color.
//
// Mirrors Python ColorFormat(ElementProxy).
type ColorFormat struct {
	rPrOwner rPrProvider
}

// newColorFormat creates a new ColorFormat proxy.
func newColorFormat(owner rPrProvider) *ColorFormat {
	return &ColorFormat{rPrOwner: owner}
}

// RGB returns the RGB color value, or nil if not set or auto.
//
// Mirrors Python ColorFormat.rgb (getter).
func (cf *ColorFormat) RGB() (*RGBColor, error) {
	rPr := cf.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	val, err := rPr.ColorVal()
	if err != nil {
		return nil, fmt.Errorf("docx: reading color value: %w", err)
	}
	if val == nil || *val == "auto" {
		return nil, nil
	}
	c, err := RGBColorFromString(*val)
	if err != nil {
		return nil, fmt.Errorf("docx: parsing RGB %q: %w", *val, err)
	}
	return &c, nil
}

// SetRGB sets the RGB color. Passing nil removes the color.
//
// Mirrors Python ColorFormat.rgb (setter): always removes the existing w:color
// element first (clearing any themeColor attribute), then adds a fresh one if
// value is not nil.
func (cf *ColorFormat) SetRGB(v *RGBColor) error {
	rPr := cf.rPrOwner.RPr()
	if v == nil {
		if rPr == nil {
			return nil
		}
		rPr.RemoveColor()
		return nil
	}
	rPr = cf.rPrOwner.GetOrAddRPr()
	// Remove existing color element entirely (clears themeColor etc.)
	rPr.RemoveColor()
	// Create fresh color element with only the val attribute
	hex := v.String()
	return rPr.SetColorVal(&hex)
}

// ThemeColor returns the theme color index, or nil if not set.
//
// Mirrors Python ColorFormat.theme_color (getter).
func (cf *ColorFormat) ThemeColor() (*enum.MsoThemeColorIndex, error) {
	rPr := cf.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	return rPr.ColorTheme()
}

// SetThemeColor sets the theme color. Passing nil removes the color entirely.
//
// Mirrors Python ColorFormat.theme_color (setter).
func (cf *ColorFormat) SetThemeColor(v *enum.MsoThemeColorIndex) error {
	if v == nil {
		rPr := cf.rPrOwner.RPr()
		if rPr == nil {
			return nil
		}
		return rPr.SetColorVal(nil)
	}
	rPr := cf.rPrOwner.GetOrAddRPr()
	return rPr.SetColorTheme(v)
}

// Type returns the color type: RGB, THEME, AUTO, or nil if no color is applied.
//
// Mirrors Python ColorFormat.type (getter).
func (cf *ColorFormat) Type() (*enum.MsoColorType, error) {
	rPr := cf.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	theme, err := rPr.ColorTheme()
	if err != nil {
		return nil, fmt.Errorf("docx: reading color theme: %w", err)
	}
	if theme != nil {
		ct := enum.MsoColorTypeTheme
		return &ct, nil
	}
	val, err := rPr.ColorVal()
	if err != nil {
		return nil, fmt.Errorf("docx: reading color val: %w", err)
	}
	if val == nil {
		return nil, nil
	}
	if *val == "auto" {
		ct := enum.MsoColorTypeAuto
		return &ct, nil
	}
	ct := enum.MsoColorTypeRGB
	return &ct, nil
}
