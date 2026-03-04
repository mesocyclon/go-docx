package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// color_test.go — ColorFormat (Batch 1)
// Mirrors Python: tests/dml/test_color.py
// -----------------------------------------------------------------------

// Mirrors Python: it_can_change_its_RGB_value — full set of cases
func TestColorFormat_SetRGB(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		newRGB   *RGBColor
		wantRGB  string
		wantNil  bool
	}{
		{
			"set_on_empty",
			``,
			rgbPtr(0xFF, 0x00, 0x00),
			"FF0000",
			false,
		},
		{
			"set_on_existing",
			`<w:rPr><w:color w:val="3C2F80"/></w:rPr>`,
			rgbPtr(0x00, 0xFF, 0x00),
			"00FF00",
			false,
		},
		{
			"remove",
			`<w:rPr><w:color w:val="3C2F80"/></w:rPr>`,
			nil,
			"",
			true,
		},
		{
			"set_over_auto",
			`<w:rPr><w:color w:val="auto"/></w:rPr>`,
			rgbPtr(0x00, 0x00, 0xFF),
			"0000FF",
			false,
		},
		{
			"set_over_theme",
			`<w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr>`,
			rgbPtr(0x12, 0x34, 0x56),
			"123456",
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			cf := newColorFormat(r)
			if err := cf.SetRGB(tt.newRGB); err != nil {
				t.Fatal(err)
			}
			got, err := cf.RGB()
			if err != nil {
				t.Fatal(err)
			}
			if tt.wantNil {
				if got != nil {
					t.Errorf("RGB() = %v, want nil", got)
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil RGB")
			}
			if got.String() != tt.wantRGB {
				t.Errorf("RGB() = %q, want %q", got.String(), tt.wantRGB)
			}
		})
	}
}

// Mirrors Python: it_knows_its_theme_color (getter)
func TestColorFormat_ThemeColor_Getter(t *testing.T) {
	// No theme color
	r := makeR(t, `<w:rPr><w:color w:val="FF0000"/></w:rPr>`)
	cf := newColorFormat(r)
	tc, err := cf.ThemeColor()
	if err != nil {
		t.Fatal(err)
	}
	if tc != nil {
		t.Errorf("ThemeColor() = %v, want nil for plain RGB", tc)
	}

	// With theme color
	r2 := makeR(t, `<w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr>`)
	cf2 := newColorFormat(r2)
	tc2, err := cf2.ThemeColor()
	if err != nil {
		t.Fatal(err)
	}
	if tc2 == nil {
		t.Fatal("expected non-nil ThemeColor")
	}
	if *tc2 != enum.MsoThemeColorIndexAccent1 {
		t.Errorf("ThemeColor() = %v, want Accent1", *tc2)
	}
}

// Mirrors Python: it_can_change_its_theme_color (setter)
func TestColorFormat_ThemeColor_Setter(t *testing.T) {
	// Set theme color
	r := makeR(t, ``)
	cf := newColorFormat(r)
	accent2 := enum.MsoThemeColorIndexAccent2
	if err := cf.SetThemeColor(&accent2); err != nil {
		t.Fatal(err)
	}
	tc, _ := cf.ThemeColor()
	if tc == nil || *tc != enum.MsoThemeColorIndexAccent2 {
		t.Errorf("ThemeColor() = %v, want Accent2", tc)
	}

	// Type should be Theme
	ct, err := cf.Type()
	if err != nil {
		t.Fatal(err)
	}
	if ct == nil || *ct != enum.MsoColorTypeTheme {
		t.Errorf("Type() = %v, want Theme", ct)
	}

	// Remove theme color
	if err := cf.SetThemeColor(nil); err != nil {
		t.Fatal(err)
	}
}

// Mirrors Python: ColorFormat.type cases
func TestColorFormat_Type_AllCases(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		isNil    bool
		expected enum.MsoColorType
	}{
		{"none", ``, true, 0},
		{"rgb", `<w:rPr><w:color w:val="FF0000"/></w:rPr>`, false, enum.MsoColorTypeRGB},
		{"auto", `<w:rPr><w:color w:val="auto"/></w:rPr>`, false, enum.MsoColorTypeAuto},
		{"theme", `<w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr>`, false, enum.MsoColorTypeTheme},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			cf := newColorFormat(r)
			got, err := cf.Type()
			if err != nil {
				t.Fatal(err)
			}
			if tt.isNil {
				if got != nil {
					t.Errorf("Type() = %v, want nil", *got)
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil Type")
			}
			if *got != tt.expected {
				t.Errorf("Type() = %v, want %v", *got, tt.expected)
			}
		})
	}
}

// helper
func rgbPtr(r, g, b byte) *RGBColor {
	c := NewRGBColor(r, g, b)
	return &c
}
