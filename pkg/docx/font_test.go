package docx

import (
	"testing"
)

// -----------------------------------------------------------------------
// font_test.go — Font (Batch 1)
// Mirrors Python: tests/text/test_font.py
// -----------------------------------------------------------------------

// fontBoolPropGetter is a func that reads a bool prop from Font.
type fontBoolPropGetter func(*Font) *bool

// fontBoolPropSetter is a func that writes a bool prop to Font.
type fontBoolPropSetter func(*Font, *bool) error

// fontBoolPropPair bundles get/set for table-driven tests.
type fontBoolPropPair struct {
	name   string
	xmlTag string // XML element name (e.g. "b", "i", "caps")
	get    fontBoolPropGetter
	set    fontBoolPropSetter
}

var fontBoolProps = []fontBoolPropPair{
	{"AllCaps", "caps", (*Font).AllCaps, (*Font).SetAllCaps},
	{"Bold", "b", (*Font).Bold, (*Font).SetBold},
	{"ComplexScript", "cs", (*Font).ComplexScript, (*Font).SetComplexScript},
	{"CsBold", "bCs", (*Font).CsBold, (*Font).SetCsBold},
	{"CsItalic", "iCs", (*Font).CsItalic, (*Font).SetCsItalic},
	{"DoubleStrike", "dstrike", (*Font).DoubleStrike, (*Font).SetDoubleStrike},
	{"Emboss", "emboss", (*Font).Emboss, (*Font).SetEmboss},
	{"Hidden", "vanish", (*Font).Hidden, (*Font).SetHidden},
	{"Italic", "i", (*Font).Italic, (*Font).SetItalic},
	{"Imprint", "imprint", (*Font).Imprint, (*Font).SetImprint},
	{"Math", "oMath", (*Font).Math, (*Font).SetMath},
	{"NoProof", "noProof", (*Font).NoProof, (*Font).SetNoProof},
	{"Outline", "outline", (*Font).Outline, (*Font).SetOutline},
	{"Rtl", "rtl", (*Font).Rtl, (*Font).SetRtl},
	{"Shadow", "shadow", (*Font).Shadow, (*Font).SetShadow},
	{"SmallCaps", "smallCaps", (*Font).SmallCaps, (*Font).SetSmallCaps},
	{"SnapToGrid", "snapToGrid", (*Font).SnapToGrid, (*Font).SetSnapToGrid},
	{"SpecVanish", "specVanish", (*Font).SpecVanish, (*Font).SetSpecVanish},
	{"Strike", "strike", (*Font).Strike, (*Font).SetStrike},
	{"WebHidden", "webHidden", (*Font).WebHidden, (*Font).SetWebHidden},
}

// Mirrors Python: it_knows_its_bool_prop_states — tests all 20 bool properties.
// Each property is tested with: absent (nil), present no val (true), val="on" (true),
// val="off" (false), val="1" (true), val="0" (false).
func TestFont_BoolPropGetters_AllProps(t *testing.T) {
	for _, prop := range fontBoolProps {
		t.Run(prop.name, func(t *testing.T) {
			// Case 1: absent → nil
			t.Run("absent", func(t *testing.T) {
				r := makeR(t, `<w:rPr/>`)
				f := newFont(r)
				compareBoolPtr(t, prop.name+"()", prop.get(f), nil)
			})
			// Case 2: present no val → true
			t.Run("default_true", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+`/></w:rPr>`)
				f := newFont(r)
				compareBoolPtr(t, prop.name+"()", prop.get(f), boolPtr(true))
			})
			// Case 3: val="on" → true
			t.Run("val_on", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+` w:val="on"/></w:rPr>`)
				f := newFont(r)
				compareBoolPtr(t, prop.name+"()", prop.get(f), boolPtr(true))
			})
			// Case 4: val="off" → false
			t.Run("val_off", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+` w:val="off"/></w:rPr>`)
				f := newFont(r)
				compareBoolPtr(t, prop.name+"()", prop.get(f), boolPtr(false))
			})
			// Case 5: val="1" → true
			t.Run("val_1", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+` w:val="1"/></w:rPr>`)
				f := newFont(r)
				compareBoolPtr(t, prop.name+"()", prop.get(f), boolPtr(true))
			})
			// Case 6: val="0" → false
			t.Run("val_0", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+` w:val="0"/></w:rPr>`)
				f := newFont(r)
				compareBoolPtr(t, prop.name+"()", prop.get(f), boolPtr(false))
			})
		})
	}
}

// Mirrors Python: it_can_change_its_bool_prop_settings — tests all 20 bool setters.
// Each has 4 transitions: empty→true, empty→false, empty→nil, and one val→opposite.
func TestFont_BoolPropSetters_AllProps(t *testing.T) {
	for _, prop := range fontBoolProps {
		t.Run(prop.name, func(t *testing.T) {
			// empty → true
			t.Run("empty_to_true", func(t *testing.T) {
				r := makeR(t, "")
				f := newFont(r)
				if err := prop.set(f, boolPtr(true)); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(f), boolPtr(true))
			})
			// empty → false
			t.Run("empty_to_false", func(t *testing.T) {
				r := makeR(t, "")
				f := newFont(r)
				if err := prop.set(f, boolPtr(false)); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(f), boolPtr(false))
			})
			// empty → nil
			t.Run("empty_to_nil", func(t *testing.T) {
				r := makeR(t, "")
				f := newFont(r)
				if err := prop.set(f, nil); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(f), nil)
			})
			// default(true) → false
			t.Run("true_to_false", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+`/></w:rPr>`)
				f := newFont(r)
				if err := prop.set(f, boolPtr(false)); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(f), boolPtr(false))
			})
			// default(true) → nil
			t.Run("true_to_nil", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+`/></w:rPr>`)
				f := newFont(r)
				if err := prop.set(f, nil); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(f), nil)
			})
			// false → true
			t.Run("false_to_true", func(t *testing.T) {
				r := makeR(t, `<w:rPr><w:`+prop.xmlTag+` w:val="0"/></w:rPr>`)
				f := newFont(r)
				if err := prop.set(f, boolPtr(true)); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(f), boolPtr(true))
			})
		})
	}
}

// Mirrors Python: it_knows_its_typeface_name (4 cases)
func TestFont_Name_Getter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected *string
	}{
		{"absent", ``, nil},
		{"rPr_only", `<w:rPr/>`, nil},
		{"rFonts_no_ascii", `<w:rPr><w:rFonts/></w:rPr>`, nil},
		{"arial", `<w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>`, strPtr("Arial")},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			f := newFont(r)
			got := f.Name()
			if got == nil && tt.expected == nil {
				return
			}
			if got == nil || tt.expected == nil || *got != *tt.expected {
				t.Errorf("Name() = %v, want %v", ptrStr(got), ptrStr(tt.expected))
			}
		})
	}
}

// Mirrors Python: it_can_change_its_typeface_name
func TestFont_Name_Setter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		value    *string
	}{
		{"set_on_empty", ``, strPtr("Foo")},
		{"set_on_rPr", `<w:rPr/>`, strPtr("Foo")},
		{"change_existing", `<w:rPr><w:rFonts w:ascii="Foo" w:hAnsi="Foo"/></w:rPr>`, strPtr("Bar")},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			f := newFont(r)
			if err := f.SetName(tt.value); err != nil {
				t.Fatal(err)
			}
			got := f.Name()
			if tt.value == nil {
				if got != nil {
					t.Errorf("Name() = %v after set nil, want nil", *got)
				}
			} else {
				if got == nil || *got != *tt.value {
					t.Errorf("Name() = %v, want %v", ptrStr(got), *tt.value)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_its_size (3 cases)
func TestFont_Size_Getter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		isNil    bool
		wantPt   float64
	}{
		{"absent", ``, true, 0},
		{"rPr_only", `<w:rPr/>`, true, 0},
		{"sz_28_halfpts_eq_14pt", `<w:rPr><w:sz w:val="28"/></w:rPr>`, false, 14.0},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			f := newFont(r)
			got, err := f.Size()
			if err != nil {
				t.Fatal(err)
			}
			if tt.isNil {
				if got != nil {
					t.Errorf("Size() = %v, want nil", got)
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil Size()")
			}
			if diff := got.Pt() - tt.wantPt; diff > 0.01 || diff < -0.01 {
				t.Errorf("Size().Pt() = %f, want %f", got.Pt(), tt.wantPt)
			}
		})
	}
}

// Mirrors Python: it_can_change_its_size (4 cases)
func TestFont_Size_Setter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		value    *Length
		wantPt   float64
		wantNil  bool
	}{
		{"set_12pt_on_empty", ``, ptPtr(12), 12.0, false},
		{"set_12pt_on_rPr", `<w:rPr/>`, ptPtr(12), 12.0, false},
		{"change_12_to_18", `<w:rPr><w:sz w:val="24"/></w:rPr>`, ptPtr(18), 18.0, false},
		{"remove", `<w:rPr><w:sz w:val="36"/></w:rPr>`, nil, 0, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			f := newFont(r)
			if err := f.SetSize(tt.value); err != nil {
				t.Fatal(err)
			}
			got, err := f.Size()
			if err != nil {
				t.Fatal(err)
			}
			if tt.wantNil {
				if got != nil {
					t.Errorf("Size() = %v after set nil, want nil", got)
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil Size()")
			}
			if diff := got.Pt() - tt.wantPt; diff > 0.5 || diff < -0.5 {
				t.Errorf("Size().Pt() = %f, want ~%f", got.Pt(), tt.wantPt)
			}
		})
	}
}

// Mirrors Python: it_provides_access_to_its_color_object
func TestFont_Color(t *testing.T) {
	r := makeR(t, `<w:rPr><w:color w:val="FF0000"/></w:rPr>`)
	f := newFont(r)
	c := f.Color()
	if c == nil {
		t.Fatal("Color() returned nil")
	}
	rgb, err := c.RGB()
	if err != nil {
		t.Fatal(err)
	}
	if rgb == nil || rgb.String() != "FF0000" {
		t.Errorf("Color().RGB() = %v, want FF0000", rgb)
	}
}

// Mirrors Python: it_knows_its_underline_type
func TestFont_Underline_Getter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		isNil    bool
		isSingle bool
		isNone   bool
	}{
		{"absent", ``, true, false, false},
		{"single", `<w:rPr><w:u w:val="single"/></w:rPr>`, false, true, false},
		{"none_explicit", `<w:rPr><w:u w:val="none"/></w:rPr>`, false, false, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			f := newFont(r)
			got, err := f.Underline()
			if err != nil {
				t.Fatal(err)
			}
			if tt.isNil {
				if got != nil {
					t.Error("expected nil Underline")
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil Underline")
			}
			if got.IsSingle() != tt.isSingle {
				t.Errorf("IsSingle() = %v, want %v", got.IsSingle(), tt.isSingle)
			}
			if got.IsNone() != tt.isNone {
				t.Errorf("IsNone() = %v, want %v", got.IsNone(), tt.isNone)
			}
		})
	}
}

// Mirrors Python: it_can_change_its_underline_type
func TestFont_Underline_Setter(t *testing.T) {
	// set single
	r := makeR(t, "")
	f := newFont(r)
	single := UnderlineSingle()
	if err := f.SetUnderline(&single); err != nil {
		t.Fatal(err)
	}
	got, err := f.Underline()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || !got.IsSingle() {
		t.Error("expected single underline")
	}
	// set nil
	if err := f.SetUnderline(nil); err != nil {
		t.Fatal(err)
	}
	got2, err := f.Underline()
	if err != nil {
		t.Fatal(err)
	}
	if got2 != nil {
		t.Error("expected nil underline after set nil")
	}
}

// helper: ptPtr returns a *Length for a point value.
func ptPtr(v float64) *Length {
	l := Pt(v)
	return &l
}
