package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// font_extended_test.go — Font subscript, superscript, highlight (Batch 2)
// Mirrors Python: tests/text/test_font.py — subscript/superscript/highlight
// -----------------------------------------------------------------------

// Mirrors Python: it_knows_whether_it_is_subscript (5 cases)
func TestFont_Subscript_Getter(t *testing.T) {
	tests := []struct {
		name     string
		rprXml   string
		expected *bool
	}{
		{"no_rPr", "", nil},
		{"empty_rPr", "<w:rPr/>", nil},
		{"baseline", `<w:rPr><w:vertAlign w:val="baseline"/></w:rPr>`, boolPtr(false)},
		{"subscript", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, boolPtr(true)},
		{"superscript", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, boolPtr(false)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.rprXml)
			run := newRun(r, nil)
			font := run.Font()

			got, err := font.Subscript()
			if err != nil {
				t.Fatalf("Subscript(): %v", err)
			}
			compareBoolPtr(t, "Subscript()", got, tt.expected)
		})
	}
}

// Mirrors Python: it_can_change_whether_it_is_subscript (10 cases)
func TestFont_Subscript_Setter(t *testing.T) {
	tests := []struct {
		name     string
		rprXml   string
		value    *bool
		wantAttr string // expected w:val on vertAlign, or "" if vertAlign removed
		wantGone bool   // true if vertAlign should be absent
	}{
		// from empty
		{"empty_to_true", "", boolPtr(true), "subscript", false},
		{"empty_to_false", "", boolPtr(false), "", true},
		{"empty_to_nil", "", nil, "", true},
		// from subscript
		{"sub_to_true", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, boolPtr(true), "subscript", false},
		{"sub_to_false", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, boolPtr(false), "", true},
		{"sub_to_nil", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, nil, "", true},
		// from superscript
		{"super_to_true", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, boolPtr(true), "subscript", false},
		{"super_to_false", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, boolPtr(false), "superscript", false},
		{"super_to_nil", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, nil, "", true},
		// from baseline
		{"baseline_to_true", `<w:rPr><w:vertAlign w:val="baseline"/></w:rPr>`, boolPtr(true), "subscript", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.rprXml)
			run := newRun(r, nil)
			font := run.Font()

			if err := font.SetSubscript(tt.value); err != nil {
				t.Fatalf("SetSubscript: %v", err)
			}

			// Verify via getter round-trip
			got, err := font.Subscript()
			if err != nil {
				t.Fatalf("Subscript() after set: %v", err)
			}
			if tt.value != nil && *tt.value {
				if got == nil || !*got {
					t.Errorf("after SetSubscript(true): got %s", ptrBoolStr(got))
				}
			}

			// Verify XML
			rPr := r.RPr()
			if rPr == nil {
				if !tt.wantGone {
					t.Error("rPr is nil but expected vertAlign")
				}
				return
			}
			va := rPr.RawElement().FindElement("vertAlign")
			if tt.wantGone {
				if va != nil {
					t.Error("expected vertAlign to be removed")
				}
			} else {
				if va == nil {
					t.Fatal("expected vertAlign element")
				}
				gotVal := va.SelectAttrValue("w:val", "")
				if gotVal != tt.wantAttr {
					t.Errorf("vertAlign w:val = %q, want %q", gotVal, tt.wantAttr)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_whether_it_is_superscript (5 cases)
func TestFont_Superscript_Getter(t *testing.T) {
	tests := []struct {
		name     string
		rprXml   string
		expected *bool
	}{
		{"no_rPr", "", nil},
		{"empty_rPr", "<w:rPr/>", nil},
		{"baseline", `<w:rPr><w:vertAlign w:val="baseline"/></w:rPr>`, boolPtr(false)},
		{"subscript", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, boolPtr(false)},
		{"superscript", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, boolPtr(true)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.rprXml)
			run := newRun(r, nil)
			font := run.Font()

			got, err := font.Superscript()
			if err != nil {
				t.Fatalf("Superscript(): %v", err)
			}
			compareBoolPtr(t, "Superscript()", got, tt.expected)
		})
	}
}

// Mirrors Python: it_can_change_whether_it_is_superscript (10 cases)
func TestFont_Superscript_Setter(t *testing.T) {
	tests := []struct {
		name     string
		rprXml   string
		value    *bool
		wantAttr string
		wantGone bool
	}{
		// from empty
		{"empty_to_true", "", boolPtr(true), "superscript", false},
		{"empty_to_false", "", boolPtr(false), "", true},
		{"empty_to_nil", "", nil, "", true},
		// from superscript
		{"super_to_true", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, boolPtr(true), "superscript", false},
		{"super_to_false", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, boolPtr(false), "", true},
		{"super_to_nil", `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`, nil, "", true},
		// from subscript
		{"sub_to_true", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, boolPtr(true), "superscript", false},
		{"sub_to_false", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, boolPtr(false), "subscript", false},
		{"sub_to_nil", `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`, nil, "", true},
		// from baseline
		{"baseline_to_true", `<w:rPr><w:vertAlign w:val="baseline"/></w:rPr>`, boolPtr(true), "superscript", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.rprXml)
			run := newRun(r, nil)
			font := run.Font()

			if err := font.SetSuperscript(tt.value); err != nil {
				t.Fatalf("SetSuperscript: %v", err)
			}

			rPr := r.RPr()
			if rPr == nil {
				if !tt.wantGone {
					t.Error("rPr is nil but expected vertAlign")
				}
				return
			}
			va := rPr.RawElement().FindElement("vertAlign")
			if tt.wantGone {
				if va != nil {
					t.Error("expected vertAlign to be removed")
				}
			} else {
				if va == nil {
					t.Fatal("expected vertAlign element")
				}
				gotVal := va.SelectAttrValue("w:val", "")
				if gotVal != tt.wantAttr {
					t.Errorf("vertAlign w:val = %q, want %q", gotVal, tt.wantAttr)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_its_highlight_color (4 cases)
func TestFont_HighlightColor_Getter(t *testing.T) {
	tests := []struct {
		name     string
		rprXml   string
		expected *enum.WdColorIndex
	}{
		{"no_rPr", "", nil},
		{"empty_rPr", "<w:rPr/>", nil},
		{"auto", `<w:rPr><w:highlight w:val="default"/></w:rPr>`, colorIndexPtr(enum.WdColorIndexAuto)},
		{"blue", `<w:rPr><w:highlight w:val="blue"/></w:rPr>`, colorIndexPtr(enum.WdColorIndexBlue)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.rprXml)
			run := newRun(r, nil)
			font := run.Font()

			got, err := font.HighlightColor()
			if err != nil {
				t.Fatalf("HighlightColor(): %v", err)
			}
			if tt.expected == nil {
				if got != nil {
					t.Errorf("HighlightColor() = %v, want nil", *got)
				}
			} else {
				if got == nil {
					t.Fatalf("HighlightColor() = nil, want %v", *tt.expected)
				}
				if *got != *tt.expected {
					t.Errorf("HighlightColor() = %v, want %v", *got, *tt.expected)
				}
			}
		})
	}
}

// Mirrors Python: it_can_change_its_highlight_color (6 cases)
func TestFont_HighlightColor_Setter(t *testing.T) {
	tests := []struct {
		name     string
		rprXml   string
		value    *enum.WdColorIndex
		wantAttr string // expected w:val on highlight, or "" if removed
		wantGone bool
	}{
		{"empty_to_auto", "", colorIndexPtr(enum.WdColorIndexAuto), "default", false},
		{"rPr_to_green", "<w:rPr/>", colorIndexPtr(enum.WdColorIndexBrightGreen), "green", false},
		{"green_to_yellow", `<w:rPr><w:highlight w:val="green"/></w:rPr>`, colorIndexPtr(enum.WdColorIndexYellow), "yellow", false},
		{"yellow_to_nil", `<w:rPr><w:highlight w:val="yellow"/></w:rPr>`, nil, "", true},
		{"rPr_to_nil", "<w:rPr/>", nil, "", true},
		{"empty_to_nil", "", nil, "", true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.rprXml)
			run := newRun(r, nil)
			font := run.Font()

			if err := font.SetHighlightColor(tt.value); err != nil {
				t.Fatalf("SetHighlightColor: %v", err)
			}

			rPr := r.RPr()
			if rPr == nil {
				if !tt.wantGone {
					t.Error("rPr is nil but expected highlight")
				}
				return
			}
			hl := rPr.RawElement().FindElement("highlight")
			if tt.wantGone {
				if hl != nil {
					t.Error("expected highlight to be removed")
				}
			} else {
				if hl == nil {
					t.Fatal("expected highlight element")
				}
				gotVal := hl.SelectAttrValue("w:val", "")
				if gotVal != tt.wantAttr {
					t.Errorf("highlight w:val = %q, want %q", gotVal, tt.wantAttr)
				}
			}
		})
	}
}

// --- helpers ---

func colorIndexPtr(v enum.WdColorIndex) *enum.WdColorIndex { return &v }
