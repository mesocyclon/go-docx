package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// parfmt_test.go — ParagraphFormat (Batch 1)
// Mirrors Python: tests/text/test_parfmt.py
// -----------------------------------------------------------------------

// Mirrors Python: it_knows_its_space_before / it_can_change
func TestParagraphFormat_SpaceBefore(t *testing.T) {
	// Getter: set
	p := makeP(t, `<w:pPr><w:spacing w:before="120"/></w:pPr>`)
	pf := newParagraph(p, nil).ParagraphFormat()
	got, err := pf.SpaceBefore()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 120 {
		t.Errorf("SpaceBefore() = %v, want 120", got)
	}

	// Getter: nil
	p2 := makeP(t, ``)
	pf2 := newParagraph(p2, nil).ParagraphFormat()
	got2, err := pf2.SpaceBefore()
	if err != nil {
		t.Fatal(err)
	}
	if got2 != nil {
		t.Errorf("SpaceBefore() = %v, want nil", got2)
	}

	// Setter
	p3 := makeP(t, ``)
	pf3 := newParagraph(p3, nil).ParagraphFormat()
	v := 200
	if err := pf3.SetSpaceBefore(&v); err != nil {
		t.Fatal(err)
	}
	got3, _ := pf3.SpaceBefore()
	if got3 == nil || *got3 != 200 {
		t.Errorf("SpaceBefore() after set = %v, want 200", got3)
	}

	// Remove
	if err := pf3.SetSpaceBefore(nil); err != nil {
		t.Fatal(err)
	}
	got4, _ := pf3.SpaceBefore()
	if got4 != nil {
		t.Errorf("SpaceBefore() after set nil = %v, want nil", got4)
	}
}

// Mirrors Python: it_knows_its_space_after / it_can_change
func TestParagraphFormat_SpaceAfter_SetAndGet(t *testing.T) {
	p := makeP(t, ``)
	pf := newParagraph(p, nil).ParagraphFormat()

	// Initially nil
	got, err := pf.SpaceAfter()
	if err != nil {
		t.Fatal(err)
	}
	if got != nil {
		t.Errorf("SpaceAfter() initially = %v, want nil", got)
	}

	// Set
	v := 240
	if err := pf.SetSpaceAfter(&v); err != nil {
		t.Fatal(err)
	}
	got2, _ := pf.SpaceAfter()
	if got2 == nil || *got2 != 240 {
		t.Errorf("SpaceAfter() after set = %v, want 240", got2)
	}
}

// Mirrors Python: it_knows_its_first_line_indent / it_can_change
func TestParagraphFormat_FirstLineIndent(t *testing.T) {
	// Getter: nil
	p := makeP(t, ``)
	pf := newParagraph(p, nil).ParagraphFormat()
	got, err := pf.FirstLineIndent()
	if err != nil {
		t.Fatal(err)
	}
	if got != nil {
		t.Errorf("FirstLineIndent() = %v, want nil", got)
	}

	// Setter
	v := 720
	if err := pf.SetFirstLineIndent(&v); err != nil {
		t.Fatal(err)
	}
	got2, _ := pf.FirstLineIndent()
	if got2 == nil || *got2 != 720 {
		t.Errorf("FirstLineIndent() after set = %v, want 720", got2)
	}

	// Remove
	if err := pf.SetFirstLineIndent(nil); err != nil {
		t.Fatal(err)
	}
	got3, _ := pf.FirstLineIndent()
	if got3 != nil {
		t.Errorf("FirstLineIndent() after nil = %v, want nil", got3)
	}
}

// Mirrors Python: it_knows_its_left_indent / it_can_change
func TestParagraphFormat_LeftIndent(t *testing.T) {
	p := makeP(t, `<w:pPr><w:ind w:left="720"/></w:pPr>`)
	pf := newParagraph(p, nil).ParagraphFormat()
	got, err := pf.LeftIndent()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 720 {
		t.Errorf("LeftIndent() = %v, want 720", got)
	}

	// Set to different
	v := 1440
	if err := pf.SetLeftIndent(&v); err != nil {
		t.Fatal(err)
	}
	got2, _ := pf.LeftIndent()
	if got2 == nil || *got2 != 1440 {
		t.Errorf("LeftIndent() after set = %v, want 1440", got2)
	}
}

// Mirrors Python: it_knows_its_right_indent / it_can_change
func TestParagraphFormat_RightIndent(t *testing.T) {
	p := makeP(t, ``)
	pf := newParagraph(p, nil).ParagraphFormat()

	// Nil initially
	got, err := pf.RightIndent()
	if err != nil {
		t.Fatal(err)
	}
	if got != nil {
		t.Errorf("RightIndent() = %v, want nil", got)
	}

	v := 360
	if err := pf.SetRightIndent(&v); err != nil {
		t.Fatal(err)
	}
	got2, _ := pf.RightIndent()
	if got2 == nil || *got2 != 360 {
		t.Errorf("RightIndent() after set = %v, want 360", got2)
	}
}

// Mirrors Python: it_knows_its_line_spacing (complete cases)
func TestParagraphFormat_LineSpacing(t *testing.T) {
	tests := []struct {
		name       string
		innerXml   string
		isNil      bool
		isMultiple bool
		multiple   float64
		twips      int
	}{
		{"nil_when_absent", ``, true, false, 0, 0},
		{"double_spacing", `<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>`, false, true, 2.0, 0},
		{"single_spacing", `<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>`, false, true, 1.0, 0},
		{"exact_240", `<w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>`, false, false, 0, 240},
		{"at_least_300", `<w:pPr><w:spacing w:line="300" w:lineRule="atLeast"/></w:pPr>`, false, false, 0, 300},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			pf := newParagraph(p, nil).ParagraphFormat()
			got, err := pf.LineSpacing()
			if err != nil {
				t.Fatal(err)
			}
			if tt.isNil {
				if got != nil {
					t.Error("expected nil LineSpacing")
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil LineSpacing")
			}
			if got.IsMultiple() != tt.isMultiple {
				t.Errorf("IsMultiple() = %v, want %v", got.IsMultiple(), tt.isMultiple)
			}
			if tt.isMultiple {
				if diff := got.Multiple() - tt.multiple; diff > 0.01 || diff < -0.01 {
					t.Errorf("Multiple() = %f, want %f", got.Multiple(), tt.multiple)
				}
			} else {
				if got.Twips() != tt.twips {
					t.Errorf("Twips() = %d, want %d", got.Twips(), tt.twips)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_its_line_spacing_rule
func TestParagraphFormat_LineSpacingRule(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		isNil    bool
		expected enum.WdLineSpacing
	}{
		{"nil_when_absent", ``, true, 0},
		{"single", `<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>`, false, enum.WdLineSpacingSingle},
		{"one_point_five", `<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>`, false, enum.WdLineSpacingOnePointFive},
		{"double", `<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>`, false, enum.WdLineSpacingDouble},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			pf := newParagraph(p, nil).ParagraphFormat()
			got, err := pf.LineSpacingRule()
			if err != nil {
				t.Fatal(err)
			}
			if tt.isNil {
				if got != nil {
					t.Errorf("LineSpacingRule() = %v, want nil", *got)
				}
				return
			}
			if got == nil {
				t.Fatal("expected non-nil LineSpacingRule")
			}
			if *got != tt.expected {
				t.Errorf("LineSpacingRule() = %v, want %v", *got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_can_change_its_line_spacing_rule
func TestParagraphFormat_SetLineSpacingRule(t *testing.T) {
	tests := []struct {
		name string
		rule enum.WdLineSpacing
	}{
		{"single", enum.WdLineSpacingSingle},
		{"one_point_five", enum.WdLineSpacingOnePointFive},
		{"double", enum.WdLineSpacingDouble},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, ``)
			pf := newParagraph(p, nil).ParagraphFormat()
			if err := pf.SetLineSpacingRule(tt.rule); err != nil {
				t.Fatal(err)
			}
			got, err := pf.LineSpacingRule()
			if err != nil {
				t.Fatal(err)
			}
			if got == nil || *got != tt.rule {
				t.Errorf("LineSpacingRule() after set = %v, want %v", got, tt.rule)
			}
		})
	}
}

// Mirrors Python: it_knows_its_on_off_prop_values — KeepTogether, KeepWithNext, PageBreakBefore, WidowControl
func TestParagraphFormat_OnOffProps(t *testing.T) {
	type onOffProp struct {
		name   string
		xmlTag string
		get    func(*ParagraphFormat) *bool
		set    func(*ParagraphFormat, *bool) error
	}
	props := []onOffProp{
		{"KeepTogether", "keepLines", (*ParagraphFormat).KeepTogether, (*ParagraphFormat).SetKeepTogether},
		{"KeepWithNext", "keepNext", (*ParagraphFormat).KeepWithNext, (*ParagraphFormat).SetKeepWithNext},
		{"PageBreakBefore", "pageBreakBefore", (*ParagraphFormat).PageBreakBefore, (*ParagraphFormat).SetPageBreakBefore},
		{"WidowControl", "widowControl", (*ParagraphFormat).WidowControl, (*ParagraphFormat).SetWidowControl},
	}

	for _, prop := range props {
		t.Run(prop.name, func(t *testing.T) {
			// absent → nil
			t.Run("absent", func(t *testing.T) {
				p := makeP(t, ``)
				pf := newParagraph(p, nil).ParagraphFormat()
				compareBoolPtr(t, prop.name, prop.get(pf), nil)
			})
			// present → true
			t.Run("present_true", func(t *testing.T) {
				p := makeP(t, `<w:pPr><w:`+prop.xmlTag+`/></w:pPr>`)
				pf := newParagraph(p, nil).ParagraphFormat()
				compareBoolPtr(t, prop.name, prop.get(pf), boolPtr(true))
			})
			// set to true
			t.Run("set_true", func(t *testing.T) {
				p := makeP(t, ``)
				pf := newParagraph(p, nil).ParagraphFormat()
				if err := prop.set(pf, boolPtr(true)); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(pf), boolPtr(true))
			})
			// set to nil
			t.Run("set_nil", func(t *testing.T) {
				p := makeP(t, `<w:pPr><w:`+prop.xmlTag+`/></w:pPr>`)
				pf := newParagraph(p, nil).ParagraphFormat()
				if err := prop.set(pf, nil); err != nil {
					t.Fatal(err)
				}
				compareBoolPtr(t, prop.name, prop.get(pf), nil)
			})
		})
	}
}

// Mirrors Python: it_provides_access_to_its_tab_stops
func TestParagraphFormat_TabStops(t *testing.T) {
	p := makeP(t, `<w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>`)
	pf := newParagraph(p, nil).ParagraphFormat()
	ts := pf.TabStops()
	if ts == nil {
		t.Fatal("TabStops() returned nil")
	}
	if ts.Len() != 1 {
		t.Errorf("TabStops.Len() = %d, want 1", ts.Len())
	}
}
