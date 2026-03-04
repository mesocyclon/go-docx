package enum

import "testing"

// ---------------------------------------------------------------------------
// WdParagraphAlignment
// ---------------------------------------------------------------------------

func TestWdParagraphAlignmentFromXml(t *testing.T) {
	t.Parallel()
	tests := []struct {
		xml  string
		want WdParagraphAlignment
	}{
		{"left", WdParagraphAlignmentLeft},
		{"center", WdParagraphAlignmentCenter},
		{"right", WdParagraphAlignmentRight},
		{"both", WdParagraphAlignmentJustify},
		{"distribute", WdParagraphAlignmentDistribute},
		{"mediumKashida", WdParagraphAlignmentJustifyMed},
		{"highKashida", WdParagraphAlignmentJustifyHi},
		{"lowKashida", WdParagraphAlignmentJustifyLow},
		{"thaiDistribute", WdParagraphAlignmentThaiJustify},
	}
	for _, tc := range tests {
		t.Run(tc.xml, func(t *testing.T) {
			t.Parallel()
			got, err := WdParagraphAlignmentFromXml(tc.xml)
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if got != tc.want {
				t.Errorf("WdParagraphAlignmentFromXml(%q) = %d, want %d", tc.xml, got, tc.want)
			}
		})
	}
}

func TestWdParagraphAlignmentToXml(t *testing.T) {
	t.Parallel()
	// JUSTIFY maps to "both", not "justify"
	got, err := WdParagraphAlignmentJustify.ToXml()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != "both" {
		t.Errorf("JUSTIFY.ToXml() = %q, want %q", got, "both")
	}
	got, err = WdParagraphAlignmentCenter.ToXml()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != "center" {
		t.Errorf("CENTER.ToXml() = %q, want %q", got, "center")
	}
}

func TestWdParagraphAlignmentRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdParagraphAlignmentToXml {
		got, err := WdParagraphAlignmentFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

func TestWdParagraphAlignmentFromXmlError(t *testing.T) {
	t.Parallel()
	_, err := WdParagraphAlignmentFromXml("nonexistent")
	if err == nil {
		t.Error("expected error for nonexistent XML value, got nil")
	}
}

// ---------------------------------------------------------------------------
// MsoThemeColorIndex
// ---------------------------------------------------------------------------

func TestMsoThemeColorIndexFromXml(t *testing.T) {
	t.Parallel()
	got, err := MsoThemeColorIndexFromXml("accent1")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != MsoThemeColorIndexAccent1 {
		t.Errorf("MsoThemeColorIndexFromXml(\"accent1\") = %d, want %d", got, MsoThemeColorIndexAccent1)
	}
}

func TestMsoThemeColorIndexToXml(t *testing.T) {
	t.Parallel()
	got, err := MsoThemeColorIndexAccent1.ToXml()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != "accent1" {
		t.Errorf("ACCENT_1.ToXml() = %q, want %q", got, "accent1")
	}
}

func TestMsoThemeColorIndexNotThemeColorNoXml(t *testing.T) {
	t.Parallel()
	_, err := MsoThemeColorIndexNotThemeColor.ToXml()
	if err == nil {
		t.Error("expected error for NOT_THEME_COLOR.ToXml(), got nil")
	}
}

func TestMsoThemeColorIndexRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range msoThemeColorIndexToXml {
		got, err := MsoThemeColorIndexFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

// ---------------------------------------------------------------------------
// WdColorIndex
// ---------------------------------------------------------------------------

func TestWdColorIndexRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdColorIndexToXml {
		got, err := WdColorIndexFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

// ---------------------------------------------------------------------------
// WdLineSpacing
// ---------------------------------------------------------------------------

func TestWdLineSpacingFromXml(t *testing.T) {
	t.Parallel()
	tests := []struct {
		xml  string
		want WdLineSpacing
	}{
		{"atLeast", WdLineSpacingAtLeast},
		{"exact", WdLineSpacingExactly},
		{"auto", WdLineSpacingMultiple},
	}
	for _, tc := range tests {
		t.Run(tc.xml, func(t *testing.T) {
			t.Parallel()
			got, err := WdLineSpacingFromXml(tc.xml)
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if got != tc.want {
				t.Errorf("got %d, want %d", got, tc.want)
			}
		})
	}
}

func TestWdLineSpacingUnmappedToXml(t *testing.T) {
	t.Parallel()
	// SINGLE, ONE_POINT_FIVE, DOUBLE have no XML mapping
	_, err := WdLineSpacingSingle.ToXml()
	if err == nil {
		t.Error("expected error for SINGLE.ToXml(), got nil")
	}
}

// ---------------------------------------------------------------------------
// WdTabAlignment
// ---------------------------------------------------------------------------

func TestWdTabAlignmentRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdTabAlignmentToXml {
		got, err := WdTabAlignmentFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

// ---------------------------------------------------------------------------
// WdTabLeader
// ---------------------------------------------------------------------------

func TestWdTabLeaderRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdTabLeaderToXml {
		got, err := WdTabLeaderFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

// ---------------------------------------------------------------------------
// WdUnderline
// ---------------------------------------------------------------------------

func TestWdUnderlineFromXml(t *testing.T) {
	t.Parallel()
	tests := []struct {
		xml  string
		want WdUnderline
	}{
		{"none", WdUnderlineNone},
		{"single", WdUnderlineSingle},
		{"double", WdUnderlineDouble},
		{"wave", WdUnderlineWavy},
		{"wavyDouble", WdUnderlineWavyDouble},
		{"dashLongHeavy", WdUnderlineDashLongHeavy},
	}
	for _, tc := range tests {
		t.Run(tc.xml, func(t *testing.T) {
			t.Parallel()
			got, err := WdUnderlineFromXml(tc.xml)
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if got != tc.want {
				t.Errorf("got %d, want %d", got, tc.want)
			}
		})
	}
}

func TestWdUnderlineInheritedNoXml(t *testing.T) {
	t.Parallel()
	_, err := WdUnderlineInherited.ToXml()
	if err == nil {
		t.Error("expected error for INHERITED.ToXml(), got nil")
	}
}

func TestWdUnderlineRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdUnderlineToXml {
		got, err := WdUnderlineFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

// ---------------------------------------------------------------------------
// Section enums
// ---------------------------------------------------------------------------

func TestWdHeaderFooterIndexRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdHeaderFooterIndexToXml {
		got, err := WdHeaderFooterIndexFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q", xml)
		}
	}
}

func TestWdOrientationRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdOrientationToXml {
		got, err := WdOrientationFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q", xml)
		}
	}
}

func TestWdSectionStartRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdSectionStartToXml {
		got, err := WdSectionStartFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q", xml)
		}
	}
}

// ---------------------------------------------------------------------------
// Style enums
// ---------------------------------------------------------------------------

func TestWdStyleTypeRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdStyleTypeToXml {
		got, err := WdStyleTypeFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q, got=%d, want=%d", xml, got, val)
		}
	}
}

func TestWdStyleTypeValues(t *testing.T) {
	t.Parallel()
	// Verify exact values match Python source
	if WdStyleTypeParagraph != 1 {
		t.Errorf("PARAGRAPH = %d, want 1", WdStyleTypeParagraph)
	}
	if WdStyleTypeCharacter != 2 {
		t.Errorf("CHARACTER = %d, want 2", WdStyleTypeCharacter)
	}
	if WdStyleTypeTable != 3 {
		t.Errorf("TABLE = %d, want 3", WdStyleTypeTable)
	}
	if WdStyleTypeList != 4 {
		t.Errorf("LIST = %d, want 4", WdStyleTypeList)
	}
}

// ---------------------------------------------------------------------------
// Table enums
// ---------------------------------------------------------------------------

func TestWdTableAlignmentRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdTableAlignmentToXml {
		got, err := WdTableAlignmentFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q", xml)
		}
	}
}

func TestWdCellVerticalAlignmentRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdCellVerticalAlignmentToXml {
		got, err := WdCellVerticalAlignmentFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q", xml)
		}
	}
}

func TestWdRowHeightRuleRoundTrip(t *testing.T) {
	t.Parallel()
	for val, xml := range wdRowHeightRuleToXml {
		got, err := WdRowHeightRuleFromXml(xml)
		if err != nil {
			t.Fatalf("round-trip error for %q: %v", xml, err)
		}
		if got != val {
			t.Errorf("round-trip failed: xml=%q", xml)
		}
	}
}

func TestWdRowHeightRuleExactly(t *testing.T) {
	t.Parallel()
	// EXACTLY maps to "exact" (not "exactly")
	got, err := WdRowHeightRuleExactly.ToXml()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != "exact" {
		t.Errorf("EXACTLY.ToXml() = %q, want %q", got, "exact")
	}
}

// ---------------------------------------------------------------------------
// Shape enums (no XML mapping, just verify values)
// ---------------------------------------------------------------------------

func TestWdInlineShapeTypeValues(t *testing.T) {
	t.Parallel()
	if WdInlineShapeTypeChart != 12 {
		t.Errorf("CHART = %d, want 12", WdInlineShapeTypeChart)
	}
	if WdInlineShapeTypePicture != 3 {
		t.Errorf("PICTURE = %d, want 3", WdInlineShapeTypePicture)
	}
	if WdInlineShapeTypeNotImplemented != -6 {
		t.Errorf("NOT_IMPLEMENTED = %d, want -6", WdInlineShapeTypeNotImplemented)
	}
}

// ---------------------------------------------------------------------------
// Generic FromXml error
// ---------------------------------------------------------------------------

func TestFromXmlError(t *testing.T) {
	t.Parallel()
	_, err := FromXml(map[string]int{"a": 1}, "nonexistent")
	if err == nil {
		t.Error("expected error for nonexistent key, got nil")
	}
}

// ---------------------------------------------------------------------------
// Unmapped enum value error tests (verifies unified ToXml behavior)
// ---------------------------------------------------------------------------

func TestToXmlReturnsErrorForUnmappedValues(t *testing.T) {
	t.Parallel()

	// WdColorIndex.INHERITED has no XML mapping
	_, err := WdColorIndexInherited.ToXml()
	if err == nil {
		t.Error("expected error for WdColorIndexInherited.ToXml(), got nil")
	}

	// An arbitrary invalid int cast to a fully-mapped enum should also error
	_, err = WdHeaderFooterIndex(999).ToXml()
	if err == nil {
		t.Error("expected error for WdHeaderFooterIndex(999).ToXml(), got nil")
	}

	_, err = WdOrientation(999).ToXml()
	if err == nil {
		t.Error("expected error for WdOrientation(999).ToXml(), got nil")
	}

	_, err = WdStyleType(999).ToXml()
	if err == nil {
		t.Error("expected error for WdStyleType(999).ToXml(), got nil")
	}

	_, err = WdParagraphAlignment(999).ToXml()
	if err == nil {
		t.Error("expected error for WdParagraphAlignment(999).ToXml(), got nil")
	}
}
