package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func TestStyleIdFromName(t *testing.T) {
	cases := []struct {
		name, expected string
	}{
		{"Heading 1", "Heading1"},
		{"heading 1", "Heading1"},
		{"caption", "Caption"},
		{"Normal", "Normal"},
		{"Table of Contents", "TableofContents"},
		{"Body Text", "BodyText"},
	}
	for _, c := range cases {
		if got := StyleIdFromName(c.name); got != c.expected {
			t.Errorf("StyleIdFromName(%q) = %q, want %q", c.name, got, c.expected)
		}
	}
}

func TestCT_Styles_GetByID(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("Heading1"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := s.SetNameVal("heading 1"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}

	found := styles.GetByID("Heading1")
	if found == nil {
		t.Fatal("expected to find style by ID")
	}
	if nv, err := found.NameVal(); err != nil {
		t.Fatalf("NameVal: %v", err)
	} else if nv != "heading 1" {
		t.Errorf("expected name 'heading 1', got %q", nv)
	}

	if styles.GetByID("NoSuchStyle") != nil {
		t.Error("expected nil for unknown style")
	}
}

func TestCT_Styles_GetByName(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := s.SetNameVal("Normal"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}

	found := styles.GetByName("Normal")
	if found == nil {
		t.Fatal("expected to find style by name")
	}
	if found.StyleId() != "Normal" {
		t.Errorf("expected styleId 'Normal', got %q", found.StyleId())
	}
}

func TestCT_Styles_DefaultFor(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	xmlType, err := enum.WdStyleTypeParagraph.ToXml()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if err := s.SetType(xmlType); err != nil {
		t.Fatalf("SetType: %v", err)
	}
	if err := s.SetDefault(true); err != nil {
		t.Fatalf("SetDefault: %v", err)
	}

	def, err := styles.DefaultFor(enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if def == nil {
		t.Fatal("expected default style")
	}
	if def.StyleId() != "Normal" {
		t.Errorf("expected Normal, got %q", def.StyleId())
	}

	// No default for character
	defChar, err := styles.DefaultFor(enum.WdStyleTypeCharacter)
	if err != nil {
		t.Fatal(err)
	}
	if defChar != nil {
		t.Error("expected nil for character type")
	}
}

func TestCT_Styles_AddStyleOfType(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s, err := styles.AddStyleOfType("My Custom Style", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}

	if s.StyleId() != "MyCustomStyle" {
		t.Errorf("expected styleId MyCustomStyle, got %q", s.StyleId())
	}
	if nv, err := s.NameVal(); err != nil {
		t.Fatalf("NameVal: %v", err)
	} else if nv != "My Custom Style" {
		t.Errorf("expected name 'My Custom Style', got %q", nv)
	}
	if s.Type() != "paragraph" {
		t.Errorf("expected type paragraph, got %q", s.Type())
	}
	if !s.CustomStyle() {
		t.Error("expected customStyle=true for non-builtin")
	}

	// Builtin
	b, err := styles.AddStyleOfType("Heading 1", enum.WdStyleTypeParagraph, true)
	if err != nil {
		t.Fatal(err)
	}
	if b.CustomStyle() {
		t.Error("expected customStyle=false for builtin")
	}
	if b.StyleId() != "Heading1" {
		t.Errorf("expected Heading1, got %q", b.StyleId())
	}
}

func TestCT_Style_NameVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if nv, err := s.NameVal(); err != nil {
		t.Fatalf("NameVal: %v", err)
	} else if nv != "" {
		t.Errorf("expected empty, got %q", nv)
	}
	if err := s.SetNameVal("Normal"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}
	if nv, err := s.NameVal(); err != nil {
		t.Fatalf("NameVal: %v", err)
	} else if nv != "Normal" {
		t.Errorf("expected Normal, got %q", nv)
	}
	if err := s.SetNameVal(""); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}
	if nv, err := s.NameVal(); err != nil {
		t.Fatalf("NameVal: %v", err)
	} else if nv != "" {
		t.Errorf("expected empty after clear, got %q", nv)
	}
}

func TestCT_Style_BasedOnVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if bv, err := s.BasedOnVal(); err != nil {
		t.Fatalf("BasedOnVal: %v", err)
	} else if bv != "" {
		t.Errorf("expected empty, got %q", bv)
	}
	if err := s.SetBasedOnVal("Normal"); err != nil {
		t.Fatalf("SetBasedOnVal: %v", err)
	}
	if bv, err := s.BasedOnVal(); err != nil {
		t.Fatalf("BasedOnVal: %v", err)
	} else if bv != "Normal" {
		t.Errorf("expected Normal, got %q", bv)
	}
}

func TestCT_Style_NextVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetNextVal("Normal"); err != nil {
		t.Fatalf("SetNextVal: %v", err)
	}
	if nxv, err := s.NextVal(); err != nil {
		t.Fatalf("NextVal: %v", err)
	} else if nxv != "Normal" {
		t.Errorf("expected Normal, got %q", nxv)
	}
	if err := s.SetNextVal(""); err != nil {
		t.Fatalf("SetNextVal: %v", err)
	}
	if nxv, err := s.NextVal(); err != nil {
		t.Fatalf("NextVal: %v", err)
	} else if nxv != "" {
		t.Errorf("expected empty, got %q", nxv)
	}
}

func TestCT_Style_LockedVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.LockedVal() {
		t.Error("expected false by default")
	}
	if err := s.SetLockedVal(true); err != nil {
		t.Fatalf("SetLockedVal: %v", err)
	}
	if !s.LockedVal() {
		t.Error("expected true")
	}
	if err := s.SetLockedVal(false); err != nil {
		t.Fatalf("SetLockedVal: %v", err)
	}
	if s.LockedVal() {
		t.Error("expected false after clear")
	}
}

func TestCT_Style_SemiHiddenVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.SemiHiddenVal() {
		t.Error("expected false")
	}
	if err := s.SetSemiHiddenVal(true); err != nil {
		t.Fatalf("SetSemiHiddenVal: %v", err)
	}
	if !s.SemiHiddenVal() {
		t.Error("expected true")
	}
}

func TestCT_Style_UnhideWhenUsedVal_RoundTrip(t *testing.T) {
	t.Parallel()

	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s, err := styles.AddStyleOfType("Custom", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}

	if s.UnhideWhenUsedVal() {
		t.Error("expected false by default")
	}

	if err := s.SetUnhideWhenUsedVal(true); err != nil {
		t.Fatalf("SetUnhideWhenUsedVal: %v", err)
	}
	if !s.UnhideWhenUsedVal() {
		t.Error("expected true")
	}

	if err := s.SetUnhideWhenUsedVal(false); err != nil {
		t.Fatalf("SetUnhideWhenUsedVal: %v", err)
	}
	if s.UnhideWhenUsedVal() {
		t.Error("expected false after removing")
	}
}

func TestCT_Style_QFormatVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	s.SetQFormatVal(true)
	if !s.QFormatVal() {
		t.Error("expected true")
	}
	s.SetQFormatVal(false)
	if s.QFormatVal() {
		t.Error("expected false")
	}
}

func TestCT_Style_UiPriorityVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if uv, err := s.UiPriorityVal(); err != nil {
		t.Fatalf("UiPriorityVal: %v", err)
	} else if uv != nil {
		t.Error("expected nil")
	}
	v := 99
	if err := s.SetUiPriorityVal(&v); err != nil {
		t.Fatalf("SetUiPriorityVal: %v", err)
	}
	got, err := s.UiPriorityVal()
	if err != nil {
		t.Fatalf("UiPriorityVal: %v", err)
	}
	if got == nil || *got != 99 {
		t.Errorf("expected 99, got %v", got)
	}
	if err := s.SetUiPriorityVal(nil); err != nil {
		t.Fatalf("SetUiPriorityVal: %v", err)
	}
	if uv, err := s.UiPriorityVal(); err != nil {
		t.Fatalf("UiPriorityVal: %v", err)
	} else if uv != nil {
		t.Error("expected nil after clear")
	}
}

func TestCT_Style_BaseStyle(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	normal := styles.AddStyle()
	if err := normal.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := normal.SetNameVal("Normal"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}

	heading := styles.AddStyle()
	if err := heading.SetStyleId("Heading1"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := heading.SetBasedOnVal("Normal"); err != nil {
		t.Fatalf("SetBasedOnVal: %v", err)
	}

	base := heading.BaseStyle()
	if base == nil {
		t.Fatal("expected base style")
	}
	if base.StyleId() != "Normal" {
		t.Errorf("expected Normal, got %q", base.StyleId())
	}
}

func TestCT_Style_NextStyle(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	normal := styles.AddStyle()
	if err := normal.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}

	heading := styles.AddStyle()
	if err := heading.SetStyleId("Heading1"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := heading.SetNextVal("Normal"); err != nil {
		t.Fatalf("SetNextVal: %v", err)
	}

	next := heading.NextStyle()
	if next == nil {
		t.Fatal("expected next style")
	}
	if next.StyleId() != "Normal" {
		t.Errorf("expected Normal, got %q", next.StyleId())
	}
}

func TestCT_Style_Delete(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("ToDelete"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if styles.GetByID("ToDelete") == nil {
		t.Fatal("style should exist before delete")
	}
	s.Delete()
	if styles.GetByID("ToDelete") != nil {
		t.Error("style should be removed after delete")
	}
}

func TestCT_Style_IsBuiltin(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s, err := styles.AddStyleOfType("Normal", enum.WdStyleTypeParagraph, true)
	if err != nil {
		t.Fatal(err)
	}
	if !s.IsBuiltin() {
		t.Error("expected builtin")
	}
	custom, err := styles.AddStyleOfType("My Style", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	if custom.IsBuiltin() {
		t.Error("expected not builtin")
	}
}

func TestCT_LatentStyles_GetByName(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()
	exc := ls.AddLsdException()
	if err := exc.SetName("Heading 1"); err != nil {
		t.Fatalf("SetName: %v", err)
	}
	prio := 9
	if err := exc.SetUiPriority(&prio); err != nil {
		t.Fatalf("SetUiPriority: %v", err)
	}

	found := ls.GetByName("Heading 1")
	if found == nil {
		t.Fatal("expected to find exception")
	}
	if up, err := found.UiPriority(); err != nil {
		t.Fatalf("UiPriority: %v", err)
	} else if up == nil || *up != 9 {
		t.Errorf("expected priority 9, got %v", up)
	}
	if ls.GetByName("NoSuch") != nil {
		t.Error("expected nil for unknown")
	}
}

func TestCT_LatentStyles_BoolProp(t *testing.T) {
	t.Parallel()

	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()

	// Initially false (not set)
	if ls.BoolProp("w:defSemiHidden") {
		t.Error("expected false for unset property")
	}

	if err := ls.SetBoolProp("w:defSemiHidden", true); err != nil {
		t.Fatalf("SetBoolProp: %v", err)
	}
	if !ls.BoolProp("w:defSemiHidden") {
		t.Error("expected true after setting")
	}

	if err := ls.SetBoolProp("w:defSemiHidden", false); err != nil {
		t.Fatalf("SetBoolProp: %v", err)
	}
	if ls.BoolProp("w:defSemiHidden") {
		t.Error("expected false after clearing")
	}
}

func TestCT_LsdException_Delete(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()
	exc := ls.AddLsdException()
	if err := exc.SetName("ToRemove"); err != nil {
		t.Fatalf("SetName: %v", err)
	}
	if ls.GetByName("ToRemove") == nil {
		t.Fatal("should exist")
	}
	exc.Delete()
	if ls.GetByName("ToRemove") != nil {
		t.Error("should be removed")
	}
}

func TestCT_LsdException_OnOffProp(t *testing.T) {
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()
	exc := ls.AddLsdException()
	if err := exc.SetName("Test"); err != nil {
		t.Fatalf("SetName: %v", err)
	}

	// nil by default (attr not set)
	if exc.OnOffProp("w:locked") != nil {
		// Note: SetLocked in generated code defaults to false removal
		// OnOffProp reads raw attr
	}
	tr := true
	if err := exc.SetOnOffProp("w:locked", &tr); err != nil {
		t.Fatalf("SetOnOffProp: %v", err)
	}
	got := exc.OnOffProp("w:locked")
	if got == nil || !*got {
		t.Error("expected locked=true")
	}
	if err := exc.SetOnOffProp("w:locked", nil); err != nil {
		t.Fatalf("SetOnOffProp: %v", err)
	}
	if exc.OnOffProp("w:locked") != nil {
		t.Error("expected nil after removal")
	}
}
