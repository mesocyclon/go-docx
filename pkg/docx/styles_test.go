package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// styles_test.go — Styles / BaseStyle (Batch 1)
// Mirrors Python: tests/styles/test_styles.py + test_style.py
// -----------------------------------------------------------------------

func makeStylesFromDoc(t *testing.T) *Styles {
	t.Helper()
	doc := mustNewDoc(t)
	ss, err := doc.Styles()
	if err != nil {
		t.Fatal(err)
	}
	return ss
}

// Mirrors Python: Styles.contains
func TestStyles_Contains(t *testing.T) {
	ss := makeStylesFromDoc(t)
	// "Normal" should always exist in a default doc
	if !ss.Contains("Normal") {
		t.Error("Contains(Normal) = false, want true")
	}
	if ss.Contains("Nonexistent Style XYZ") {
		t.Error("Contains(Nonexistent) = true, want false")
	}
}

// Mirrors Python: Styles.get
func TestStyles_Get(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.Get("Normal")
	if err != nil {
		t.Fatalf("Get(Normal): %v", err)
	}
	if style == nil {
		t.Fatal("Get(Normal) returned nil")
	}
	// Name should round-trip
	gotName, err := style.Name()
	if err != nil {
		t.Fatal(err)
	}
	if gotName != "Normal" {
		t.Errorf("Name() = %q, want %q", gotName, "Normal")
	}
}

// Mirrors Python: Styles.iter / Styles.len
func TestStyles_Iter_Len(t *testing.T) {
	ss := makeStylesFromDoc(t)
	length := ss.Len()
	if length == 0 {
		t.Error("Len() = 0, want > 0")
	}
	iter := ss.Iter()
	if len(iter) != length {
		t.Errorf("len(Iter()) = %d, want %d", len(iter), length)
	}
}

// Mirrors Python: Styles.add_style
func TestStyles_AddStyle(t *testing.T) {
	ss := makeStylesFromDoc(t)
	newStyle, err := ss.AddStyle("MyCustomPara", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	if newStyle == nil {
		t.Fatal("AddStyle returned nil")
	}
	if !ss.Contains("MyCustomPara") {
		t.Error("Contains(MyCustomPara) = false after AddStyle")
	}
	// Duplicate should error
	_, err = ss.AddStyle("MyCustomPara", enum.WdStyleTypeParagraph, false)
	if err == nil {
		t.Error("expected error for duplicate AddStyle")
	}
}

// Mirrors Python: Styles.default
func TestStyles_Default(t *testing.T) {
	ss := makeStylesFromDoc(t)
	def, err := ss.Default(enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if def == nil {
		t.Error("Default(Paragraph) returned nil")
	}
}

// Mirrors Python: Styles.get_by_id
func TestStyles_GetByID(t *testing.T) {
	ss := makeStylesFromDoc(t)
	// Nil → default
	def, err := ss.GetByID(nil, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if def == nil {
		t.Error("GetByID(nil) returned nil")
	}
	// Known ID
	id := "Normal"
	style, err := ss.GetByID(&id, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if style == nil {
		t.Error("GetByID(Normal) returned nil")
	}
}

// Mirrors Python: BaseStyle.name / set_name
func TestBaseStyle_Name(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.AddStyle("TestNameProp", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	gotN1, err := style.Name()
	if err != nil {
		t.Fatal(err)
	}
	if gotN1 != "TestNameProp" {
		t.Errorf("Name() = %q, want %q", gotN1, "TestNameProp")
	}
	if err := style.SetName("Renamed"); err != nil {
		t.Fatal(err)
	}
	gotN2, err := style.Name()
	if err != nil {
		t.Fatal(err)
	}
	if gotN2 != "Renamed" {
		t.Errorf("Name() after rename = %q, want %q", gotN2, "Renamed")
	}
}

// Mirrors Python: BaseStyle.style_id / set_style_id
func TestBaseStyle_StyleID(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.AddStyle("TestIDProp", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	origID := style.StyleID()
	if origID == "" {
		t.Error("StyleID() is empty")
	}
	if err := style.SetStyleID("CustomID"); err != nil {
		t.Fatal(err)
	}
	if style.StyleID() != "CustomID" {
		t.Errorf("StyleID() after set = %q, want %q", style.StyleID(), "CustomID")
	}
}

// Mirrors Python: BaseStyle.type
func TestBaseStyle_Type(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.AddStyle("TestTypeProp", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	got, err := style.Type()
	if err != nil {
		t.Fatal(err)
	}
	if got != enum.WdStyleTypeParagraph {
		t.Errorf("Type() = %v, want Paragraph", got)
	}
}

// Mirrors Python: BaseStyle.hidden / locked / quick_style
func TestBaseStyle_BoolProps(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.AddStyle("TestBoolProps", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}

	// Hidden
	if style.Hidden() {
		t.Error("Hidden() = true initially")
	}
	if err := style.SetHidden(true); err != nil {
		t.Fatal(err)
	}
	if !style.Hidden() {
		t.Error("Hidden() = false after SetHidden(true)")
	}

	// Locked
	if style.Locked() {
		t.Error("Locked() = true initially")
	}
	if err := style.SetLocked(true); err != nil {
		t.Fatal(err)
	}
	if !style.Locked() {
		t.Error("Locked() = false after SetLocked(true)")
	}

	// QuickStyle
	style.SetQuickStyle(true)
	if !style.QuickStyle() {
		t.Error("QuickStyle() = false after SetQuickStyle(true)")
	}
}

// Mirrors Python: BaseStyle.priority
func TestBaseStyle_Priority(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.AddStyle("TestPriorityProp", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}

	// Set
	v := 42
	if err := style.SetPriority(&v); err != nil {
		t.Fatal(err)
	}
	got, err := style.Priority()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 42 {
		t.Errorf("Priority() = %v, want 42", got)
	}

	// Remove
	if err := style.SetPriority(nil); err != nil {
		t.Fatal(err)
	}
	got2, _ := style.Priority()
	if got2 != nil {
		t.Errorf("Priority() after remove = %v, want nil", got2)
	}
}

// Mirrors Python: BaseStyle.delete
func TestBaseStyle_Delete(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.AddStyle("ToDelete", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	if !ss.Contains("ToDelete") {
		t.Fatal("style not found after add")
	}
	style.Delete()
	if ss.Contains("ToDelete") {
		t.Error("style still found after Delete()")
	}
}

// Mirrors Python: BaseStyle.paragraph_format
func TestBaseStyle_ParagraphFormat2(t *testing.T) {
	ss := makeStylesFromDoc(t)
	style, err := ss.Get("Normal")
	if err != nil {
		t.Fatal(err)
	}
	pf := style.ParagraphFormat()
	if pf == nil {
		t.Error("ParagraphFormat() returned nil")
	}
}

// Mirrors Python: Styles.get_style_id (name→id resolution)
func TestStyles_GetStyleID(t *testing.T) {
	ss := makeStylesFromDoc(t)
	// nil → nil
	id, err := ss.GetStyleID(nil, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id != nil {
		t.Errorf("GetStyleID(nil) = %v, want nil", *id)
	}
}
