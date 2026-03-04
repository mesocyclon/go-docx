package docx

import (
	"fmt"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// latent_styles_test.go — LatentStyles + LatentStyle (Batch 2)
// Mirrors Python: tests/styles/test_latent.py
// -----------------------------------------------------------------------

// helper: build a LatentStyles from inner XML
func makeLatentStyles(t *testing.T, innerXml string) *LatentStyles {
	t.Helper()
	xml := fmt.Sprintf(
		`<w:latentStyles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"%s>%s</w:latentStyles>`,
		"", innerXml,
	)
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	ls := &oxml.CT_LatentStyles{Element: oxml.WrapElement(el)}
	return &LatentStyles{element: ls}
}

func makeLatentStylesWithAttrs(t *testing.T, attrs string, innerXml string) *LatentStyles {
	t.Helper()
	xml := fmt.Sprintf(
		`<w:latentStyles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" %s>%s</w:latentStyles>`,
		attrs, innerXml,
	)
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	ls := &oxml.CT_LatentStyles{Element: oxml.WrapElement(el)}
	return &LatentStyles{element: ls}
}

// -----------------------------------------------------------------------
// LatentStyles collection tests
// -----------------------------------------------------------------------

// Mirrors Python: it_knows_how_many_latent_styles_it_contains
func TestLatentStyles_Len(t *testing.T) {
	tests := []struct {
		name  string
		inner string
		want  int
	}{
		{"empty", "", 0},
		{"one", `<w:lsdException w:name="Foo"/>`, 1},
		{"two", `<w:lsdException w:name="Foo"/><w:lsdException w:name="Bar"/>`, 2},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStyles(t, tt.inner)
			if got := ls.Len(); got != tt.want {
				t.Errorf("Len() = %d, want %d", got, tt.want)
			}
		})
	}
}

// Mirrors Python: it_can_iterate_over_its_latent_styles
func TestLatentStyles_Iter(t *testing.T) {
	tests := []struct {
		name  string
		inner string
		want  int
	}{
		{"empty", "", 0},
		{"one", `<w:lsdException w:name="Foo"/>`, 1},
		{"two", `<w:lsdException w:name="Foo"/><w:lsdException w:name="Bar"/>`, 2},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStyles(t, tt.inner)
			items := ls.Iter()
			if len(items) != tt.want {
				t.Errorf("Iter() len = %d, want %d", len(items), tt.want)
			}
		})
	}
}

// Mirrors Python: it_can_get_a_latent_style_by_name
func TestLatentStyles_Get(t *testing.T) {
	tests := []struct {
		name      string
		inner     string
		styleName string
	}{
		{"first", `<w:lsdException w:name="Ab"/><w:lsdException w:name="Cd"/>`, "Ab"},
		{"second", `<w:lsdException w:name="Ab"/><w:lsdException w:name="Cd"/>`, "Cd"},
		// Python: case-insensitive lookup "heading 1" → "Heading 1"
		{"case_insensitive", `<w:lsdException w:name="heading 1"/>`, "Heading 1"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStyles(t, tt.inner)
			style, err := ls.Get(tt.styleName)
			if err != nil {
				t.Fatalf("Get(%q): %v", tt.styleName, err)
			}
			if style == nil {
				t.Fatal("Get returned nil")
			}
		})
	}
}

// Mirrors Python: it_raises_on_latent_style_not_found
func TestLatentStyles_Get_NotFound(t *testing.T) {
	ls := makeLatentStyles(t, "")
	_, err := ls.Get("Foobar")
	if err == nil {
		t.Error("expected error for non-existent latent style")
	}
}

// Mirrors Python: it_can_add_a_latent_style
func TestLatentStyles_AddLatentStyle(t *testing.T) {
	ls := makeLatentStyles(t, "")
	style := ls.AddLatentStyle("Heading 1")

	if style == nil {
		t.Fatal("AddLatentStyle returned nil")
	}
	if ls.Len() != 1 {
		t.Errorf("Len() = %d after add, want 1", ls.Len())
	}
	// Verify the style can be retrieved
	got, err := ls.Get("Heading 1")
	if err != nil {
		t.Fatalf("Get after add: %v", err)
	}
	if got == nil {
		t.Error("Get returned nil after add")
	}
}

// Mirrors Python: it_knows_its_default_priority
func TestLatentStyles_DefaultPriority(t *testing.T) {
	tests := []struct {
		name  string
		attrs string
		want  *int
	}{
		{"absent", "", nil},
		{"set_42", `w:defUIPriority="42"`, intPtr(42)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStylesWithAttrs(t, tt.attrs, "")
			got, err := ls.DefaultPriority()
			if err != nil {
				t.Fatalf("DefaultPriority: %v", err)
			}
			if tt.want == nil {
				if got != nil {
					t.Errorf("DefaultPriority() = %d, want nil", *got)
				}
			} else {
				if got == nil || *got != *tt.want {
					t.Errorf("DefaultPriority() = %v, want %d", got, *tt.want)
				}
			}
		})
	}
}

// Mirrors Python: it_can_change_its_default_priority
func TestLatentStyles_SetDefaultPriority(t *testing.T) {
	tests := []struct {
		name  string
		attrs string
		value *int
	}{
		{"set_42", "", intPtr(42)},
		{"change_24_to_42", `w:defUIPriority="24"`, intPtr(42)},
		{"remove", `w:defUIPriority="24"`, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStylesWithAttrs(t, tt.attrs, "")
			if err := ls.SetDefaultPriority(tt.value); err != nil {
				t.Fatalf("SetDefaultPriority: %v", err)
			}
			got, _ := ls.DefaultPriority()
			if tt.value == nil {
				if got != nil {
					t.Errorf("after set nil: got %d", *got)
				}
			} else {
				if got == nil || *got != *tt.value {
					t.Errorf("after set %d: got %v", *tt.value, got)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_its_load_count
func TestLatentStyles_LoadCount(t *testing.T) {
	tests := []struct {
		name  string
		attrs string
		want  *int
	}{
		{"absent", "", nil},
		{"set_42", `w:count="42"`, intPtr(42)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStylesWithAttrs(t, tt.attrs, "")
			got, err := ls.LoadCount()
			if err != nil {
				t.Fatalf("LoadCount: %v", err)
			}
			if tt.want == nil {
				if got != nil {
					t.Errorf("LoadCount() = %d, want nil", *got)
				}
			} else {
				if got == nil || *got != *tt.want {
					t.Errorf("LoadCount() = %v, want %d", got, *tt.want)
				}
			}
		})
	}
}

// Mirrors Python: it_can_change_its_load_count
func TestLatentStyles_SetLoadCount(t *testing.T) {
	tests := []struct {
		name  string
		attrs string
		value *int
	}{
		{"set_42", "", intPtr(42)},
		{"change", `w:count="24"`, intPtr(42)},
		{"remove", `w:count="24"`, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStylesWithAttrs(t, tt.attrs, "")
			if err := ls.SetLoadCount(tt.value); err != nil {
				t.Fatalf("SetLoadCount: %v", err)
			}
			got, _ := ls.LoadCount()
			if tt.value == nil {
				if got != nil {
					t.Errorf("after set nil: got %d", *got)
				}
			} else {
				if got == nil || *got != *tt.value {
					t.Errorf("after set %d: got %v", *tt.value, got)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_its_boolean_properties (default_to_hidden, etc.)
func TestLatentStyles_BoolProps_Getter(t *testing.T) {
	tests := []struct {
		name   string
		attrs  string
		prop   string
		getter func(*LatentStyles) bool
		want   bool
	}{
		{"hidden_absent", "", "default_to_hidden", (*LatentStyles).DefaultToHidden, false},
		{"locked_absent", "", "default_to_locked", (*LatentStyles).DefaultToLocked, false},
		{"quickstyle_absent", "", "default_to_quick_style", (*LatentStyles).DefaultToQuickStyle, false},
		{"unhide_absent", "", "default_to_unhide_when_used", (*LatentStyles).DefaultToUnhideWhenUsed, false},
		{"hidden_true", `w:defSemiHidden="1"`, "default_to_hidden", (*LatentStyles).DefaultToHidden, true},
		{"locked_false", `w:defLockedState="0"`, "default_to_locked", (*LatentStyles).DefaultToLocked, false},
		{"quickstyle_true", `w:defQFormat="1"`, "default_to_quick_style", (*LatentStyles).DefaultToQuickStyle, true},
		{"unhide_false", `w:defUnhideWhenUsed="0"`, "default_to_unhide_when_used", (*LatentStyles).DefaultToUnhideWhenUsed, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStylesWithAttrs(t, tt.attrs, "")
			got := tt.getter(ls)
			if got != tt.want {
				t.Errorf("%s = %v, want %v", tt.prop, got, tt.want)
			}
		})
	}
}

// Mirrors Python: it_can_change_its_boolean_properties
func TestLatentStyles_BoolProps_Setter(t *testing.T) {
	tests := []struct {
		name   string
		attrs  string
		setter func(*LatentStyles, bool) error
		getter func(*LatentStyles) bool
		value  bool
	}{
		{"set_hidden_true", "", (*LatentStyles).SetDefaultToHidden, (*LatentStyles).DefaultToHidden, true},
		{"set_locked_false", "", (*LatentStyles).SetDefaultToLocked, (*LatentStyles).DefaultToLocked, false},
		{"set_quickstyle_true", "", (*LatentStyles).SetDefaultToQuickStyle, (*LatentStyles).DefaultToQuickStyle, true},
		{"set_unhide_false", "", (*LatentStyles).SetDefaultToUnhideWhenUsed, (*LatentStyles).DefaultToUnhideWhenUsed, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ls := makeLatentStylesWithAttrs(t, tt.attrs, "")
			if err := tt.setter(ls, tt.value); err != nil {
				t.Fatalf("setter: %v", err)
			}
			got := tt.getter(ls)
			if got != tt.value {
				t.Errorf("after set %v: got %v", tt.value, got)
			}
		})
	}
}

// -----------------------------------------------------------------------
// LatentStyle (individual lsdException) tests
// -----------------------------------------------------------------------

// helper: build a LatentStyle from an lsdException XML
func makeLatentStyleItem(t *testing.T, attrs string) (*LatentStyles, *LatentStyle) {
	t.Helper()
	inner := fmt.Sprintf(`<w:lsdException %s/>`, attrs)
	ls := makeLatentStyles(t, inner)
	items := ls.Iter()
	if len(items) != 1 {
		t.Fatalf("expected 1 item, got %d", len(items))
	}
	return ls, items[0]
}

// Mirrors Python: it_knows_its_name
func TestLatentStyle_Name(t *testing.T) {
	_, style := makeLatentStyleItem(t, `w:name="heading 1"`)
	got := style.Name()
	if got != "Heading 1" {
		t.Errorf("Name() = %q, want %q", got, "Heading 1")
	}
}

// Mirrors Python: it_knows_its_priority
func TestLatentStyle_Priority(t *testing.T) {
	tests := []struct {
		name  string
		attrs string
		want  *int
	}{
		{"absent", `w:name="Foo"`, nil},
		{"set_42", `w:name="Foo" w:uiPriority="42"`, intPtr(42)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, style := makeLatentStyleItem(t, tt.attrs)
			got, err := style.Priority()
			if err != nil {
				t.Fatalf("Priority: %v", err)
			}
			if tt.want == nil {
				if got != nil {
					t.Errorf("Priority() = %d, want nil", *got)
				}
			} else {
				if got == nil || *got != *tt.want {
					t.Errorf("Priority() = %v, want %d", got, *tt.want)
				}
			}
		})
	}
}

// Mirrors Python: it_can_change_its_priority
func TestLatentStyle_SetPriority(t *testing.T) {
	tests := []struct {
		name  string
		attrs string
		value *int
	}{
		{"set_42", `w:name="Foo"`, intPtr(42)},
		{"change", `w:name="Foo" w:uiPriority="42"`, intPtr(24)},
		{"remove", `w:name="Foo" w:uiPriority="24"`, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, style := makeLatentStyleItem(t, tt.attrs)
			if err := style.SetPriority(tt.value); err != nil {
				t.Fatalf("SetPriority: %v", err)
			}
			got, _ := style.Priority()
			if tt.value == nil {
				if got != nil {
					t.Errorf("after set nil: got %d", *got)
				}
			} else {
				if got == nil || *got != *tt.value {
					t.Errorf("after set %d: got %v", *tt.value, got)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_its_on_off_properties (12 cases)
func TestLatentStyle_OnOffProps_Getter(t *testing.T) {
	type getter func(*LatentStyle) *bool
	tests := []struct {
		name   string
		attrs  string
		fn     getter
		expect *bool
	}{
		// absent → nil
		{"hidden_nil", `w:name="Foo"`, (*LatentStyle).Hidden, nil},
		{"locked_nil", `w:name="Foo"`, (*LatentStyle).Locked, nil},
		{"quickstyle_nil", `w:name="Foo"`, (*LatentStyle).QuickStyle, nil},
		{"unhide_nil", `w:name="Foo"`, (*LatentStyle).UnhideWhenUsed, nil},
		// true
		{"hidden_true", `w:name="Foo" w:semiHidden="1"`, (*LatentStyle).Hidden, boolPtr(true)},
		{"locked_true", `w:name="Foo" w:locked="1"`, (*LatentStyle).Locked, boolPtr(true)},
		{"quickstyle_true", `w:name="Foo" w:qFormat="1"`, (*LatentStyle).QuickStyle, boolPtr(true)},
		{"unhide_true", `w:name="Foo" w:unhideWhenUsed="1"`, (*LatentStyle).UnhideWhenUsed, boolPtr(true)},
		// false
		{"hidden_false", `w:name="Foo" w:semiHidden="0"`, (*LatentStyle).Hidden, boolPtr(false)},
		{"locked_false", `w:name="Foo" w:locked="0"`, (*LatentStyle).Locked, boolPtr(false)},
		{"quickstyle_false", `w:name="Foo" w:qFormat="0"`, (*LatentStyle).QuickStyle, boolPtr(false)},
		{"unhide_false", `w:name="Foo" w:unhideWhenUsed="0"`, (*LatentStyle).UnhideWhenUsed, boolPtr(false)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, style := makeLatentStyleItem(t, tt.attrs)
			got := tt.fn(style)
			compareBoolPtr(t, tt.name, got, tt.expect)
		})
	}
}

// Mirrors Python: it_can_change_its_on_off_properties (7 cases)
func TestLatentStyle_OnOffProps_Setter(t *testing.T) {
	type setter func(*LatentStyle, *bool) error
	type getter func(*LatentStyle) *bool
	tests := []struct {
		name   string
		attrs  string
		set    setter
		get    getter
		value  *bool
		expect *bool
	}{
		{"set_hidden_true", `w:name="Foo"`,
			(*LatentStyle).SetHidden, (*LatentStyle).Hidden, boolPtr(true), boolPtr(true)},
		{"hidden_true_to_false", `w:name="Foo" w:semiHidden="1"`,
			(*LatentStyle).SetHidden, (*LatentStyle).Hidden, boolPtr(false), boolPtr(false)},
		{"hidden_false_to_nil", `w:name="Foo" w:semiHidden="0"`,
			(*LatentStyle).SetHidden, (*LatentStyle).Hidden, nil, nil},
		{"set_locked_true", `w:name="Foo"`,
			(*LatentStyle).SetLocked, (*LatentStyle).Locked, boolPtr(true), boolPtr(true)},
		{"set_quickstyle_false", `w:name="Foo"`,
			(*LatentStyle).SetQuickStyle, (*LatentStyle).QuickStyle, boolPtr(false), boolPtr(false)},
		{"set_unhide_true", `w:name="Foo"`,
			(*LatentStyle).SetUnhideWhenUsed, (*LatentStyle).UnhideWhenUsed, boolPtr(true), boolPtr(true)},
		{"locked_true_to_nil", `w:name="Foo" w:locked="1"`,
			(*LatentStyle).SetLocked, (*LatentStyle).Locked, nil, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, style := makeLatentStyleItem(t, tt.attrs)
			if err := tt.set(style, tt.value); err != nil {
				t.Fatalf("setter: %v", err)
			}
			got := tt.get(style)
			compareBoolPtr(t, tt.name, got, tt.expect)
		})
	}
}

// Mirrors Python: it_can_delete_itself
func TestLatentStyle_Delete(t *testing.T) {
	ls := makeLatentStyles(t, `<w:lsdException w:name="Foo"/>`)
	if ls.Len() != 1 {
		t.Fatalf("precondition: Len() = %d, want 1", ls.Len())
	}
	items := ls.Iter()
	items[0].Delete()

	if ls.Len() != 0 {
		t.Errorf("after delete: Len() = %d, want 0", ls.Len())
	}
}
