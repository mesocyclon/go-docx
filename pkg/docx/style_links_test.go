package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// style_links_test.go — BaseStyle links, Font, ParagraphFormat (Batch 2)
// Mirrors Python: tests/styles/test_style.py — CharacterStyle / ParagraphStyle
// -----------------------------------------------------------------------

// helper: wrap oxml CT_Styles into domain Styles
func makeDomainStyles(t *testing.T, innerXml string) *Styles {
	t.Helper()
	ct := makeStyles(t, innerXml)
	return newStyles(ct)
}

// Mirrors Python DescribeCharacterStyle: it_knows_which_style_it_is_based_on
func TestBaseStyle_BaseStyleObj(t *testing.T) {
	tests := []struct {
		name    string
		xml     string
		idx     int // style index to test
		wantNil bool
		wantID  string // expected base style's styleId
	}{
		{
			"has_base_style",
			`<w:style w:styleId="Foo"/><w:style><w:basedOn w:val="Foo"/></w:style>`,
			1, false, "Foo",
		},
		{
			"base_style_not_found",
			`<w:style w:styleId="Foo"/><w:style><w:basedOn w:val="Bar"/></w:style>`,
			1, true, "",
		},
		{
			"no_basedOn",
			`<w:style w:styleId="Foo"/>`,
			0, true, "",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			styles := makeDomainStyles(t, tt.xml)
			items := styles.Iter()
			if tt.idx >= len(items) {
				t.Fatalf("style index %d out of range (%d styles)", tt.idx, len(items))
			}
			style := items[tt.idx]

			base := style.BaseStyleObj()
			if tt.wantNil {
				if base != nil {
					t.Errorf("BaseStyleObj() = %q, want nil", base.StyleID())
				}
			} else {
				if base == nil {
					t.Fatal("BaseStyleObj() = nil, want non-nil")
				}
				if base.StyleID() != tt.wantID {
					t.Errorf("BaseStyleObj().StyleID() = %q, want %q", base.StyleID(), tt.wantID)
				}
			}
		})
	}
}

// Mirrors Python: it_can_change_its_base_style
func TestBaseStyle_SetBaseStyle(t *testing.T) {
	tests := []struct {
		name     string
		xml      string
		setToNil bool
		setToIdx int // index of style to set as base (ignored if setToNil)
		wantVal  string
		wantGone bool
	}{
		{
			"set_basedOn",
			`<w:style w:styleId="Foo"/><w:style/>`,
			false, 0, "Foo", false,
		},
		{
			"change_basedOn",
			`<w:style w:styleId="Foo"/><w:style w:styleId="Bar"/><w:style><w:basedOn w:val="Foo"/></w:style>`,
			false, 1, "Bar", false,
		},
		{
			"remove_basedOn",
			`<w:style w:styleId="Foo"/><w:style><w:basedOn w:val="Foo"/></w:style>`,
			true, 0, "", true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			styles := makeDomainStyles(t, tt.xml)
			items := styles.Iter()
			target := items[len(items)-1] // last style is always the one under test

			var err error
			if tt.setToNil {
				err = target.SetBaseStyle(nil)
			} else {
				err = target.SetBaseStyle(items[tt.setToIdx])
			}
			if err != nil {
				t.Fatalf("SetBaseStyle: %v", err)
			}

			// Verify via XML
			bo := target.CT_Style().RawElement().FindElement("basedOn")
			if tt.wantGone {
				if bo != nil {
					t.Error("expected basedOn to be removed")
				}
			} else {
				if bo == nil {
					t.Fatal("expected basedOn element")
				}
				gotVal := bo.SelectAttrValue("w:val", "")
				if gotVal != tt.wantVal {
					t.Errorf("basedOn w:val = %q, want %q", gotVal, tt.wantVal)
				}
			}
		})
	}
}

// Mirrors Python DescribeParagraphStyle: it_knows_its_next_paragraph_style
func TestBaseStyle_NextParagraphStyle(t *testing.T) {
	// Build styles: H1 → next:Body, H2 → next:Char(character type), Body(no next), Foo → next:Bar(not found)
	xml := `<w:style w:type="paragraph" w:styleId="H1"><w:next w:val="Body"/></w:style>` +
		`<w:style w:type="paragraph" w:styleId="H2"><w:next w:val="Char"/></w:style>` +
		`<w:style w:type="paragraph" w:styleId="Body"/>` +
		`<w:style w:type="paragraph" w:styleId="Foo"><w:next w:val="Bar"/></w:style>` +
		`<w:style w:type="character" w:styleId="Char"/>`

	styles := makeDomainStyles(t, xml)
	items := styles.Iter()

	tests := []struct {
		name       string
		styleIdx   int
		wantSelfID string // if returns self, match this ID
	}{
		// H1.next = Body → should return Body style
		{"H1_next_is_Body", 0, "Body"},
		// H2.next = Char (character style, different type) → returns self
		{"H2_next_wrong_type_returns_self", 1, "H2"},
		// Body has no next → returns self
		{"Body_no_next_returns_self", 2, "Body"},
		// Foo.next = Bar (not found) → falls through to self
		{"Foo_not_found_returns_self", 3, "Foo"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			style := items[tt.styleIdx]
			next := style.NextParagraphStyle()
			if next == nil {
				t.Fatal("NextParagraphStyle returned nil")
			}
			if next.StyleID() != tt.wantSelfID {
				t.Errorf("NextParagraphStyle().StyleID() = %q, want %q", next.StyleID(), tt.wantSelfID)
			}
		})
	}
}

// Mirrors Python: it_provides_access_to_its_font
func TestBaseStyle_Font(t *testing.T) {
	xml := `<w:style w:styleId="Foo"/>`
	styles := makeDomainStyles(t, xml)
	items := styles.Iter()
	style := items[0]

	font := style.Font()
	if font == nil {
		t.Fatal("Font() returned nil")
	}
	// Font should be usable — set a property and verify
	if err := font.SetBold(boolPtr(true)); err != nil {
		t.Fatalf("font.SetBold: %v", err)
	}
}

// Mirrors Python: it_provides_access_to_its_paragraph_format
func TestBaseStyle_ParagraphFormat(t *testing.T) {
	xml := `<w:style w:styleId="Foo"/>`
	styles := makeDomainStyles(t, xml)
	items := styles.Iter()
	style := items[0]

	pf := style.ParagraphFormat()
	if pf == nil {
		t.Fatal("ParagraphFormat() returned nil")
	}
}

// Mirrors Python BaseStyle: name, style_id, type, hidden, locked, quick_style, priority
func TestBaseStyle_Properties(t *testing.T) {
	xml := `<w:style w:type="paragraph" w:styleId="Heading1">` +
		`<w:name w:val="heading 1"/>` +
		`<w:qFormat/>` +
		`<w:uiPriority w:val="9"/>` +
		`</w:style>`
	styles := makeDomainStyles(t, xml)
	items := styles.Iter()
	style := items[0]

	// Name
	gotSN, err := style.Name()
	if err != nil {
		t.Fatal(err)
	}
	if gotSN != "Heading 1" {
		t.Errorf("Name() = %q, want %q", gotSN, "Heading 1")
	}
	// StyleID
	if got := style.StyleID(); got != "Heading1" {
		t.Errorf("StyleID() = %q, want %q", got, "Heading1")
	}
	// Type
	got, err := style.Type()
	if err != nil {
		t.Fatalf("Type: %v", err)
	}
	if got != enum.WdStyleTypeParagraph {
		t.Errorf("Type() = %v, want WdStyleTypeParagraph", got)
	}
	// QuickStyle (qFormat present without val → true)
	if !style.QuickStyle() {
		t.Error("QuickStyle() = false, want true")
	}
	// Priority
	pri, err := style.Priority()
	if err != nil {
		t.Fatalf("Priority: %v", err)
	}
	if pri == nil || *pri != 9 {
		t.Errorf("Priority() = %v, want 9", pri)
	}
	// Hidden (absent → false)
	if style.Hidden() {
		t.Error("Hidden() = true, want false")
	}
	// Locked (absent → false)
	if style.Locked() {
		t.Error("Locked() = true, want false")
	}
}

// Test UnhideWhenUsed getter
func TestBaseStyle_UnhideWhenUsed(t *testing.T) {
	tests := []struct {
		name string
		xml  string
		want bool
	}{
		{"absent", `<w:style w:styleId="Foo"/>`, false},
		{"present", `<w:style w:styleId="Foo"><w:unhideWhenUsed/></w:style>`, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			styles := makeDomainStyles(t, tt.xml)
			style := styles.Iter()[0]
			if got := style.UnhideWhenUsed(); got != tt.want {
				t.Errorf("UnhideWhenUsed() = %v, want %v", got, tt.want)
			}
		})
	}
}
