package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// run_test.go — Run (Batch 1)
// Mirrors Python: tests/text/test_run.py
// -----------------------------------------------------------------------

// Mirrors Python: it_knows_its_bool_prop_states (6 cases)
func TestRun_BoolPropGetters(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		prop     string
		expected *bool
	}{
		{"bold_nil", `<w:rPr/>`, "bold", nil},
		{"bold_default_true", `<w:rPr><w:b/></w:rPr>`, "bold", boolPtr(true)},
		{"bold_on", `<w:rPr><w:b w:val="on"/></w:rPr>`, "bold", boolPtr(true)},
		{"bold_off", `<w:rPr><w:b w:val="off"/></w:rPr>`, "bold", boolPtr(false)},
		{"bold_1", `<w:rPr><w:b w:val="1"/></w:rPr>`, "bold", boolPtr(true)},
		{"italic_0", `<w:rPr><w:i w:val="0"/></w:rPr>`, "italic", boolPtr(false)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			run := newRun(r, nil)
			var got *bool
			switch tt.prop {
			case "bold":
				got = run.Bold()
			case "italic":
				got = run.Italic()
			}
			compareBoolPtr(t, tt.prop, got, tt.expected)
		})
	}
}

// Mirrors Python: it_can_change_its_bool_prop_settings (12 cases)
func TestRun_BoolPropSetters(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		prop     string
		value    *bool
	}{
		// nothing → True, False, None
		{"empty_to_bold_true", ``, "bold", boolPtr(true)},
		{"empty_to_bold_false", ``, "bold", boolPtr(false)},
		{"empty_to_italic_nil", ``, "italic", nil},
		// default → True, False, None
		{"default_bold_to_true", `<w:rPr><w:b/></w:rPr>`, "bold", boolPtr(true)},
		{"default_bold_to_false", `<w:rPr><w:b/></w:rPr>`, "bold", boolPtr(false)},
		{"default_italic_to_nil", `<w:rPr><w:i/></w:rPr>`, "italic", nil},
		// True → True, False, None
		{"true_bold_to_true", `<w:rPr><w:b w:val="on"/></w:rPr>`, "bold", boolPtr(true)},
		{"true_bold_to_false", `<w:rPr><w:b w:val="1"/></w:rPr>`, "bold", boolPtr(false)},
		{"true_bold_to_nil", `<w:rPr><w:b w:val="1"/></w:rPr>`, "bold", nil},
		// False → True, False, None
		{"false_italic_to_true", `<w:rPr><w:i w:val="false"/></w:rPr>`, "italic", boolPtr(true)},
		{"false_italic_to_false", `<w:rPr><w:i w:val="0"/></w:rPr>`, "italic", boolPtr(false)},
		{"false_italic_to_nil", `<w:rPr><w:i w:val="off"/></w:rPr>`, "italic", nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			run := newRun(r, nil)
			var err error
			switch tt.prop {
			case "bold":
				err = run.SetBold(tt.value)
			case "italic":
				err = run.SetItalic(tt.value)
			}
			if err != nil {
				t.Fatalf("Set%s: %v", tt.prop, err)
			}
			// Re-read and verify
			var got *bool
			switch tt.prop {
			case "bold":
				got = run.Bold()
			case "italic":
				got = run.Italic()
			}
			compareBoolPtr(t, "after set "+tt.prop, got, tt.value)
		})
	}
}

// Mirrors Python: it_can_add_text (4 cases)
func TestRun_AddText(t *testing.T) {
	tests := []struct {
		name    string
		initial string
		addText string
	}{
		{"add_to_empty", ``, "foo"},
		{"add_to_existing", `<w:t>foo</w:t>`, "bar"},
		{"add_trailing_space", ``, "fo "},
		{"add_mid_space", ``, "f o"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.initial)
			run := newRun(r, nil)
			run.AddText(tt.addText)
			text := run.Text()
			expected := ""
			if tt.initial != "" {
				// Extract text from initial
				initR := makeR(t, tt.initial)
				expected = (&Run{r: initR}).Text()
			}
			expected += tt.addText
			if text != expected {
				t.Errorf("Text() after AddText = %q, want %q", text, expected)
			}
		})
	}
}

// Mirrors Python: it_can_add_a_break (6 break types)
func TestRun_AddBreak(t *testing.T) {
	tests := []struct {
		name      string
		breakType enum.WdBreakType
	}{
		{"line", enum.WdBreakTypeLine},
		{"page", enum.WdBreakTypePage},
		{"column", enum.WdBreakTypeColumn},
		{"line_clear_left", enum.WdBreakTypeLineClearLeft},
		{"line_clear_right", enum.WdBreakTypeLineClearRight},
		{"line_clear_all", enum.WdBreakTypeLineClearAll},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, "")
			run := newRun(r, nil)
			if err := run.AddBreak(tt.breakType); err != nil {
				t.Fatalf("AddBreak: %v", err)
			}
			// Verify a w:br element was added
			brElem := r.RawElement().FindElement("br")
			if brElem == nil {
				t.Fatal("expected w:br element after AddBreak")
			}
			// Verify attributes based on break type.
			// Note: "textWrapping" is the OOXML default for w:type, so
			// SetType("textWrapping") omits the attribute. The getter
			// returns "textWrapping" when absent. We check raw attrs
			// accordingly: expect "" for textWrapping (omitted).
			wantType, wantClear, _ := breakTypeToAttrs(tt.breakType)
			wantRawType := wantType
			if wantRawType == "textWrapping" {
				wantRawType = "" // omitted because it is the default
			}
			gotType := brElem.SelectAttrValue("w:type", "")
			gotClear := brElem.SelectAttrValue("w:clear", "")
			if gotType != wantRawType {
				t.Errorf("br type attr = %q, want %q", gotType, wantRawType)
			}
			if gotClear != wantClear {
				t.Errorf("br clear attr = %q, want %q", gotClear, wantClear)
			}
		})
	}
}

// Mirrors Python: dict[break_type] → KeyError for unsupported types
func TestRun_AddBreak_UnsupportedType(t *testing.T) {
	r := makeR(t, "")
	run := newRun(r, nil)
	// Section breaks are not valid in Run.AddBreak
	err := run.AddBreak(enum.WdBreakTypeSectionNextPage)
	if err == nil {
		t.Error("expected error for unsupported break type")
	}
}

// Mirrors Python: it_can_add_a_tab
func TestRun_AddTab(t *testing.T) {
	r := makeR(t, `<w:t>foo</w:t>`)
	run := newRun(r, nil)
	run.AddTab()
	// Verify a w:tab element was added
	tabElem := r.RawElement().FindElement("tab")
	if tabElem == nil {
		t.Fatal("expected w:tab element after AddTab")
	}
}

// Mirrors Python: it_knows_the_text_it_contains (4 cases)
func TestRun_Text(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected string
	}{
		{"empty", ``, ""},
		{"simple", `<w:t>foobar</w:t>`, "foobar"},
		{"mixed_tab_cr", `<w:t>abc</w:t><w:tab/><w:t>def</w:t><w:cr/>`, "abc\tdef\n"},
		{"page_break_and_tab", `<w:br w:type="page"/><w:t>abc</w:t><w:t>def</w:t><w:tab/>`, "abcdef\t"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			run := newRun(r, nil)
			if got := run.Text(); got != tt.expected {
				t.Errorf("Text() = %q, want %q", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_can_replace_the_text_it_contains (4 cases)
func TestRun_SetText(t *testing.T) {
	tests := []struct {
		name     string
		newText  string
		expected string
	}{
		{"plain", "abc  def", "abc  def"},
		{"with_tab", "abc\tdef", "abc\tdef"},
		{"with_newline", "abc\ndef", "abc\ndef"},
		{"with_cr", "abc\rdef", "abc\ndef"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, `<w:t>should get deleted</w:t>`)
			run := newRun(r, nil)
			run.SetText(tt.newText)
			if got := run.Text(); got != tt.expected {
				t.Errorf("Text() after SetText(%q) = %q, want %q", tt.newText, got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_can_remove_its_content_but_keep_formatting (6 cases)
func TestRun_Clear(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		wantRPr  bool
	}{
		{"empty", ``, false},
		{"text_only", `<w:t>foo</w:t>`, false},
		{"br_only", `<w:br/>`, false},
		{"rPr_only", `<w:rPr/>`, true},
		{"rPr_and_text", `<w:rPr/><w:t>foo</w:t>`, true},
		{"rPr_complex_and_content", `<w:rPr><w:b/><w:i/></w:rPr><w:t>foo</w:t><w:cr/><w:t>bar</w:t>`, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			run := newRun(r, nil)
			run.Clear()
			if got := run.Text(); got != "" {
				t.Errorf("Text() after Clear = %q, want empty", got)
			}
			hasRPr := r.RPr() != nil
			if hasRPr != tt.wantRPr {
				t.Errorf("rPr present = %v, want %v", hasRPr, tt.wantRPr)
			}
		})
	}
}

// Mirrors Python: it_knows_whether_it_contains_a_page_break
func TestRun_ContainsPageBreak(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected bool
	}{
		{"no_content", ``, false},
		{"text_only", `<w:t>foobar</w:t>`, false},
		{"has_page_break", `<w:t>abc</w:t><w:lastRenderedPageBreak/><w:t>def</w:t>`, true},
		{"two_page_breaks", `<w:lastRenderedPageBreak/><w:lastRenderedPageBreak/>`, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			run := newRun(r, nil)
			if got := run.ContainsPageBreak(); got != tt.expected {
				t.Errorf("ContainsPageBreak() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_knows_its_underline_type (6 cases)
func TestRun_Underline_Getter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		isNil    bool
		isSingle bool
		isNone   bool
	}{
		{"absent", ``, true, false, false},
		{"u_no_val", `<w:rPr><w:u/></w:rPr>`, true, false, false},
		{"single", `<w:rPr><w:u w:val="single"/></w:rPr>`, false, true, false},
		{"none", `<w:rPr><w:u w:val="none"/></w:rPr>`, false, false, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.innerXml)
			run := newRun(r, nil)
			got, err := run.Underline()
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

// Mirrors Python: it_can_change_its_underline_type (10 cases)
func TestRun_Underline_Setter(t *testing.T) {
	t.Run("set_single", func(t *testing.T) {
		r := makeR(t, "")
		run := newRun(r, nil)
		v := UnderlineSingle()
		if err := run.SetUnderline(&v); err != nil {
			t.Fatal(err)
		}
		got, err := run.Underline()
		if err != nil {
			t.Fatal(err)
		}
		if got == nil || !got.IsSingle() {
			t.Error("expected single underline after set")
		}
	})
	t.Run("set_none", func(t *testing.T) {
		r := makeR(t, "")
		run := newRun(r, nil)
		v := UnderlineNone()
		if err := run.SetUnderline(&v); err != nil {
			t.Fatal(err)
		}
		got, err := run.Underline()
		if err != nil {
			t.Fatal(err)
		}
		if got == nil || !got.IsNone() {
			t.Error("expected none underline after set")
		}
	})
	t.Run("set_nil_inherits", func(t *testing.T) {
		r := makeR(t, `<w:rPr><w:u w:val="single"/></w:rPr>`)
		run := newRun(r, nil)
		if err := run.SetUnderline(nil); err != nil {
			t.Fatal(err)
		}
		got, err := run.Underline()
		if err != nil {
			t.Fatal(err)
		}
		if got != nil {
			t.Error("expected nil underline after set nil")
		}
	})
}

// Mirrors Python: it_provides_access_to_its_font
func TestRun_Font(t *testing.T) {
	r := makeR(t, "")
	run := newRun(r, nil)
	f := run.Font()
	if f == nil {
		t.Fatal("Font() returned nil")
	}
}

// Mirrors Python: it_can_mark_a_comment_reference_range
func TestRun_MarkCommentRange(t *testing.T) {
	pXml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:t>referenced text</w:t></w:r></w:p>`
	p := makeP(t, `<w:r><w:t>referenced text</w:t></w:r>`)
	_ = pXml
	runs := (&Paragraph{p: p}).Runs()
	if len(runs) == 0 {
		t.Fatal("expected at least one run")
	}
	run := runs[0]
	lastRun := runs[0]

	if err := run.MarkCommentRange(lastRun, 42); err != nil {
		t.Fatal(err)
	}
	// Verify commentRangeStart was inserted
	cs := p.RawElement().FindElement("commentRangeStart")
	if cs == nil {
		t.Error("expected w:commentRangeStart after MarkCommentRange")
	} else {
		if id := cs.SelectAttrValue("w:id", ""); id != "42" {
			t.Errorf("commentRangeStart w:id = %q, want %q", id, "42")
		}
	}
	// Verify commentRangeEnd was inserted
	ce := p.RawElement().FindElement("commentRangeEnd")
	if ce == nil {
		t.Error("expected w:commentRangeEnd after MarkCommentRange")
	} else {
		if id := ce.SelectAttrValue("w:id", ""); id != "42" {
			t.Errorf("commentRangeEnd w:id = %q, want %q", id, "42")
		}
	}
}
