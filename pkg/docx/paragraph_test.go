package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// paragraph_test.go — Paragraph (Batch 1)
// Mirrors Python: tests/text/test_paragraph.py
// -----------------------------------------------------------------------

func wdAlignPtr(v enum.WdParagraphAlignment) *enum.WdParagraphAlignment { return &v }

// Mirrors Python: it_knows_its_alignment_value
func TestParagraph_Alignment_Getter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected *enum.WdParagraphAlignment
	}{
		{"center", `<w:pPr><w:jc w:val="center"/></w:pPr>`, wdAlignPtr(enum.WdParagraphAlignmentCenter)},
		{"nil_when_absent", ``, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			got, err := para.Alignment()
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if got == nil && tt.expected == nil {
				return
			}
			if got == nil || tt.expected == nil || *got != *tt.expected {
				t.Errorf("Alignment() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_can_change_its_alignment_value (all 4 transitions)
func TestParagraph_Alignment_Setter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		value    *enum.WdParagraphAlignment
		wantJc   string // expected w:val on w:jc; "" means jc should be absent
	}{
		// none → set
		{"none_to_left", ``, wdAlignPtr(enum.WdParagraphAlignmentLeft), "left"},
		// set → change
		{"left_to_center", `<w:pPr><w:jc w:val="left"/></w:pPr>`, wdAlignPtr(enum.WdParagraphAlignmentCenter), "center"},
		// set → remove
		{"left_to_nil", `<w:pPr><w:jc w:val="left"/></w:pPr>`, nil, ""},
		// none → remove (noop)
		{"none_to_nil", ``, nil, ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			if err := para.SetAlignment(tt.value); err != nil {
				t.Fatalf("SetAlignment: %v", err)
			}
			// Verify: re-read alignment
			got, err := para.Alignment()
			if err != nil {
				t.Fatalf("Alignment() after set: %v", err)
			}
			if tt.wantJc == "" {
				if got != nil {
					t.Errorf("expected nil alignment, got %v", *got)
				}
			} else {
				if got == nil {
					t.Fatalf("expected non-nil alignment %q", tt.wantJc)
				}
				gotXml, _ := got.ToXml()
				if gotXml != tt.wantJc {
					t.Errorf("alignment XML = %q, want %q", gotXml, tt.wantJc)
				}
			}
		})
	}
}

// Mirrors Python: it_knows_the_text_it_contains (10 cases)
func TestParagraph_Text_AllCases(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected string
	}{
		{"empty_p", ``, ""},
		{"empty_run", `<w:r/>`, ""},
		{"empty_t", `<w:r><w:t/></w:r>`, ""},
		{"simple_text", `<w:r><w:t>foo</w:t></w:r>`, "foo"},
		{"two_t_elements", `<w:r><w:t>foo</w:t><w:t>bar</w:t></w:r>`, "foobar"},
		{"text_with_space", `<w:r><w:t>fo </w:t><w:t>bar</w:t></w:r>`, "fo bar"},
		{"tab_in_run", `<w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>`, "foo\tbar"},
		{"br_in_run", `<w:r><w:t>foo</w:t><w:br/><w:t>bar</w:t></w:r>`, "foo\nbar"},
		{"cr_in_run", `<w:r><w:t>foo</w:t><w:cr/><w:t>bar</w:t></w:r>`, "foo\nbar"},
		{"hyperlink_text",
			`<w:r><w:t>click </w:t></w:r>` +
				`<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId6"><w:r><w:t>here</w:t></w:r></w:hyperlink>` +
				`<w:r><w:t> for more</w:t></w:r>`,
			"click here for more"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			if got := para.Text(); got != tt.expected {
				t.Errorf("Text() = %q, want %q", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_can_replace_the_text_it_contains
func TestParagraph_SetText(t *testing.T) {
	p := makeP(t, `<w:r><w:t>old text</w:t></w:r>`)
	para := newParagraph(p, nil)

	if err := para.SetText("new text"); err != nil {
		t.Fatal(err)
	}
	if got := para.Text(); got != "new text" {
		t.Errorf("Text() after SetText = %q, want %q", got, "new text")
	}
}

// Mirrors Python: it_can_remove_its_content_while_preserving_formatting (4 cases)
func TestParagraph_Clear(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		wantText string
		wantPPr  bool // whether pPr should survive
	}{
		{"empty", ``, "", false},
		{"pPr_only", `<w:pPr/>`, "", true},
		{"text_removed", `<w:r><w:t>foobar</w:t></w:r>`, "", false},
		{"pPr_preserved_text_removed", `<w:pPr/><w:r><w:t>foobar</w:t></w:r>`, "", true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			para.Clear()
			if got := para.Text(); got != tt.wantText {
				t.Errorf("Text() after Clear = %q, want %q", got, tt.wantText)
			}
			// Check pPr preservation
			pPr := p.PPr()
			hasPPr := pPr != nil
			if hasPPr != tt.wantPPr {
				t.Errorf("pPr present = %v, want %v", hasPPr, tt.wantPPr)
			}
		})
	}
}

// Mirrors Python: it_knows_whether_it_contains_a_page_break
func TestParagraph_ContainsPageBreak(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected bool
	}{
		{"no_break_empty_run", `<w:r/>`, false},
		{"no_break_text_only", `<w:r><w:t>foobar</w:t></w:r>`, false},
		{"has_break_in_run", `<w:r><w:lastRenderedPageBreak/><w:lastRenderedPageBreak/></w:r>`, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			if got := para.ContainsPageBreak(); got != tt.expected {
				t.Errorf("ContainsPageBreak() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_provides_access_to_the_hyperlinks_it_contains
func TestParagraph_Hyperlinks(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		count    int
	}{
		{"no_hyperlinks_empty", ``, 0},
		{"no_hyperlinks_run_only", `<w:r/>`, 0},
		{"one_hyperlink", `<w:hyperlink/>`, 1},
		{"mixed_run_hyperlink_run", `<w:r/><w:hyperlink/><w:r/>`, 1},
		{"two_hyperlinks", `<w:r/><w:hyperlink/><w:r/><w:hyperlink/>`, 2},
		{"hyperlink_first_and_last", `<w:hyperlink/><w:r/><w:hyperlink/><w:r/>`, 2},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			hls := para.Hyperlinks()
			if len(hls) != tt.count {
				t.Errorf("len(Hyperlinks) = %d, want %d", len(hls), tt.count)
			}
		})
	}
}

// Mirrors Python: it_can_iterate_its_inner_content_items
func TestParagraph_IterInnerContent(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected []string // "Run" or "Hyperlink"
	}{
		{"empty", ``, nil},
		{"run_only", `<w:r/>`, []string{"Run"}},
		{"hyperlink_only", `<w:hyperlink/>`, []string{"Hyperlink"}},
		{"run_hyperlink_run", `<w:r/><w:hyperlink/><w:r/>`, []string{"Run", "Hyperlink", "Run"}},
		{"hyperlink_run_hyperlink", `<w:hyperlink/><w:r/><w:hyperlink/>`, []string{"Hyperlink", "Run", "Hyperlink"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := makeP(t, tt.innerXml)
			para := newParagraph(p, nil)
			items := para.IterInnerContent()
			got := make([]string, len(items))
			for i, item := range items {
				if item.IsRun() {
					got[i] = "Run"
				} else if item.IsHyperlink() {
					got[i] = "Hyperlink"
				}
			}
			if len(got) != len(tt.expected) {
				t.Fatalf("len(items) = %d, want %d; got %v", len(got), len(tt.expected), got)
			}
			for i := range got {
				if got[i] != tt.expected[i] {
					t.Errorf("item[%d] = %q, want %q", i, got[i], tt.expected[i])
				}
			}
		})
	}
}

// Mirrors Python: it_provides_access_to_its_paragraph_format
func TestParagraph_ParagraphFormat(t *testing.T) {
	p := makeP(t, ``)
	para := newParagraph(p, nil)
	pf := para.ParagraphFormat()
	if pf == nil {
		t.Fatal("ParagraphFormat() returned nil")
	}
}

// Mirrors Python: it_can_insert_a_paragraph_before_itself
func TestParagraph_InsertParagraphBefore(t *testing.T) {
	// Construct a body with a single paragraph
	bodyXml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:t>original</w:t></w:r></w:p></w:body>`
	el := mustParseXml(t, bodyXml)
	// Find the w:p inside the body
	pElem := el.RawElement().FindElement("p")
	if pElem == nil {
		t.Fatal("could not find w:p in body")
	}

	pCT := wrapAsCT_P(t, pElem)
	para := newParagraph(pCT, nil)

	newPara, err := para.InsertParagraphBefore("inserted")
	if err != nil {
		t.Fatalf("InsertParagraphBefore: %v", err)
	}
	if newPara == nil {
		t.Fatal("InsertParagraphBefore returned nil")
	}
	if newPara.Text() != "inserted" {
		t.Errorf("new paragraph text = %q, want %q", newPara.Text(), "inserted")
	}
	// Original paragraph should still be accessible
	if para.Text() != "original" {
		t.Errorf("original paragraph text = %q, want %q", para.Text(), "original")
	}
}
