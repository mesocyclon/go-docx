package docx

import (
	"strings"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// pagebreak_test.go — RenderedPageBreak (Batch 2)
// Mirrors Python: tests/text/test_pagebreak.py
// -----------------------------------------------------------------------

// helper: build a paragraph from XML, find lrpb elements
func buildParagraphFromXml(t *testing.T, xml string) (*Paragraph, []*RenderedPageBreak) {
	t.Helper()
	fullXml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + xml + `</w:p>`
	el, err := oxml.ParseXml([]byte(fullXml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	p := &oxml.CT_P{Element: oxml.WrapElement(el)}
	para := newParagraph(p, nil)
	rpbs := para.RenderedPageBreaks()
	return para, rpbs
}

// helper: normalize XML for comparison (strip namespace declarations, compress whitespace)
func normalizeXml(s string) string {
	s = strings.ReplaceAll(s, "\n", "")
	s = strings.ReplaceAll(s, "\t", "")
	// Collapse multiple spaces
	for strings.Contains(s, "  ") {
		s = strings.ReplaceAll(s, "  ", " ")
	}
	return strings.TrimSpace(s)
}

// Mirrors Python: it_can_split_off_the_preceding_paragraph_content_when_in_a_run
func TestRenderedPageBreak_PrecedingFragment_InRun(t *testing.T) {
	// <w:p><w:pPr><w:ind/></w:pPr><w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r><w:r><w:t>barfoo</w:t></w:r></w:p>
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r>` +
		`<w:r><w:t>barfoo</w:t></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].PrecedingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag == nil {
		t.Fatal("PrecedingParagraphFragment returned nil")
	}

	text := frag.Text()
	if text != "foo" {
		t.Errorf("preceding text = %q, want %q", text, "foo")
	}
}

// Mirrors Python: it_produces_None_for_preceding_when_leading
func TestRenderedPageBreak_PrecedingFragment_Leading(t *testing.T) {
	// page-break at start of paragraph → no preceding content
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:r><w:lastRenderedPageBreak/><w:t>foo</w:t><w:t>bar</w:t></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].PrecedingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag != nil {
		t.Error("expected nil for preceding fragment when page break is leading")
	}
}

// Mirrors Python: it_can_split_off_the_following_paragraph_content_when_in_a_run
func TestRenderedPageBreak_FollowingFragment_InRun(t *testing.T) {
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r>` +
		`<w:r><w:t>baz</w:t></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].FollowingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag == nil {
		t.Fatal("FollowingParagraphFragment returned nil")
	}

	text := frag.Text()
	// Following should contain "bar" + "baz" (= "barbaz")
	if text != "barbaz" {
		t.Errorf("following text = %q, want %q", text, "barbaz")
	}
}

// Mirrors Python: it_produces_None_for_following_when_trailing
func TestRenderedPageBreak_FollowingFragment_Trailing(t *testing.T) {
	// page-break is the last content → no following
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:r><w:t>foo</w:t><w:t>bar</w:t><w:lastRenderedPageBreak/></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].FollowingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag != nil {
		t.Error("expected nil for following fragment when page break is trailing")
	}
}

// Mirrors Python: and_it_can_split_off_the_preceding_paragraph_content_when_in_a_hyperlink
func TestRenderedPageBreak_PrecedingFragment_InHyperlink(t *testing.T) {
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:hyperlink><w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r></w:hyperlink>` +
		`<w:r><w:t>barfoo</w:t></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].PrecedingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag == nil {
		t.Fatal("PrecedingParagraphFragment returned nil")
	}

	// Python expected: w:p/(w:pPr/w:ind,w:hyperlink/w:r/(w:t"foo",w:t"bar"))
	// The entire hyperlink goes into preceding fragment (with lrpb removed)
	text := frag.Text()
	if text != "foobar" {
		t.Errorf("preceding hyperlink text = %q, want %q", text, "foobar")
	}
}

// Mirrors Python: and_it_can_split_off_the_following_paragraph_content_when_in_a_hyperlink
func TestRenderedPageBreak_FollowingFragment_InHyperlink(t *testing.T) {
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:hyperlink><w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r></w:hyperlink>` +
		`<w:r><w:t>baz</w:t></w:r>` +
		`<w:r><w:t>qux</w:t></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].FollowingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag == nil {
		t.Fatal("FollowingParagraphFragment returned nil")
	}

	// Python expected: w:p/(w:pPr/w:ind,w:r/w:t"baz",w:r/w:t"qux")
	// Hyperlink is excluded from following; only subsequent runs included
	text := frag.Text()
	if text != "bazqux" {
		t.Errorf("following hyperlink text = %q, want %q", text, "bazqux")
	}
}

// Mirrors Python: it_can_split_off_the_following_paragraph_content multi-run
func TestRenderedPageBreak_FollowingFragment_MultiRun(t *testing.T) {
	xml := `<w:pPr><w:ind/></w:pPr>` +
		`<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r>` +
		`<w:r><w:t>foo</w:t></w:r>`
	_, rpbs := buildParagraphFromXml(t, xml)
	if len(rpbs) == 0 {
		t.Fatal("no rendered page breaks found")
	}

	frag, err := rpbs[0].FollowingParagraphFragment()
	if err != nil {
		t.Fatal(err)
	}
	if frag == nil {
		t.Fatal("FollowingParagraphFragment returned nil")
	}

	text := frag.Text()
	if text != "barfoo" {
		t.Errorf("following text = %q, want %q", text, "barfoo")
	}
}

// Verify that RenderedPageBreaks list finds all breaks
func TestParagraph_RenderedPageBreaks_Count(t *testing.T) {
	tests := []struct {
		name string
		xml  string
		want int
	}{
		{"none", `<w:r><w:t>text</w:t></w:r>`, 0},
		{"one_in_run", `<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>`, 1},
		{"two_in_run", `<w:r><w:lastRenderedPageBreak/><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>`, 2},
		{"one_in_hyperlink", `<w:hyperlink><w:r><w:lastRenderedPageBreak/><w:t>b</w:t></w:r></w:hyperlink>`, 1},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, rpbs := buildParagraphFromXml(t, tt.xml)
			if len(rpbs) != tt.want {
				t.Errorf("RenderedPageBreaks() len = %d, want %d", len(rpbs), tt.want)
			}
		})
	}
}
