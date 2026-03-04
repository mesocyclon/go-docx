package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// section_hdrftr_test.go — Section header/footer accessors (Batch 2)
// Mirrors Python: tests/test_section.py — _BaseHeaderFooter, header/footer access
// -----------------------------------------------------------------------

// helper: build a sectPr from inner XML
func makeSectPrFromXml(t *testing.T, innerXml string) *oxml.CT_SectPr {
	t.Helper()
	xml := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		innerXml + `</w:sectPr>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	return &oxml.CT_SectPr{Element: oxml.WrapElement(el)}
}

// Mirrors Python: it_knows_when_its_linked_to_the_previous_header_or_footer
// IsLinkedToPrevious = true when no headerReference exists for that type
func TestHeader_IsLinkedToPrevious(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		index    enum.WdHeaderFooterIndex
		want     bool
	}{
		{
			"no_ref_primary",
			"", // no headerReference at all
			enum.WdHeaderFooterIndexPrimary,
			true, // linked = no definition
		},
		{
			"has_ref_primary",
			`<w:headerReference w:type="default" r:id="rId1"/>`,
			enum.WdHeaderFooterIndexPrimary,
			false, // has definition → not linked
		},
		{
			"has_ref_first_page",
			`<w:headerReference w:type="first" r:id="rId2"/>`,
			enum.WdHeaderFooterIndexFirstPage,
			false,
		},
		{
			"no_ref_first_page_when_default_exists",
			`<w:headerReference w:type="default" r:id="rId1"/>`,
			enum.WdHeaderFooterIndexFirstPage,
			true, // no first-page ref → linked
		},
		{
			"has_ref_even_page",
			`<w:headerReference w:type="even" r:id="rId3"/>`,
			enum.WdHeaderFooterIndexEvenPage,
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sectPr := makeSectPrFromXml(t, tt.innerXml)
			header := newHeader(sectPr, nil, tt.index)
			got := header.IsLinkedToPrevious()
			if got != tt.want {
				t.Errorf("IsLinkedToPrevious() = %v, want %v", got, tt.want)
			}
		})
	}
}

// Mirrors Python footer: same logic
func TestFooter_IsLinkedToPrevious(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		index    enum.WdHeaderFooterIndex
		want     bool
	}{
		{
			"no_ref",
			"",
			enum.WdHeaderFooterIndexPrimary,
			true,
		},
		{
			"has_ref",
			`<w:footerReference w:type="default" r:id="rId1"/>`,
			enum.WdHeaderFooterIndexPrimary,
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sectPr := makeSectPrFromXml(t, tt.innerXml)
			footer := newFooter(sectPr, nil, tt.index)
			got := footer.IsLinkedToPrevious()
			if got != tt.want {
				t.Errorf("IsLinkedToPrevious() = %v, want %v", got, tt.want)
			}
		})
	}
}

// Mirrors Python: it_can_change_whether_the_document_has_distinct_odd_and_even_headers (4 transitions)
func TestSection_DifferentFirstPageHeaderFooter_Setter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		value    bool
		wantPres bool // titlePg element should be present
	}{
		{"empty_set_true", "", true, true},
		{"titlePg_set_false", "<w:titlePg/>", false, false},
		{"titlePg_val1_set_true", `<w:titlePg w:val="1"/>`, true, true},
		{"titlePg_valoff_set_false", `<w:titlePg w:val="off"/>`, false, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sectPr := makeSectPrFromXml(t, tt.innerXml)
			section := newSection(sectPr, nil)
			if err := section.SetDifferentFirstPageHeaderFooter(tt.value); err != nil {
				t.Fatalf("SetDifferentFirstPageHeaderFooter: %v", err)
			}
			got := section.DifferentFirstPageHeaderFooter()
			if got != tt.value {
				t.Errorf("DifferentFirstPageHeaderFooter() = %v after set %v", got, tt.value)
			}
		})
	}
}

// Verify header/footer accessor methods return non-nil objects
func TestSection_HeaderFooter_Accessors(t *testing.T) {
	sectPr := makeSectPrFromXml(t, "")
	section := newSection(sectPr, nil)

	// All six accessor methods should return non-nil
	if h := section.Header(); h == nil {
		t.Error("Header() returned nil")
	}
	if f := section.Footer(); f == nil {
		t.Error("Footer() returned nil")
	}
	if h := section.EvenPageHeader(); h == nil {
		t.Error("EvenPageHeader() returned nil")
	}
	if f := section.EvenPageFooter(); f == nil {
		t.Error("EvenPageFooter() returned nil")
	}
	if h := section.FirstPageHeader(); h == nil {
		t.Error("FirstPageHeader() returned nil")
	}
	if f := section.FirstPageFooter(); f == nil {
		t.Error("FirstPageFooter() returned nil")
	}
}

// Mirrors Python: it_knows_when_it_displays_a_distinct_first_page_header (getter)
func TestSection_DifferentFirstPageHeaderFooter_Getter(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		want     bool
	}{
		{"absent", "", false},
		{"present_no_val", "<w:titlePg/>", true},
		{"val_true", `<w:titlePg w:val="true"/>`, true},
		{"val_1", `<w:titlePg w:val="1"/>`, true},
		{"val_false", `<w:titlePg w:val="false"/>`, false},
		{"val_0", `<w:titlePg w:val="0"/>`, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sectPr := makeSectPrFromXml(t, tt.innerXml)
			section := newSection(sectPr, nil)
			got := section.DifferentFirstPageHeaderFooter()
			if got != tt.want {
				t.Errorf("DifferentFirstPageHeaderFooter() = %v, want %v", got, tt.want)
			}
		})
	}
}
