package docx

import (
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// testutil_test.go — Additional shared test helpers for Batch 1 tests
// -----------------------------------------------------------------------

// wrapAsCT_P wraps an existing etree.Element as a CT_P.
func wrapAsCT_P(t *testing.T, el *etree.Element) *oxml.CT_P {
	t.Helper()
	return &oxml.CT_P{Element: oxml.WrapElement(el)}
}

// intPtr returns a pointer to int.
func intPtr(v int) *int { return &v }

// makeSectPr creates a <w:sectPr> element from inner XML.
func makeSectPr(t *testing.T, innerXml string) *oxml.CT_SectPr {
	t.Helper()
	xml := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:sectPr>`
	el := mustParseXml(t, xml)
	return &oxml.CT_SectPr{Element: *el}
}

// makeStyles creates a <w:styles> element from inner XML.
func makeStyles(t *testing.T, innerXml string) *oxml.CT_Styles {
	t.Helper()
	xml := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:styles>`
	el := mustParseXml(t, xml)
	return &oxml.CT_Styles{Element: *el}
}

// makeComments creates a <w:comments> element from inner XML.
func makeComments(t *testing.T, innerXml string) *oxml.CT_Comments {
	t.Helper()
	xml := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:comments>`
	el := mustParseXml(t, xml)
	return &oxml.CT_Comments{Element: *el}
}

// makePPr creates a <w:pPr> element from inner XML.
func makePPr(t *testing.T, innerXml string) *oxml.CT_PPr {
	t.Helper()
	xml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:pPr>`
	el := mustParseXml(t, xml)
	return &oxml.CT_PPr{Element: *el}
}

// compareBoolPtr checks two *bool pointers for equality.
func compareBoolPtr(t *testing.T, name string, got, want *bool) {
	t.Helper()
	if got == nil && want == nil {
		return
	}
	if got == nil || want == nil {
		t.Errorf("%s = %v, want %v", name, ptrBoolStr(got), ptrBoolStr(want))
		return
	}
	if *got != *want {
		t.Errorf("%s = %v, want %v", name, *got, *want)
	}
}

func ptrBoolStr(p *bool) string {
	if p == nil {
		return "<nil>"
	}
	if *p {
		return "true"
	}
	return "false"
}
