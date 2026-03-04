package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// hyperlink_test.go — Hyperlink (Batch 1)
// Mirrors Python: tests/text/test_hyperlink.py
// -----------------------------------------------------------------------

func makeHyperlink(t *testing.T, xml string) *Hyperlink {
	t.Helper()
	fullXml := `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		xml + `</w:hyperlink>`
	el := mustParseXml(t, fullXml)
	h := &oxml.CT_Hyperlink{Element: *el}
	return newHyperlink(h, nil)
}

// Mirrors Python: it_knows_the_hyperlink_address (internal — no rId)
func TestHyperlink_Address_Internal(t *testing.T) {
	hl := makeHyperlink(t, ` w:anchor="bookmark1"><w:r><w:t>See here</w:t></w:r>`)
	// No rId means address is ""
	if hl.Address() != "" {
		t.Errorf("Address() = %q, want empty for internal link", hl.Address())
	}
}

// Mirrors Python: it_knows_the_visible_text
func TestHyperlink_Text(t *testing.T) {
	tests := []struct {
		name     string
		xml      string
		expected string
	}{
		{"single_run", `><w:r><w:t>Click here</w:t></w:r>`, "Click here"},
		{"two_runs", `><w:r><w:t>Click </w:t></w:r><w:r><w:t>here</w:t></w:r>`, "Click here"},
		{"empty", `>`, ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			hl := makeHyperlink(t, tt.xml)
			if got := hl.Text(); got != tt.expected {
				t.Errorf("Text() = %q, want %q", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_knows_the_full_url (fragment + address concat)
func TestHyperlink_URL_WithFragment(t *testing.T) {
	// Internal link with anchor/fragment → URL is just fragment
	hl := makeHyperlink(t, ` w:anchor="_Toc147925734"><w:r><w:t>See section 1</w:t></w:r>`)
	if hl.Fragment() != "_Toc147925734" {
		t.Errorf("Fragment() = %q, want %q", hl.Fragment(), "_Toc147925734")
	}
	// URL should be empty since address is empty
	if hl.URL() != "" {
		t.Errorf("URL() = %q, want empty (no address)", hl.URL())
	}
}

// Mirrors Python: it_knows_whether_it_contains_a_page_break
func TestHyperlink_ContainsPageBreak(t *testing.T) {
	// No page break
	hl := makeHyperlink(t, `><w:r><w:t>no break</w:t></w:r>`)
	runs := hl.Runs()
	containsBreak := false
	for _, r := range runs {
		if r.ContainsPageBreak() {
			containsBreak = true
		}
	}
	if containsBreak {
		t.Error("expected no page break in hyperlink")
	}
}

// Mirrors Python: Hyperlink.runs
func TestHyperlink_Runs_Multiple(t *testing.T) {
	hl := makeHyperlink(t, `><w:r><w:t>word1</w:t></w:r><w:r><w:t> word2</w:t></w:r>`)
	runs := hl.Runs()
	if len(runs) != 2 {
		t.Fatalf("len(Runs) = %d, want 2", len(runs))
	}
	if runs[0].Text() != "word1" {
		t.Errorf("runs[0].Text() = %q, want %q", runs[0].Text(), "word1")
	}
	if runs[1].Text() != " word2" {
		t.Errorf("runs[1].Text() = %q, want %q", runs[1].Text(), " word2")
	}
}
