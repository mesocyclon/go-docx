package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// section_test.go — Section / Sections (Batch 1)
// Mirrors Python: tests/test_section.py
// -----------------------------------------------------------------------

// Mirrors Python: Sections.it_knows_how_many_sections
func TestSections_Len(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	if sections.Len() < 1 {
		t.Errorf("Sections.Len() = %d, want >= 1", sections.Len())
	}
}

// Mirrors Python: Sections.it_can_iterate
func TestSections_Iter(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	iter := sections.Iter()
	if len(iter) != sections.Len() {
		t.Errorf("len(Iter()) = %d, want %d", len(iter), sections.Len())
	}
	for i, s := range iter {
		if s == nil {
			t.Errorf("Iter()[%d] is nil", i)
		}
	}
}

// Mirrors Python: Sections.it_can_access_by_index
func TestSections_Get(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	s, err := sections.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	if s == nil {
		t.Error("Get(0) returned nil")
	}

	// Out of range
	_, err = sections.Get(999)
	if err == nil {
		t.Error("expected error for Get(999)")
	}
}

// Mirrors Python: it_knows_its_start_type / it_can_change
func TestSection_StartType(t *testing.T) {
	sectPr := makeSectPr(t, ``)
	sec := newSection(sectPr, nil)

	// Set to new page
	if err := sec.SetStartType(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}
	got, err := sec.StartType()
	if err != nil {
		t.Fatal(err)
	}
	if got != enum.WdSectionStartNewPage {
		t.Errorf("StartType() = %v, want %v", got, enum.WdSectionStartNewPage)
	}

	// Change to continuous
	if err := sec.SetStartType(enum.WdSectionStartContinuous); err != nil {
		t.Fatal(err)
	}
	got2, err := sec.StartType()
	if err != nil {
		t.Fatal(err)
	}
	if got2 != enum.WdSectionStartContinuous {
		t.Errorf("StartType() = %v, want %v", got2, enum.WdSectionStartContinuous)
	}
}

// Mirrors Python: it_knows_its_page_orientation / it_can_change
func TestSection_Orientation(t *testing.T) {
	sectPr := makeSectPr(t, `<w:pgSz w:w="12240" w:h="15840"/>`)
	sec := newSection(sectPr, nil)

	// Set to landscape
	if err := sec.SetOrientation(enum.WdOrientationLandscape); err != nil {
		t.Fatal(err)
	}
	got, err := sec.Orientation()
	if err != nil {
		t.Fatal(err)
	}
	if got != enum.WdOrientationLandscape {
		t.Errorf("Orientation() = %v, want LANDSCAPE", got)
	}

	// Set to portrait
	if err := sec.SetOrientation(enum.WdOrientationPortrait); err != nil {
		t.Fatal(err)
	}
	got2, err := sec.Orientation()
	if err != nil {
		t.Fatal(err)
	}
	if got2 != enum.WdOrientationPortrait {
		t.Errorf("Orientation() = %v, want PORTRAIT", got2)
	}
}

// Mirrors Python: it_knows_its_page_dimensions (complete set/get)
func TestSection_PageDimensions_SetGet(t *testing.T) {
	sectPr := makeSectPr(t, `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>`)
	sec := newSection(sectPr, nil)

	// Page Width
	w, err := sec.PageWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w == nil || *w != 12240 {
		t.Errorf("PageWidth = %v, want 12240", w)
	}
	newW := 15840
	if err := sec.SetPageWidth(&newW); err != nil {
		t.Fatal(err)
	}
	w2, _ := sec.PageWidth()
	if w2 == nil || *w2 != 15840 {
		t.Errorf("PageWidth after set = %v, want 15840", w2)
	}

	// Page Height
	h, err := sec.PageHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h == nil || *h != 15840 {
		t.Errorf("PageHeight = %v, want 15840", h)
	}
}

// Mirrors Python: margins set/get
func TestSection_Margins_SetGet(t *testing.T) {
	sectPr := makeSectPr(t, `<w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>`)
	sec := newSection(sectPr, nil)

	// Read
	top, err := sec.TopMargin()
	if err != nil {
		t.Fatal(err)
	}
	if top == nil || *top != 1440 {
		t.Errorf("TopMargin = %v, want 1440", top)
	}

	// Set
	v := 2000
	if err := sec.SetTopMargin(&v); err != nil {
		t.Fatal(err)
	}
	top2, _ := sec.TopMargin()
	if top2 == nil || *top2 != 2000 {
		t.Errorf("TopMargin after set = %v, want 2000", top2)
	}

	// Bottom
	bot, _ := sec.BottomMargin()
	if bot == nil || *bot != 1440 {
		t.Errorf("BottomMargin = %v, want 1440", bot)
	}

	// Left
	left, _ := sec.LeftMargin()
	if left == nil || *left != 1800 {
		t.Errorf("LeftMargin = %v, want 1800", left)
	}

	// Right
	right, _ := sec.RightMargin()
	if right == nil || *right != 1800 {
		t.Errorf("RightMargin = %v, want 1800", right)
	}
}

// Mirrors Python: it_knows_when_it_displays_a_distinct_first_page_header
func TestSection_DifferentFirstPageHeaderFooter(t *testing.T) {
	// Without titlePg
	sectPr1 := makeSectPr(t, ``)
	sec1 := newSection(sectPr1, nil)
	if sec1.DifferentFirstPageHeaderFooter() {
		t.Error("expected false when titlePg absent")
	}

	// With titlePg
	sectPr2 := makeSectPr(t, `<w:titlePg/>`)
	sec2 := newSection(sectPr2, nil)
	if !sec2.DifferentFirstPageHeaderFooter() {
		t.Error("expected true when titlePg present")
	}

	// Set
	if err := sec1.SetDifferentFirstPageHeaderFooter(true); err != nil {
		t.Fatal(err)
	}
	if !sec1.DifferentFirstPageHeaderFooter() {
		t.Error("expected true after SetDifferentFirstPageHeaderFooter(true)")
	}
}

// Helper: check Sections from a document with body-level sectPr
func makeSectionsDoc(t *testing.T, bodySectPrXml string) *oxml.CT_Document {
	t.Helper()
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body><w:p/><w:sectPr>` + bodySectPrXml + `</w:sectPr></w:body></w:document>`
	el := mustParseXml(t, xml)
	return &oxml.CT_Document{Element: *el}
}

// -----------------------------------------------------------------------
// Header/Footer.getOrAddDefinition — iterative prior-section walk
// -----------------------------------------------------------------------

// mustGetSection is a test helper that returns sections.Get(idx) or fails.
func mustGetSection(t *testing.T, doc *Document, idx int) *Section {
	t.Helper()
	sec, err := doc.Sections().Get(idx)
	if err != nil {
		t.Fatalf("Sections().Get(%d): %v", idx, err)
	}
	return sec
}

// Single section, no existing header definition.
// getOrAddDefinition should create a new definition (addDefinition path).
func TestHeaderGetOrAddDef_SingleSection_CreatesNew(t *testing.T) {
	doc := mustNewDoc(t)
	sec := mustGetSection(t, doc, 0)
	hdr := sec.Header()

	// Initially linked (no definition)
	if !hdr.IsLinkedToPrevious() {
		t.Fatal("expected IsLinkedToPrevious=true before any access")
	}

	// AddParagraph triggers getOrAddDefinition → addDefinition (no prior section)
	p, err := hdr.AddParagraph("created")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	if p.Text() != "created" {
		t.Errorf("paragraph text = %q, want %q", p.Text(), "created")
	}

	// Now the section has its own definition
	if hdr.IsLinkedToPrevious() {
		t.Error("expected IsLinkedToPrevious=false after AddParagraph")
	}
}

// Two sections: sec0 has a header, sec1 is linked.
// sec1.Header().Paragraphs() should resolve to sec0's header part.
func TestHeaderGetOrAddDef_WalksToPrior(t *testing.T) {
	doc := mustNewDoc(t)

	// Give sec0 a header with recognizable text.
	sec0 := mustGetSection(t, doc, 0)
	_, err := sec0.Header().AddParagraph("sec0-header")
	if err != nil {
		t.Fatalf("sec0 AddParagraph: %v", err)
	}

	// Add a second section — its header is linked (no own definition).
	sec1, err := doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		t.Fatalf("AddSection: %v", err)
	}
	hdr1 := sec1.Header()
	if !hdr1.IsLinkedToPrevious() {
		t.Fatal("sec1 header should be linked to previous")
	}

	// Paragraphs() walks to sec0's definition.
	paras, err := hdr1.Paragraphs()
	if err != nil {
		t.Fatalf("sec1 Paragraphs: %v", err)
	}
	found := false
	for _, p := range paras {
		if p.Text() == "sec0-header" {
			found = true
			break
		}
	}
	if !found {
		t.Error("sec1 header paragraphs should contain text from sec0 header")
	}
}

// Deep chain: 5 sections, only sec0 has a header definition, rest are linked.
// The last section should resolve all the way back to sec0 via the iterative
// loop (previously would be 4 recursive calls).
func TestHeaderGetOrAddDef_DeepChain(t *testing.T) {
	doc := mustNewDoc(t)

	sec0 := mustGetSection(t, doc, 0)
	_, err := sec0.Header().AddParagraph("deep-origin")
	if err != nil {
		t.Fatal(err)
	}

	// Add 4 more sections, all linked (no own header definition).
	for i := 0; i < 4; i++ {
		if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
			t.Fatalf("AddSection %d: %v", i+1, err)
		}
	}

	// Total 5 sections; last is index 4.
	lastSec := mustGetSection(t, doc, 4)
	hdr := lastSec.Header()
	if !hdr.IsLinkedToPrevious() {
		t.Fatal("last section header should be linked")
	}

	paras, err := hdr.Paragraphs()
	if err != nil {
		t.Fatalf("Paragraphs on last section: %v", err)
	}
	found := false
	for _, p := range paras {
		if p.Text() == "deep-origin" {
			found = true
			break
		}
	}
	if !found {
		t.Error("last section header should resolve to sec0's header (deep-origin)")
	}
}

// Three sections: sec0 and sec2 are linked, sec1 has its own definition.
// sec2 should resolve to sec1 (not walk past it to sec0).
func TestHeaderGetOrAddDef_StopsAtMiddle(t *testing.T) {
	doc := mustNewDoc(t)

	// Give the initial (only) section a header.
	initSec := mustGetSection(t, doc, 0)
	_, err := initSec.Header().AddParagraph("sec0-text")
	if err != nil {
		t.Fatal(err)
	}

	// Add two more sections (total 3).
	if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}
	if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}

	// Re-fetch after structural changes — AddSectionBreak shuffles sectPr
	// elements (clones the sentinel into a paragraph), so cached pointers
	// to the old sentinel would address the wrong section.
	sec1 := mustGetSection(t, doc, 1)
	sec2 := mustGetSection(t, doc, 2)

	// Unlink sec1 and give it its own header with distinct text.
	hdr1 := sec1.Header()
	if err := hdr1.SetIsLinkedToPrevious(false); err != nil {
		t.Fatal(err)
	}
	if _, err := hdr1.AddParagraph("sec1-text"); err != nil {
		t.Fatal(err)
	}

	// sec2 is linked — should resolve to sec1 (not walk past it).
	paras, err := sec2.Header().Paragraphs()
	if err != nil {
		t.Fatalf("Paragraphs on sec2: %v", err)
	}

	hasSec1 := false
	for _, p := range paras {
		if p.Text() == "sec1-text" {
			hasSec1 = true
		}
		if p.Text() == "sec0-text" {
			t.Error("sec2 header should NOT resolve to sec0; should stop at sec1")
		}
	}
	if !hasSec1 {
		t.Error("sec2 header should resolve to sec1's header (sec1-text)")
	}
}

// Footer: same iterative walk as Header.
// Two sections, sec0 has footer, sec1 linked → sec1 resolves to sec0.
func TestFooterGetOrAddDef_WalksToPrior(t *testing.T) {
	doc := mustNewDoc(t)

	sec0 := mustGetSection(t, doc, 0)
	_, err := sec0.Footer().AddParagraph("sec0-footer")
	if err != nil {
		t.Fatalf("sec0 Footer.AddParagraph: %v", err)
	}

	sec1, err := doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		t.Fatal(err)
	}
	ftr1 := sec1.Footer()
	if !ftr1.IsLinkedToPrevious() {
		t.Fatal("sec1 footer should be linked")
	}

	paras, err := ftr1.Paragraphs()
	if err != nil {
		t.Fatalf("sec1 Footer.Paragraphs: %v", err)
	}
	found := false
	for _, p := range paras {
		if p.Text() == "sec0-footer" {
			found = true
			break
		}
	}
	if !found {
		t.Error("sec1 footer should resolve to sec0's footer (sec0-footer)")
	}
}

// Footer deep chain: 4 sections linked, only sec0 has definition.
func TestFooterGetOrAddDef_DeepChain(t *testing.T) {
	doc := mustNewDoc(t)

	sec0 := mustGetSection(t, doc, 0)
	_, err := sec0.Footer().AddParagraph("footer-origin")
	if err != nil {
		t.Fatal(err)
	}

	for i := 0; i < 3; i++ {
		if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
			t.Fatalf("AddSection %d: %v", i+1, err)
		}
	}

	lastSec := mustGetSection(t, doc, 3)
	paras, err := lastSec.Footer().Paragraphs()
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, p := range paras {
		if p.Text() == "footer-origin" {
			found = true
			break
		}
	}
	if !found {
		t.Error("last section footer should resolve to sec0's footer")
	}
}

// Unlinking a previously linked header triggers addDefinition on that section,
// giving it its own independent header part.
func TestHeaderGetOrAddDef_UnlinkCreatesOwnDefinition(t *testing.T) {
	doc := mustNewDoc(t)

	// Give the initial section a header.
	initSec := mustGetSection(t, doc, 0)
	_, err := initSec.Header().AddParagraph("original")
	if err != nil {
		t.Fatal(err)
	}

	// Add a second section.
	if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}

	// Re-fetch — AddSectionBreak clones the sentinel, so cached pointers
	// to the old sectPr now address the wrong section.
	sec0 := mustGetSection(t, doc, 0)
	sec1 := mustGetSection(t, doc, 1)
	hdr1 := sec1.Header()

	// Unlink: gives sec1 its own header definition.
	if err := hdr1.SetIsLinkedToPrevious(false); err != nil {
		t.Fatalf("SetIsLinkedToPrevious(false): %v", err)
	}
	if hdr1.IsLinkedToPrevious() {
		t.Error("expected not linked after SetIsLinkedToPrevious(false)")
	}

	// Adding text to sec1's header should NOT affect sec0.
	_, err = hdr1.AddParagraph("sec1-own")
	if err != nil {
		t.Fatal(err)
	}

	// sec0 header should still have only "original".
	sec0paras, err := sec0.Header().Paragraphs()
	if err != nil {
		t.Fatal(err)
	}
	for _, p := range sec0paras {
		if p.Text() == "sec1-own" {
			t.Error("sec0 header should NOT contain sec1's text after unlinking")
		}
	}
}
