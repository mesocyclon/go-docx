package docx

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// batch3_test.go — Batch 3: P3 edge cases (~15 tests)
//
// 19. Section inner content iteration
// 20. OPC core properties lazy creation
// 21. Document additions (AddComment flow, AddPicture flow, CoreProperties)
//
// Mirrors Python:
//   tests/test_section.py  → it_can_iterate_its_inner_content
//   tests/opc/test_package.py → it_provides_access_to_the_core_properties,
//                                it_creates_a_default_core_props_part_if_none_present
//   tests/test_document.py → it_can_add_a_comment, it_can_add_a_picture,
//                             it_provides_access_to_its_core_properties,
//                             _Body.it_can_clear_itself_of_all_content_it_holds
// -----------------------------------------------------------------------

// =======================================================================
// 19. Section inner content iteration
// =======================================================================

// makeMultiSectionDoc builds a 3-section document for IterInnerContent tests.
//
// Layout (mirrors Python fixture sct-inner-content.docx):
//
//	Section 0: <w:p>P1</w:p>  <w:tbl>T2</w:tbl>  <w:p pPr/sectPr>P3</w:p>
//	Section 1: <w:tbl>T4</w:tbl>  <w:p>P5</w:p>  <w:p pPr/sectPr>P6</w:p>
//	Section 2: <w:p>P7</w:p>  <w:p>P8</w:p>  <w:p>P9</w:p>  <w:sectPr/>
func makeMultiSectionDoc(t *testing.T) *oxml.CT_Document {
	t.Helper()
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body>` +

		// --- Section 0 ---
		`<w:p><w:r><w:t>P1</w:t></w:r></w:p>` +
		`<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>T2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>P3</w:t></w:r></w:p>` +

		// --- Section 1 ---
		`<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>T4</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:p><w:r><w:t>P5</w:t></w:r></w:p>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>P6</w:t></w:r></w:p>` +

		// --- Section 2 ---
		`<w:p><w:r><w:t>P7</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>P8</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>P9</w:t></w:r></w:p>` +
		`<w:sectPr/>` +

		`</w:body></w:document>`
	el := mustParseXml(t, xml)
	return &oxml.CT_Document{Element: *el}
}

// Mirrors Python: Section.it_can_iterate_its_inner_content (section 0)
func TestSection_IterInnerContent_Section0(t *testing.T) {
	docElm := makeMultiSectionDoc(t)
	sections := newSections(docElm, nil)
	if sections.Len() != 3 {
		t.Fatalf("precondition: Sections.Len() = %d, want 3", sections.Len())
	}

	sec, err := sections.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	items := sec.IterInnerContent()

	if len(items) != 3 {
		t.Fatalf("section 0: len(IterInnerContent()) = %d, want 3", len(items))
	}

	// item 0: paragraph "P1"
	if !items[0].IsParagraph() {
		t.Error("section 0, item 0: expected paragraph")
	}
	if got := items[0].Paragraph().Text(); got != "P1" {
		t.Errorf("section 0, item 0: text = %q, want %q", got, "P1")
	}

	// item 1: table with cell text "T2"
	if !items[1].IsTable() {
		t.Error("section 0, item 1: expected table")
	}
	if got := cellText(t, items[1].Table(), 0, 0); got != "T2" {
		t.Errorf("section 0, item 1: cell text = %q, want %q", got, "T2")
	}

	// item 2: paragraph "P3" (the section-ending paragraph)
	if !items[2].IsParagraph() {
		t.Error("section 0, item 2: expected paragraph")
	}
	if got := items[2].Paragraph().Text(); got != "P3" {
		t.Errorf("section 0, item 2: text = %q, want %q", got, "P3")
	}
}

// Mirrors Python: Section.it_can_iterate_its_inner_content (section 1)
func TestSection_IterInnerContent_Section1(t *testing.T) {
	docElm := makeMultiSectionDoc(t)
	sections := newSections(docElm, nil)

	sec, err := sections.Get(1)
	if err != nil {
		t.Fatal(err)
	}
	items := sec.IterInnerContent()

	if len(items) != 3 {
		t.Fatalf("section 1: len(IterInnerContent()) = %d, want 3", len(items))
	}

	// item 0: table with cell text "T4"
	if !items[0].IsTable() {
		t.Error("section 1, item 0: expected table")
	}
	if got := cellText(t, items[0].Table(), 0, 0); got != "T4" {
		t.Errorf("section 1, item 0: cell text = %q, want %q", got, "T4")
	}

	// item 1: paragraph "P5"
	if !items[1].IsParagraph() {
		t.Error("section 1, item 1: expected paragraph")
	}
	if got := items[1].Paragraph().Text(); got != "P5" {
		t.Errorf("section 1, item 1: text = %q, want %q", got, "P5")
	}

	// item 2: paragraph "P6" (section-ending)
	if !items[2].IsParagraph() {
		t.Error("section 1, item 2: expected paragraph")
	}
	if got := items[2].Paragraph().Text(); got != "P6" {
		t.Errorf("section 1, item 2: text = %q, want %q", got, "P6")
	}
}

// Mirrors Python: Section.it_can_iterate_its_inner_content (section 2)
func TestSection_IterInnerContent_Section2(t *testing.T) {
	docElm := makeMultiSectionDoc(t)
	sections := newSections(docElm, nil)

	sec, err := sections.Get(2)
	if err != nil {
		t.Fatal(err)
	}
	items := sec.IterInnerContent()

	if len(items) != 3 {
		t.Fatalf("section 2: len(IterInnerContent()) = %d, want 3", len(items))
	}

	for i, expected := range []string{"P7", "P8", "P9"} {
		if !items[i].IsParagraph() {
			t.Errorf("section 2, item %d: expected paragraph", i)
			continue
		}
		if got := items[i].Paragraph().Text(); got != expected {
			t.Errorf("section 2, item %d: text = %q, want %q", i, got, expected)
		}
	}
}

// Test: single-section document (body-level sectPr only)
func TestSection_IterInnerContent_SingleSection(t *testing.T) {
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body>` +
		`<w:p><w:r><w:t>Only</w:t></w:r></w:p>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:sectPr/>` +
		`</w:body></w:document>`
	el := mustParseXml(t, xml)
	docElm := &oxml.CT_Document{Element: *el}
	sections := newSections(docElm, nil)

	if sections.Len() != 1 {
		t.Fatalf("Len() = %d, want 1", sections.Len())
	}
	sec, _ := sections.Get(0)
	items := sec.IterInnerContent()

	if len(items) != 2 {
		t.Fatalf("len = %d, want 2", len(items))
	}
	if !items[0].IsParagraph() {
		t.Error("item 0: expected paragraph")
	}
	if got := items[0].Paragraph().Text(); got != "Only" {
		t.Errorf("item 0: text = %q, want %q", got, "Only")
	}
	if !items[1].IsTable() {
		t.Error("item 1: expected table")
	}
}

// Test: empty section (section break paragraph with no preceding content)
func TestSection_IterInnerContent_EmptySection(t *testing.T) {
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` + // section 0: just the break para
		`<w:p><w:r><w:t>After</w:t></w:r></w:p>` + // section 1
		`<w:sectPr/>` +
		`</w:body></w:document>`
	el := mustParseXml(t, xml)
	docElm := &oxml.CT_Document{Element: *el}
	sections := newSections(docElm, nil)

	if sections.Len() != 2 {
		t.Fatalf("Len() = %d, want 2", sections.Len())
	}

	// Section 0: just the delimiter paragraph (no text, but the para is included)
	sec0, _ := sections.Get(0)
	items0 := sec0.IterInnerContent()
	if len(items0) != 1 {
		t.Fatalf("section 0: len = %d, want 1", len(items0))
	}
	if !items0[0].IsParagraph() {
		t.Error("section 0, item 0: expected paragraph")
	}

	// Section 1: one paragraph
	sec1, _ := sections.Get(1)
	items1 := sec1.IterInnerContent()
	if len(items1) != 1 {
		t.Fatalf("section 1: len = %d, want 1", len(items1))
	}
	if got := items1[0].Paragraph().Text(); got != "After" {
		t.Errorf("section 1, item 0: text = %q, want %q", got, "After")
	}
}

// Test: section with only tables, no paragraphs
func TestSection_IterInnerContent_TablesOnly(t *testing.T) {
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` + // delimiter para
		`<w:sectPr/>` + // body-level
		`</w:body></w:document>`
	el := mustParseXml(t, xml)
	docElm := &oxml.CT_Document{Element: *el}
	sections := newSections(docElm, nil)

	sec0, _ := sections.Get(0)
	items := sec0.IterInnerContent()

	// 2 tables + 1 delimiter paragraph
	if len(items) != 3 {
		t.Fatalf("section 0: len = %d, want 3", len(items))
	}
	if !items[0].IsTable() {
		t.Error("item 0: expected table")
	}
	if !items[1].IsTable() {
		t.Error("item 1: expected table")
	}
	if !items[2].IsParagraph() {
		t.Error("item 2: expected paragraph (section delimiter)")
	}
}

// =======================================================================
// 20. OPC core properties lazy creation
// =======================================================================

// Mirrors Python: Package.it_provides_access_to_the_core_properties
func TestDocument_CoreProperties_Access(t *testing.T) {
	doc := mustNewDoc(t)
	cp, err := doc.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties() error: %v", err)
	}
	if cp == nil {
		t.Fatal("CoreProperties() returned nil")
	}
}

// Mirrors Python: Package.it_creates_a_default_core_props_part_if_none_present
//
// A new document should lazily create a default CorePropertiesPart when
// CoreProperties() is called, even though it wasn't explicitly added.
func TestDocument_CoreProperties_LazyCreation(t *testing.T) {
	doc := mustNewDoc(t)

	// First access — triggers lazy creation
	cp1, err := doc.CoreProperties()
	if err != nil {
		t.Fatalf("first call: %v", err)
	}
	if cp1 == nil {
		t.Fatal("first call returned nil")
	}

	// Second access — returns same data (idempotent)
	cp2, err := doc.CoreProperties()
	if err != nil {
		t.Fatalf("second call: %v", err)
	}
	if cp2 == nil {
		t.Fatal("second call returned nil")
	}

	// Both should report same author (both come from the same underlying part)
	if cp1.Author() != cp2.Author() {
		t.Errorf("inconsistent Author: %q vs %q", cp1.Author(), cp2.Author())
	}
}

// Verify that CoreProperties are accessible and writable on a new document.
func TestDocument_CoreProperties_DefaultValues(t *testing.T) {
	doc := mustNewDoc(t)
	cp, err := doc.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}

	// Default template may have arbitrary values; just verify getters don't panic
	_ = cp.Author()
	_ = cp.Title()
	_ = cp.LastModifiedBy()
	_ = cp.Revision()

	// Verify setters work on the default doc
	if err := cp.SetTitle("NewTitle"); err != nil {
		t.Fatalf("SetTitle: %v", err)
	}
	if got := cp.Title(); got != "NewTitle" {
		t.Errorf("Title after set = %q, want %q", got, "NewTitle")
	}
}

// Verify mutations to CoreProperties are visible after save/reload.
func TestDocument_CoreProperties_RoundTrip(t *testing.T) {
	doc := mustNewDoc(t)
	cp, err := doc.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}

	if err := cp.SetAuthor("TestAuthor"); err != nil {
		t.Fatalf("SetAuthor: %v", err)
	}
	if err := cp.SetTitle("TestTitle"); err != nil {
		t.Fatalf("SetTitle: %v", err)
	}

	// Save
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Reload
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	cp2, err := doc2.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties after reload: %v", err)
	}

	if got := cp2.Author(); got != "TestAuthor" {
		t.Errorf("Author after round-trip = %q, want %q", got, "TestAuthor")
	}
	if got := cp2.Title(); got != "TestTitle" {
		t.Errorf("Title after round-trip = %q, want %q", got, "TestTitle")
	}
}

// =======================================================================
// 21. Document additions
// =======================================================================

// Mirrors Python: Document.it_can_add_a_comment
//
// Verifies the full flow: Document.AddComment() creates a comment,
// assigns it to the comments collection, and marks the run range.
func TestDocument_AddComment_Flow(t *testing.T) {
	doc := mustNewDoc(t)

	// Add a paragraph with a run to anchor the comment
	para, err := doc.AddParagraph("referenced text")
	if err != nil {
		t.Fatal(err)
	}
	runs := para.Runs()
	if len(runs) == 0 {
		t.Fatal("expected at least one run after AddParagraph")
	}

	initials := "JD"
	comment, err := doc.AddComment(runs, "Review this", "John Doe", &initials)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	if comment == nil {
		t.Fatal("AddComment returned nil")
	}

	// Verify comment was added to the collection
	comments, err := doc.Comments()
	if err != nil {
		t.Fatal(err)
	}
	if comments.Len() != 1 {
		t.Errorf("comments.Len() = %d, want 1", comments.Len())
	}

	// Verify comment metadata
	gotCA, err := comment.Author()
	if err != nil {
		t.Fatal(err)
	}
	if gotCA != "John Doe" {
		t.Errorf("comment.Author() = %q, want %q", gotCA, "John Doe")
	}

	// Verify comment ID was assigned
	cid, err := comment.CommentID()
	if err != nil {
		t.Fatalf("CommentID: %v", err)
	}

	// Verify XML range markers are present in the paragraph
	pEl := para.p.RawElement()
	var hasStart, hasEnd bool
	for _, child := range pEl.ChildElements() {
		if child.Tag == "commentRangeStart" {
			if child.SelectAttrValue("w:id", "") == fmt.Sprintf("%d", cid) {
				hasStart = true
			}
		}
		if child.Tag == "commentRangeEnd" {
			if child.SelectAttrValue("w:id", "") == fmt.Sprintf("%d", cid) {
				hasEnd = true
			}
		}
	}
	if !hasStart {
		t.Errorf("missing w:commentRangeStart with id=%d", cid)
	}
	if !hasEnd {
		t.Errorf("missing w:commentRangeEnd with id=%d", cid)
	}
}

// Mirrors Python: Document.it_can_add_a_comment — error on empty runs
func TestDocument_AddComment_EmptyRuns(t *testing.T) {
	doc := mustNewDoc(t)

	_, err := doc.AddComment(nil, "text", "author", nil)
	if err == nil {
		t.Error("expected error when runs is nil")
	}

	_, err = doc.AddComment([]*Run{}, "text", "author", nil)
	if err == nil {
		t.Error("expected error when runs is empty")
	}
}

// Mirrors Python: Document.it_can_add_a_picture
//
// Verifies the full flow: AddPicture creates a paragraph/run, inserts
// the image, and the inline shapes count increases.
func TestDocument_AddPicture_Flow(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	shape, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}
	if shape == nil {
		t.Fatal("AddPicture returned nil")
	}

	// Verify a new paragraph was created
	paras := mustParagraphs(t, doc)
	if len(paras) < 1 {
		t.Error("expected at least one paragraph after AddPicture")
	}

	// Verify inline shapes count
	shapes, err := doc.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes: %v", err)
	}
	if shapes.Len() != 1 {
		t.Errorf("InlineShapes.Len() = %d, want 1", shapes.Len())
	}
}

// Mirrors Python: Document.it_can_add_a_picture with explicit dimensions
func TestDocument_AddPicture_WithDimensions(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	w := int64(914400) // 1 inch in EMU
	h := int64(457200) // 0.5 inch in EMU

	shape, err := doc.AddPicture(r, &w, &h)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}
	if shape == nil {
		t.Fatal("AddPicture returned nil")
	}

	// Verify dimensions were applied
	gotW, err := shape.Width()
	if err != nil {
		t.Fatalf("Width: %v", err)
	}
	if gotW != Length(w) {
		t.Errorf("Width = %d, want %d", gotW, w)
	}

	gotH, err := shape.Height()
	if err != nil {
		t.Fatalf("Height: %v", err)
	}
	if gotH != Length(h) {
		t.Errorf("Height = %d, want %d", gotH, h)
	}
}

// Mirrors Python: _Body.it_can_clear_itself_of_all_content_it_holds
//
// Python parametrize cases:
//
//	("w:body", "w:body")                            — empty body unchanged
//	("w:body/w:p", "w:body")                        — paragraph removed
//	("w:body/w:sectPr", "w:body/w:sectPr")          — sectPr preserved
//	("w:body/(w:p, w:sectPr)", "w:body/w:sectPr")   — p removed, sectPr preserved
func TestBody_ClearContent_PreservesSectPr(t *testing.T) {
	tests := []struct {
		name         string
		bodyXml      string
		wantChildren []string // expected child tag names after clear
	}{
		{
			"empty_body",
			`<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			nil,
		},
		{
			"paragraph_removed",
			`<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:body>`,
			nil,
		},
		{
			"sectPr_preserved",
			`<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sectPr/></w:body>`,
			[]string{"sectPr"},
		},
		{
			"p_removed_sectPr_preserved",
			`<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/><w:sectPr/></w:body>`,
			[]string{"sectPr"},
		},
		{
			"tbl_and_p_removed_sectPr_preserved",
			`<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/><w:tbl/><w:sectPr/></w:body>`,
			[]string{"sectPr"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			el := mustParseXml(t, tt.bodyXml)
			body := &oxml.CT_Body{Element: *el}
			body.ClearContent()

			children := body.RawElement().ChildElements()
			var gotTags []string
			for _, c := range children {
				gotTags = append(gotTags, c.Tag)
			}

			if len(gotTags) != len(tt.wantChildren) {
				t.Fatalf("child count = %d (%v), want %d (%v)", len(gotTags), gotTags, len(tt.wantChildren), tt.wantChildren)
			}
			for i := range gotTags {
				if gotTags[i] != tt.wantChildren[i] {
					t.Errorf("child[%d] = %q, want %q", i, gotTags[i], tt.wantChildren[i])
				}
			}
		})
	}
}

// =======================================================================
// Helpers
// =======================================================================

// cellText extracts the text from a table cell at (row, col).
func cellText(t *testing.T, tbl *Table, row, col int) string {
	t.Helper()
	cell, err := tbl.CellAt(row, col)
	if err != nil {
		t.Fatalf("CellAt(%d,%d): %v", row, col, err)
	}
	return cell.Text()
}
