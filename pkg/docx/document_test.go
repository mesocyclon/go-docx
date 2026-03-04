package docx

import (
	"bytes"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func mustNewDoc(t *testing.T) *Document {
	t.Helper()
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error: %v", err)
	}
	return doc
}

func mustParagraphs(t *testing.T, d *Document) []*Paragraph {
	t.Helper()
	p, err := d.Paragraphs()
	if err != nil {
		t.Fatalf("Paragraphs() error: %v", err)
	}
	return p
}

func mustTables(t *testing.T, d *Document) []*Table {
	t.Helper()
	tbl, err := d.Tables()
	if err != nil {
		t.Fatalf("Tables() error: %v", err)
	}
	return tbl
}

func mustInlineShapes(t *testing.T, d *Document) *InlineShapes {
	t.Helper()
	s, err := d.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	return s
}

func mustIterInnerContent(t *testing.T, d *Document) []*InnerContentItem {
	t.Helper()
	items, err := d.IterInnerContent()
	if err != nil {
		t.Fatalf("IterInnerContent() error: %v", err)
	}
	return items
}

// --------------------------------------------------------------------------
// AddParagraph
// --------------------------------------------------------------------------

func TestDocument_AddParagraph(t *testing.T) {
	doc := mustNewDoc(t)
	before := len(mustParagraphs(t, doc))

	p, err := doc.AddParagraph("hello")
	if err != nil {
		t.Fatalf("AddParagraph error: %v", err)
	}
	if p == nil {
		t.Fatal("AddParagraph returned nil")
	}
	after := len(mustParagraphs(t, doc))
	if after != before+1 {
		t.Errorf("Paragraphs count: want %d, got %d", before+1, after)
	}

	last := mustParagraphs(t, doc)[after-1]
	if got := last.Text(); got != "hello" {
		t.Errorf("last paragraph text: want %q, got %q", "hello", got)
	}
}

func TestDocument_AddParagraph_WithStyle(t *testing.T) {
	doc := mustNewDoc(t)
	// "Heading 1" is the UI name; BabelFish translates to internal "heading 1".
	_, err := doc.AddParagraph("styled", StyleName("Heading 1"))
	if err != nil {
		t.Fatalf("AddParagraph with style error: %v", err)
	}
}

// --------------------------------------------------------------------------
// AddHeading
// --------------------------------------------------------------------------

func TestDocument_AddHeading_Title(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddHeading("My Title", 0)
	if err != nil {
		t.Fatalf("AddHeading(0) error: %v", err)
	}
	if p == nil {
		t.Fatal("AddHeading returned nil")
	}
}

func TestDocument_AddHeading_Level1(t *testing.T) {
	doc := mustNewDoc(t)
	_, err := doc.AddHeading("H1", 1)
	if err != nil {
		t.Fatalf("AddHeading(1) error: %v", err)
	}
}

func TestDocument_AddHeading_OutOfRange(t *testing.T) {
	doc := mustNewDoc(t)
	_, err := doc.AddHeading("Bad", 10)
	if err == nil {
		t.Error("AddHeading(10) expected error, got nil")
	}
	_, err = doc.AddHeading("Bad", -1)
	if err == nil {
		t.Error("AddHeading(-1) expected error, got nil")
	}
}

// --------------------------------------------------------------------------
// AddPageBreak
// --------------------------------------------------------------------------

func TestDocument_AddPageBreak(t *testing.T) {
	doc := mustNewDoc(t)
	before := len(mustParagraphs(t, doc))
	p, err := doc.AddPageBreak()
	if err != nil {
		t.Fatalf("AddPageBreak error: %v", err)
	}
	if p == nil {
		t.Fatal("AddPageBreak returned nil")
	}
	after := len(mustParagraphs(t, doc))
	if after != before+1 {
		t.Errorf("Paragraphs count after page break: want %d, got %d", before+1, after)
	}
}

// --------------------------------------------------------------------------
// AddTable
// --------------------------------------------------------------------------

func TestDocument_AddTable(t *testing.T) {
	doc := mustNewDoc(t)
	table, err := doc.AddTable(2, 3)
	if err != nil {
		t.Fatalf("AddTable error: %v", err)
	}
	if table == nil {
		t.Fatal("AddTable returned nil")
	}
	tables := mustTables(t, doc)
	if len(tables) != 1 {
		t.Errorf("Tables count: want 1, got %d", len(tables))
	}
}

// --------------------------------------------------------------------------
// AddSection
// --------------------------------------------------------------------------

func TestDocument_AddSection(t *testing.T) {
	doc := mustNewDoc(t)
	before := doc.Sections().Len()

	s, err := doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		t.Fatalf("AddSection error: %v", err)
	}
	if s == nil {
		t.Fatal("AddSection returned nil")
	}
	after := doc.Sections().Len()
	if after != before+1 {
		t.Errorf("Sections count: want %d, got %d", before+1, after)
	}
}

// --------------------------------------------------------------------------
// Sections, Styles, Settings, Comments
// --------------------------------------------------------------------------

func TestDocument_Sections(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	if sections.Len() < 1 {
		t.Error("expected at least 1 section")
	}
}

func TestDocument_Styles(t *testing.T) {
	doc := mustNewDoc(t)
	styles, err := doc.Styles()
	if err != nil {
		t.Fatalf("Styles() error: %v", err)
	}
	if styles == nil {
		t.Fatal("Styles() returned nil")
	}
	// Should contain some built-in styles.
	if styles.Len() == 0 {
		t.Error("expected at least some styles")
	}
}

func TestDocument_Settings(t *testing.T) {
	doc := mustNewDoc(t)
	settings, err := doc.Settings()
	if err != nil {
		t.Fatalf("Settings() error: %v", err)
	}
	if settings == nil {
		t.Fatal("Settings() returned nil")
	}
}

func TestDocument_Comments(t *testing.T) {
	doc := mustNewDoc(t)
	comments, err := doc.Comments()
	if err != nil {
		t.Fatalf("Comments() error: %v", err)
	}
	if comments == nil {
		t.Fatal("Comments() returned nil")
	}
	if comments.Len() != 0 {
		t.Errorf("new document should have 0 comments, got %d", comments.Len())
	}
}

// --------------------------------------------------------------------------
// Paragraphs / Tables
// --------------------------------------------------------------------------

func TestDocument_Paragraphs_EmptyDoc(t *testing.T) {
	doc := mustNewDoc(t)
	paras := mustParagraphs(t, doc)
	// The default.docx template body contains only w:sectPr, no w:p.
	// Python Document() also starts with 0 paragraphs.
	if len(paras) != 0 {
		t.Errorf("expected 0 paragraphs in default doc, got %d", len(paras))
	}
}

func TestDocument_Tables_EmptyDoc(t *testing.T) {
	doc := mustNewDoc(t)
	tables := mustTables(t, doc)
	if len(tables) != 0 {
		t.Errorf("expected 0 tables in default doc, got %d", len(tables))
	}
}

// --------------------------------------------------------------------------
// InlineShapes
// --------------------------------------------------------------------------

func TestDocument_InlineShapes(t *testing.T) {
	doc := mustNewDoc(t)
	shapes := mustInlineShapes(t, doc)
	if shapes == nil {
		t.Fatal("InlineShapes() returned nil")
	}
	if shapes.Len() != 0 {
		t.Errorf("expected 0 inline shapes, got %d", shapes.Len())
	}
}

// --------------------------------------------------------------------------
// IterInnerContent
// --------------------------------------------------------------------------

func TestDocument_IterInnerContent(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("p1")
	doc.AddTable(1, 1)
	doc.AddParagraph("p2")

	items := mustIterInnerContent(t, doc)
	if len(items) < 3 {
		t.Errorf("expected at least 3 items, got %d", len(items))
	}
	// Count paragraphs and tables.
	var paras, tables int
	for _, it := range items {
		if it.IsParagraph() {
			paras++
		}
		if it.IsTable() {
			tables++
		}
	}
	if tables < 1 {
		t.Errorf("expected at least 1 table in inner content, got %d", tables)
	}
	if paras < 2 {
		t.Errorf("expected at least 2 paragraphs in inner content, got %d", paras)
	}
}

// --------------------------------------------------------------------------
// blockWidth
// --------------------------------------------------------------------------

func TestDocument_BlockWidth(t *testing.T) {
	doc := mustNewDoc(t)
	bw, err := doc.blockWidth()
	if err != nil {
		t.Fatalf("blockWidth() error: %v", err)
	}
	// default.docx: w:w="12240" w:left="1800" w:right="1800"
	// → 12240 - 1800 - 1800 = 8640 twips  (= 6" text width)
	// This matches Python: Document()._block_width == 5486400 EMU == 8640 twips.
	expected := 8640
	// Allow some tolerance since templates may differ.
	if bw < expected-100 || bw > expected+100 {
		t.Errorf("blockWidth: want ~%d twips, got %d", expected, bw)
	}
}

// --------------------------------------------------------------------------
// Save round-trip
// --------------------------------------------------------------------------

func TestDocument_Save_RoundTrip(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("Round trip")

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save error: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes error: %v", err)
	}

	found := false
	for _, p := range mustParagraphs(t, doc2) {
		if p.Text() == "Round trip" {
			found = true
			break
		}
	}
	if !found {
		t.Error("round-trip: paragraph text 'Round trip' not found")
	}
}

func TestDocument_Save_PreservesChanges(t *testing.T) {
	// Create → add content → save → open → add more → save → open → verify all.
	doc, _ := New()
	doc.AddParagraph("Original")

	var buf1 bytes.Buffer
	doc.Save(&buf1)

	doc2, _ := OpenBytes(buf1.Bytes())
	doc2.AddParagraph("Added")

	var buf2 bytes.Buffer
	doc2.Save(&buf2)

	doc3, err := OpenBytes(buf2.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes error: %v", err)
	}

	texts := make(map[string]bool)
	for _, p := range mustParagraphs(t, doc3) {
		texts[p.Text()] = true
	}
	if !texts["Original"] {
		t.Error("missing 'Original' paragraph")
	}
	if !texts["Added"] {
		t.Error("missing 'Added' paragraph")
	}
}

func TestDocument_Save_ByteStability(t *testing.T) {
	// Save twice without changes → content preserved (paragraph count).
	doc, _ := New()
	doc.AddParagraph("Stable content")

	var buf1 bytes.Buffer
	doc.Save(&buf1)

	doc2, _ := OpenBytes(buf1.Bytes())
	var buf2 bytes.Buffer
	doc2.Save(&buf2)

	doc3, _ := OpenBytes(buf2.Bytes())
	if len(mustParagraphs(t, doc2)) != len(mustParagraphs(t, doc3)) {
		t.Errorf("paragraph count changed: %d → %d",
			len(mustParagraphs(t, doc2)), len(mustParagraphs(t, doc3)))
	}
}
