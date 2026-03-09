package docx

import (
	"bytes"
	"strconv"
	"strings"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// ---------------------------------------------------------------------------
// Test helpers for ReplaceWithContent
// ---------------------------------------------------------------------------

// rcSourceWithParagraph creates a source doc with a single paragraph.
func rcSourceWithParagraph(t *testing.T, text string) *Document {
	t.Helper()
	src := mustNewDoc(t)
	src.AddParagraph(text)
	return src
}

// rcSourceWithParagraphAndTable creates a source doc with a paragraph + table.
func rcSourceWithParagraphAndTable(t *testing.T) *Document {
	t.Helper()
	src := mustNewDoc(t)
	src.AddParagraph("Source paragraph")
	src.AddTable(1, 2)
	return src
}

// rcInjectBlip injects a <w:p> with a nested <a:blip r:embed="rIdN"> into the
// source body and wires a relationship to the given ImagePart. Returns rId.
func rcInjectBlip(t *testing.T, src *Document, imgPart *parts.ImagePart) string {
	t.Helper()
	// Wire relationship on source.
	rId, _ := src.Part().StoryPart.GetOrAddImage(imgPart)

	// Inject drawing XML referencing that rId.
	body := src.element.Body().RawElement()
	// Insert before sectPr.
	p := etree.NewElement("w:p")
	r := p.CreateElement("w:r")
	drawing := r.CreateElement("w:drawing")
	inline := drawing.CreateElement("wp:inline")
	graphic := inline.CreateElement("a:graphic")
	graphicData := graphic.CreateElement("a:graphicData")
	blip := graphicData.CreateElement("a:blip")
	blip.CreateAttr("r:embed", rId)

	// Find sectPr to insert before it.
	children := body.ChildElements()
	inserted := false
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			body.InsertChildAt(i, p)
			inserted = true
			break
		}
	}
	if !inserted {
		body.AddChild(p)
	}
	return rId
}

// rcInjectHyperlink injects a <w:p> with <w:hyperlink r:id="rIdN"> into the
// source body and wires an external relationship. Returns rId.
func rcInjectHyperlink(t *testing.T, src *Document, url string) string {
	t.Helper()
	rId := src.Part().Rels().GetOrAddExtRel(opc.RTHyperlink, url)

	body := src.element.Body().RawElement()
	p := etree.NewElement("w:p")
	hl := p.CreateElement("w:hyperlink")
	hl.CreateAttr("r:id", rId)
	r := hl.CreateElement("w:r")
	wt := r.CreateElement("w:t")
	wt.SetText("click here")

	children := body.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			body.InsertChildAt(i, p)
			return rId
		}
	}
	body.AddChild(p)
	return rId
}

// rcBodyTexts collects text from all top-level paragraphs.
func rcBodyTexts(t *testing.T, doc *Document) []string {
	t.Helper()
	paras := mustParagraphs(t, doc)
	var texts []string
	for _, p := range paras {
		texts = append(texts, p.Text())
	}
	return texts
}

// rcBodyItemSummary returns a summary like ["P", "TBL", "P", "TBL"] of body items.
func rcBodyItemSummary(t *testing.T, doc *Document) []string {
	t.Helper()
	items := mustIterInnerContent(t, doc)
	var summary []string
	for _, it := range items {
		if it.IsParagraph() {
			summary = append(summary, "P")
		} else if it.IsTable() {
			summary = append(summary, "TBL")
		}
	}
	return summary
}

// rcCountBodyElements counts paragraphs and tables in body.
func rcCountBodyElements(t *testing.T, doc *Document) (paras, tables int) {
	t.Helper()
	items := mustIterInnerContent(t, doc)
	for _, it := range items {
		if it.IsParagraph() {
			paras++
		} else if it.IsTable() {
			tables++
		}
	}
	return
}

// rcFindAllBlipEmbeds recursively finds all r:embed values in body elements.
func rcFindAllBlipEmbeds(t *testing.T, doc *Document) []string {
	t.Helper()
	body, err := doc.getBody()
	if err != nil {
		t.Fatalf("getBody: %v", err)
	}
	var result []string
	stack := body.Element().ChildElements()
	for len(stack) > 0 {
		el := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		for _, attr := range el.Attr {
			if attr.Space == "r" && attr.Key == "embed" && attr.Value != "" {
				result = append(result, attr.Value)
			}
		}
		stack = append(stack, el.ChildElements()...)
	}
	return result
}

// ---------------------------------------------------------------------------
// Integration tests: Document.ReplaceWithContent
// ---------------------------------------------------------------------------

func TestDocument_ReplaceWithContent_Simple(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<CONTENT>]")

	source := rcSourceWithParagraph(t, "Inserted text")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// The body should contain the source paragraph text.
	found := false
	for _, txt := range rcBodyTexts(t, target) {
		if strings.Contains(txt, "Inserted text") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'Inserted text' in target body")
	}
}

func TestDocument_ReplaceWithContent_MultipleElements(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<CONTENT>]")

	source := rcSourceWithParagraphAndTable(t)

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Body should contain at least 1 paragraph + 1 table from source.
	paras, tables := rcCountBodyElements(t, target)
	if paras < 1 {
		t.Errorf("expected at least 1 paragraph, got %d", paras)
	}
	if tables < 1 {
		t.Errorf("expected at least 1 table, got %d", tables)
	}
}

func TestDocument_ReplaceWithContent_WithSurroundingText(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("Before [<CONTENT>] After")

	source := rcSourceWithParagraph(t, "Middle")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Should have: p("Before "), source content, p(" After")
	texts := rcBodyTexts(t, target)
	foundBefore := false
	foundAfter := false
	foundMiddle := false
	for _, txt := range texts {
		if strings.Contains(txt, "Before") {
			foundBefore = true
		}
		if strings.Contains(txt, "After") {
			foundAfter = true
		}
		if strings.Contains(txt, "Middle") {
			foundMiddle = true
		}
	}
	if !foundBefore {
		t.Error("surrounding text 'Before' not found")
	}
	if !foundAfter {
		t.Error("surrounding text 'After' not found")
	}
	if !foundMiddle {
		t.Error("inserted text 'Middle' not found")
	}
}

func TestDocument_ReplaceWithContent_Image(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<CONTENT>]")

	source := mustNewDoc(t)
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A} // PNG header
	imgPart := parts.NewImagePartWithMeta(
		"/word/media/image1.png", "image/png", imgBlob,
		100, 100, 72, 72, "test.png",
	)
	// Add image part to source package.
	source.wmlPkg.OpcPackage.AddPart(imgPart)
	srcRId := rcInjectBlip(t, source, imgPart)
	_ = srcRId

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Target should have a blip with remapped r:embed.
	embeds := rcFindAllBlipEmbeds(t, target)
	if len(embeds) == 0 {
		t.Fatal("no r:embed found in target after image insertion")
	}

	// Verify the relationship exists in target.
	newRId := embeds[0]
	rel := target.Part().Rels().GetByRID(newRId)
	if rel == nil {
		t.Fatalf("relationship %s not found in target", newRId)
	}
	if rel.RelType != opc.RTImage {
		t.Errorf("expected RTImage, got %s", rel.RelType)
	}
}

func TestDocument_ReplaceWithContent_ImageDedup(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")
	target.AddParagraph("[<C>]")

	source := mustNewDoc(t)
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	imgPart := parts.NewImagePartWithMeta(
		"/word/media/image1.png", "image/png", imgBlob,
		100, 100, 72, 72, "test.png",
	)
	source.wmlPkg.OpcPackage.AddPart(imgPart)
	rcInjectBlip(t, source, imgPart)

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}

	// Both inserted blips should reference the SAME image part (SHA-256 dedup).
	embeds := rcFindAllBlipEmbeds(t, target)
	if len(embeds) < 2 {
		t.Fatalf("expected at least 2 blip embeds, got %d", len(embeds))
	}

	// All embeds should resolve to the same image relationship target.
	rel0 := target.Part().Rels().GetByRID(embeds[0])
	rel1 := target.Part().Rels().GetByRID(embeds[1])
	if rel0 == nil || rel1 == nil {
		t.Fatal("one of the image relationships not found")
	}
	if rel0.TargetPart != rel1.TargetPart {
		t.Error("expected deduped image: both blips should point to same ImagePart")
	}
}

func TestDocument_ReplaceWithContent_ExternalHyperlink(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<CONTENT>]")

	source := mustNewDoc(t)
	source.AddParagraph("text before link")
	rcInjectHyperlink(t, source, "https://example.com")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Target should have an external hyperlink relationship.
	found := false
	for _, rel := range target.Part().Rels().All() {
		if rel.IsExternal && rel.RelType == opc.RTHyperlink && rel.TargetRef == "https://example.com" {
			found = true
			break
		}
	}
	if !found {
		t.Error("expected external hyperlink relationship in target")
	}
}

func TestDocument_ReplaceWithContent_InTableCell(t *testing.T) {
	target := mustNewDoc(t)
	tbl, err := target.AddTable(1, 1)
	if err != nil {
		t.Fatalf("AddTable: %v", err)
	}
	cell, err := tbl.CellAt(0, 0)
	if err != nil {
		t.Fatalf("CellAt: %v", err)
	}
	cell.SetText("[<CONTENT>]")

	source := rcSourceWithParagraph(t, "cell content")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Cell should contain the inserted paragraph.
	cellParas := cell.Paragraphs()
	found := false
	for _, p := range cellParas {
		if strings.Contains(p.Text(), "cell content") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'cell content' in cell paragraph")
	}

	// Cell must have at least one <w:p> (OOXML invariant).
	if len(cellParas) < 1 {
		t.Error("cell must have at least one paragraph (OOXML invariant)")
	}
}

func TestDocument_ReplaceWithContent_InHeader(t *testing.T) {
	target := mustNewDoc(t)
	sec := target.Sections().Iter()[0]
	hdr := sec.Header()
	_, err := hdr.AddParagraph("[<CONTENT>]")
	if err != nil {
		t.Fatalf("AddParagraph to header: %v", err)
	}

	source := rcSourceWithParagraph(t, "header content")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count < 1 {
		t.Errorf("count = %d, want >= 1", count)
	}

	// Header should contain the inserted text.
	hdrParas, err := hdr.Paragraphs()
	if err != nil {
		t.Fatalf("header Paragraphs error: %v", err)
	}
	found := false
	for _, p := range hdrParas {
		if strings.Contains(p.Text(), "header content") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'header content' in header")
	}
}

func TestDocument_ReplaceWithContent_HeaderDedup(t *testing.T) {
	target := mustNewDoc(t)
	sec0 := target.Sections().Iter()[0]
	_, err := sec0.Header().AddParagraph("[<CONTENT>]")
	if err != nil {
		t.Fatalf("AddParagraph to sec0 header: %v", err)
	}

	// sec1: linked to previous (shares sec0's header part).
	_, err = target.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		t.Fatalf("AddSection: %v", err)
	}

	source := rcSourceWithParagraph(t, "dedup content")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	// The tag should be replaced only once (deduplication).
	if count != 1 {
		t.Errorf("total count = %d, want 1 (dedup should prevent double replacement)", count)
	}
}

func TestDocument_ReplaceWithContent_SectionBreak(t *testing.T) {
	target := mustNewDoc(t)
	p, err := target.AddParagraph("Before [<CONTENT>] After")
	if err != nil {
		t.Fatalf("AddParagraph error: %v", err)
	}

	// Add sectPr to this paragraph (simulates section break).
	sectPr := &oxml.CT_SectPr{Element: oxml.WrapElement(oxml.OxmlElement("w:sectPr"))}
	p.CT_P().SetSectPr(sectPr)

	source := rcSourceWithParagraph(t, "inserted")

	_, err = target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}

	// The sectPr should be on the LAST paragraph fragment.
	body, err := target.getBody()
	if err != nil {
		t.Fatalf("getBody error: %v", err)
	}
	var lastP *etree.Element
	for _, child := range body.Element().ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			lastP = child
		}
	}
	if lastP == nil {
		t.Fatal("no paragraph found in body after replacement")
	}

	hasSectPr := false
	for _, child := range lastP.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			for _, sub := range child.ChildElements() {
				if sub.Space == "w" && sub.Tag == "sectPr" {
					hasSectPr = true
				}
			}
		}
	}
	if !hasSectPr {
		t.Error("sectPr should be on the last paragraph fragment, but was not found")
	}
}

func TestDocument_ReplaceWithContent_MultipleOccurrences(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")
	target.AddParagraph("[<C>]")

	source := rcSourceWithParagraph(t, "repeated")

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}

	// Should find the inserted text twice.
	n := 0
	for _, txt := range rcBodyTexts(t, target) {
		if strings.Contains(txt, "repeated") {
			n++
		}
	}
	if n < 2 {
		t.Errorf("expected 'repeated' at least 2 times, found %d", n)
	}
}

func TestDocument_ReplaceWithContent_MultipleInOneParagraph(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("A[<C>]B[<C>]D")

	source := rcSourceWithParagraph(t, "X")

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}

	// Should produce: p("A"), source, p("B"), source, p("D") — at least 3 paras.
	texts := rcBodyTexts(t, target)
	foundA := false
	foundB := false
	foundD := false
	xCount := 0
	for _, txt := range texts {
		if strings.Contains(txt, "A") && !strings.Contains(txt, "After") {
			foundA = true
		}
		if txt == "B" {
			foundB = true
		}
		if txt == "D" {
			foundD = true
		}
		if strings.Contains(txt, "X") {
			xCount++
		}
	}
	if !foundA {
		t.Error("text fragment 'A' not found")
	}
	if !foundB {
		t.Error("text fragment 'B' not found")
	}
	if !foundD {
		t.Error("text fragment 'D' not found")
	}
	if xCount < 2 {
		t.Errorf("expected source text 'X' at least 2 times, got %d", xCount)
	}
}

func TestDocument_ReplaceWithContent_EmptySource(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("Before [<C>] After")

	// Source with empty body (only sectPr).
	source := mustNewDoc(t)
	body := source.element.Body().RawElement()
	var toRemove []*etree.Element
	for _, child := range body.ChildElements() {
		if !(child.Space == "w" && child.Tag == "sectPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		body.RemoveChild(child)
	}

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	// Placeholder is removed but surrounding text preserved.
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	allText := ""
	for _, txt := range rcBodyTexts(t, target) {
		allText += txt
	}
	if !strings.Contains(allText, "Before") {
		t.Error("'Before' text lost")
	}
	if !strings.Contains(allText, "After") {
		t.Error("'After' text lost")
	}
	if strings.Contains(allText, "[<C>]") {
		t.Error("placeholder was not removed")
	}
}

func TestDocument_ReplaceWithContent_NoMatch(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("no tags here")

	source := rcSourceWithParagraph(t, "unused")

	count, err := target.ReplaceWithContent("[<MISSING>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("count = %d, want 0", count)
	}

	// Document should be unchanged.
	texts := rcBodyTexts(t, target)
	if len(texts) != 1 || texts[0] != "no tags here" {
		t.Errorf("document was modified unexpectedly: %v", texts)
	}
}

func TestDocument_ReplaceWithContent_EmptyOld(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")

	source := rcSourceWithParagraph(t, "unused")

	count, err := target.ReplaceWithContent("", ContentData{Source: source})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("count = %d, want 0 for empty old", count)
	}
}

func TestDocument_ReplaceWithContent_NilSource(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")

	_, err := target.ReplaceWithContent("[<C>]", ContentData{Source: nil})
	if err == nil {
		t.Fatal("expected error for nil Source")
	}
	if !strings.Contains(err.Error(), "nil") {
		t.Errorf("unexpected error message: %v", err)
	}
}

func TestDocument_ReplaceWithContent_RoundTrip(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("Before [<C>] After")

	source := rcSourceWithParagraphAndTable(t)

	_, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}

	// Save → reopen.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save error: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes error: %v", err)
	}

	// Verify structure survived.
	paras2 := mustParagraphs(t, doc2)
	tables2 := mustTables(t, doc2)

	if len(tables2) < 1 {
		t.Error("round-trip: expected at least 1 table")
	}

	foundBefore := false
	foundAfter := false
	foundSource := false
	for _, p := range paras2 {
		txt := p.Text()
		if strings.Contains(txt, "Before") {
			foundBefore = true
		}
		if strings.Contains(txt, "After") {
			foundAfter = true
		}
		if strings.Contains(txt, "Source paragraph") {
			foundSource = true
		}
	}
	if !foundBefore {
		t.Error("round-trip: 'Before' text lost")
	}
	if !foundAfter {
		t.Error("round-trip: 'After' text lost")
	}
	if !foundSource {
		t.Error("round-trip: 'Source paragraph' text lost")
	}
}

func TestDocument_ReplaceWithContent_SourceNotModified(t *testing.T) {
	source := rcSourceWithParagraph(t, "original text")

	// Capture source state.
	srcBody := source.element.Body().RawElement()
	var origTexts []string
	for _, child := range srcBody.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			for _, r := range child.ChildElements() {
				if r.Tag == "r" {
					for _, tc := range r.ChildElements() {
						if tc.Tag == "t" {
							origTexts = append(origTexts, tc.Text())
						}
					}
				}
			}
		}
	}

	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")

	_, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}

	// Verify source is unchanged.
	var afterTexts []string
	for _, child := range srcBody.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			for _, r := range child.ChildElements() {
				if r.Tag == "r" {
					for _, tc := range r.ChildElements() {
						if tc.Tag == "t" {
							afterTexts = append(afterTexts, tc.Text())
						}
					}
				}
			}
		}
	}

	if len(origTexts) != len(afterTexts) {
		t.Fatalf("source text count changed: %d → %d", len(origTexts), len(afterTexts))
	}
	for i, txt := range origTexts {
		if txt != afterTexts[i] {
			t.Errorf("source text[%d] changed: %q → %q", i, txt, afterTexts[i])
		}
	}
}

func TestDocument_ReplaceWithContent_StylePreservation(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")

	// Source with a styled paragraph.
	source := mustNewDoc(t)
	source.AddParagraph("Styled heading", StyleName("Heading 1"))

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Find the inserted paragraph and check it has a pStyle element.
	body, err := target.getBody()
	if err != nil {
		t.Fatalf("getBody: %v", err)
	}
	foundStyle := false
	for _, child := range body.Element().ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			for _, ppr := range child.ChildElements() {
				if ppr.Space == "w" && ppr.Tag == "pPr" {
					for _, ps := range ppr.ChildElements() {
						if ps.Space == "w" && ps.Tag == "pStyle" {
							foundStyle = true
						}
					}
				}
			}
		}
	}
	if !foundStyle {
		t.Error("expected pStyle from source to be preserved in target")
	}
}

func TestDocument_ReplaceWithContent_SelfReference(t *testing.T) {
	// Source == target: should work without panic.
	doc := mustNewDoc(t)
	doc.AddParagraph("intro")
	doc.AddParagraph("[<SELF>]")
	doc.AddParagraph("outro")

	count, err := doc.ReplaceWithContent("[<SELF>]", ContentData{Source: doc})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count < 1 {
		t.Errorf("count = %d, want >= 1", count)
	}

	// "intro" should appear at least twice (original + inserted copy).
	n := 0
	for _, txt := range rcBodyTexts(t, doc) {
		if strings.Contains(txt, "intro") {
			n++
		}
	}
	if n < 2 {
		t.Errorf("expected 'intro' at least 2 times after self-reference, got %d", n)
	}
}

func TestDocument_ReplaceWithContent_CrossRunTag(t *testing.T) {
	target := mustNewDoc(t)

	// Build a paragraph where the tag is split across two runs:
	// <w:p><w:r><w:t>Text [<CO</w:t></w:r><w:r><w:t>NTENT>] rest</w:t></w:r></w:p>
	body, err := target.getBody()
	if err != nil {
		t.Fatalf("getBody: %v", err)
	}
	pEl := etree.NewElement("w:p")
	r1 := pEl.CreateElement("w:r")
	t1 := r1.CreateElement("w:t")
	t1.SetText("Text [<CO")
	r2 := pEl.CreateElement("w:r")
	t2 := r2.CreateElement("w:t")
	t2.SetText("NTENT>] rest")
	body.insertBeforeSectPr(pEl)

	source := rcSourceWithParagraph(t, "cross-run inserted")

	count, err := target.ReplaceWithContent("[<CONTENT>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	texts := rcBodyTexts(t, target)
	foundInserted := false
	foundText := false
	foundRest := false
	for _, txt := range texts {
		if strings.Contains(txt, "cross-run inserted") {
			foundInserted = true
		}
		if strings.Contains(txt, "Text") {
			foundText = true
		}
		if strings.Contains(txt, "rest") {
			foundRest = true
		}
	}
	if !foundInserted {
		t.Error("inserted content not found")
	}
	if !foundText {
		t.Error("text before tag not found")
	}
	if !foundRest {
		t.Error("text after tag not found")
	}
}

func TestDocument_ReplaceWithContent_ImageRoundTrip(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("[<C>]")

	source := mustNewDoc(t)
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	imgPart := parts.NewImagePartWithMeta(
		"/word/media/image1.png", "image/png", imgBlob,
		100, 100, 72, 72, "test.png",
	)
	source.wmlPkg.OpcPackage.AddPart(imgPart)
	rcInjectBlip(t, source, imgPart)

	_, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}

	// Save → reopen.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save error: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes error: %v", err)
	}

	// After round-trip, should still have image relationship.
	embeds := rcFindAllBlipEmbeds(t, doc2)
	if len(embeds) == 0 {
		t.Error("round-trip: no r:embed found after save/open")
	}
	for _, rId := range embeds {
		rel := doc2.Part().Rels().GetByRID(rId)
		if rel == nil {
			t.Errorf("round-trip: relationship %s not found", rId)
		} else if rel.RelType != opc.RTImage {
			t.Errorf("round-trip: relationship %s type = %s, want RTImage", rId, rel.RelType)
		}
	}
}

func TestDocument_ReplaceWithContent_InComment(t *testing.T) {
	target := mustNewDoc(t)
	comments, err := target.Comments()
	if err != nil {
		t.Fatalf("Comments: %v", err)
	}
	c, err := comments.AddComment("[<C>]", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	_ = c

	source := rcSourceWithParagraph(t, "comment content")

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count < 1 {
		t.Errorf("count = %d, want >= 1", count)
	}
}

func TestDocument_ReplaceWithContent_HeaderImageRelationship(t *testing.T) {
	// Verify that images in headers get wired to the header's StoryPart,
	// not the body's StoryPart.
	target := mustNewDoc(t)
	sec := target.Sections().Iter()[0]
	_, err := sec.Header().AddParagraph("[<C>]")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}

	source := mustNewDoc(t)
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	imgPart := parts.NewImagePartWithMeta(
		"/word/media/image1.png", "image/png", imgBlob,
		100, 100, 72, 72, "test.png",
	)
	source.wmlPkg.OpcPackage.AddPart(imgPart)
	rcInjectBlip(t, source, imgPart)

	count, err := target.ReplaceWithContent("[<C>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count < 1 {
		t.Errorf("count = %d, want >= 1", count)
	}

	// The image relationship should be on the HEADER's rels, not body's.
	hdr := sec.Header()
	bic, err := hdr.blockItemContainer()
	if err != nil {
		t.Fatalf("blockItemContainer: %v", err)
	}
	hdrRels := bic.part.Rels()
	foundInHdr := false
	for _, rel := range hdrRels.All() {
		if rel.RelType == opc.RTImage {
			foundInHdr = true
			break
		}
	}
	if !foundInHdr {
		t.Error("image relationship should be on header StoryPart, not found")
	}
}

func TestDocument_ReplaceWithContent_MultipleSourceTypes(t *testing.T) {
	// Test replacing two different placeholders with two different sources.
	target := mustNewDoc(t)
	target.AddParagraph("[<A>]")
	target.AddParagraph("[<B>]")

	sourceA := rcSourceWithParagraph(t, "Source A content")
	sourceB := rcSourceWithParagraphAndTable(t)

	countA, err := target.ReplaceWithContent("[<A>]", ContentData{Source: sourceA})
	if err != nil {
		t.Fatalf("ReplaceWithContent A error: %v", err)
	}
	countB, err := target.ReplaceWithContent("[<B>]", ContentData{Source: sourceB})
	if err != nil {
		t.Fatalf("ReplaceWithContent B error: %v", err)
	}

	if countA != 1 {
		t.Errorf("countA = %d, want 1", countA)
	}
	if countB != 1 {
		t.Errorf("countB = %d, want 1", countB)
	}

	texts := rcBodyTexts(t, target)
	foundA := false
	foundB := false
	for _, txt := range texts {
		if strings.Contains(txt, "Source A content") {
			foundA = true
		}
		if strings.Contains(txt, "Source paragraph") {
			foundB = true
		}
	}
	if !foundA {
		t.Error("source A content not found")
	}
	if !foundB {
		t.Error("source B content not found")
	}

	// Source B also had a table.
	tables := mustTables(t, target)
	if len(tables) < 1 {
		t.Error("expected at least 1 table from source B")
	}
}

// ---------------------------------------------------------------------------
// Body snapshot / rollback tests (Step 1 — protection against corruption)
// ---------------------------------------------------------------------------

func TestDocument_SnapshotBody_RestoresOnRestore(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("original paragraph")

	// Take snapshot.
	snap, err := doc.snapshotBody()
	if err != nil {
		t.Fatalf("snapshotBody: %v", err)
	}

	// Mutate body.
	doc.AddParagraph("added after snapshot")

	// Before restore: should see both paragraphs.
	texts := rcBodyTexts(t, doc)
	foundAdded := false
	for _, txt := range texts {
		if strings.Contains(txt, "added after snapshot") {
			foundAdded = true
		}
	}
	if !foundAdded {
		t.Fatal("expected 'added after snapshot' before restore")
	}

	// Restore.
	doc.restoreBody(snap)

	// After restore: should only see original paragraph.
	texts = rcBodyTexts(t, doc)
	for _, txt := range texts {
		if strings.Contains(txt, "added after snapshot") {
			t.Error("'added after snapshot' should not exist after restore")
		}
	}
	foundOrig := false
	for _, txt := range texts {
		if strings.Contains(txt, "original paragraph") {
			foundOrig = true
		}
	}
	if !foundOrig {
		t.Error("expected 'original paragraph' after restore")
	}
}

func TestDocument_SnapshotBody_IndependentCopy(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("immutable text")

	snap, err := doc.snapshotBody()
	if err != nil {
		t.Fatalf("snapshotBody: %v", err)
	}

	// Mutate body by clearing it.
	body, err := doc.getBody()
	if err != nil {
		t.Fatalf("getBody: %v", err)
	}
	body.ClearContent()

	// Snapshot should still contain the original text.
	// Verify by restoring and checking.
	doc.restoreBody(snap)
	texts := rcBodyTexts(t, doc)
	found := false
	for _, txt := range texts {
		if strings.Contains(txt, "immutable text") {
			found = true
		}
	}
	if !found {
		t.Error("snapshot was mutated — expected 'immutable text' after restore")
	}
}

func TestDocument_ReplaceWithContent_BodyPreservedOnSuccess(t *testing.T) {
	target := mustNewDoc(t)
	target.AddParagraph("keep this")
	target.AddParagraph("[<TAG>]")

	source := rcSourceWithParagraph(t, "inserted content")

	count, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Body should be modified (not rolled back).
	afterTexts := rcBodyTexts(t, target)

	// "keep this" should still be there.
	foundKeep := false
	for _, txt := range afterTexts {
		if strings.Contains(txt, "keep this") {
			foundKeep = true
		}
	}
	if !foundKeep {
		t.Error("expected 'keep this' to remain after successful replacement")
	}

	// "inserted content" should be there.
	foundInserted := false
	for _, txt := range afterTexts {
		if strings.Contains(txt, "inserted content") {
			foundInserted = true
		}
	}
	if !foundInserted {
		t.Error("expected 'inserted content' after successful replacement")
	}

	// "[<TAG>]" should be gone.
	for _, txt := range afterTexts {
		if strings.Contains(txt, "[<TAG>]") {
			t.Error("tag should have been replaced")
		}
	}
}

func TestDocument_ReplaceWithContent_BodyUntouchedOnEarlyError(t *testing.T) {
	// NilSource triggers an early return BEFORE snapshotBody is called
	// (line 622 returns before line 626). The defer does not fire because
	// snap is never created. This verifies the pre-snapshot guard path:
	// body must remain intact.
	target := mustNewDoc(t)
	target.AddParagraph("before")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: nil})
	if err == nil {
		t.Fatal("expected error for nil source")
	}

	// Body should be untouched — no snapshot, no mutation, no restore.
	texts := rcBodyTexts(t, target)
	found := false
	for _, txt := range texts {
		if strings.Contains(txt, "before") {
			found = true
		}
	}
	if !found {
		t.Error("body should be untouched after nil source error")
	}
}

func TestDocument_ReplaceWithContent_BodyRolledBackOnPostMutationError(t *testing.T) {
	// This test verifies the defer-triggered rollback: body is mutated
	// by successful replacement, but a later phase (comments) fails,
	// and the defer restores the body to its pre-call state.
	//
	// Setup: wire a broken CommentsPart (nil root element) into the
	// target. HasCommentsPart() returns true (relationship exists with
	// non-nil TargetPart), but CommentsElement() returns error because
	// the XmlPart has no root element.
	//
	// Pipeline execution:
	//   1. Phase 1 (import resources)     — succeeds (simple source)
	//   2. Body prep                       — succeeds
	//   3. Body replacement                — succeeds, body MUTATED
	//   4. Headers/footers                 — succeeds (no tags in headers)
	//   5. Comments                        — FAILS (nil element)
	//   6. defer fires                     — body RESTORED
	target := mustNewDoc(t)
	target.AddParagraph("original text")
	target.AddParagraph("[<TAG>]")

	// Wire a broken CommentsPart: XmlPart with no root element.
	// NewXmlPart parses the XML proc-inst but there's no root element,
	// so Element() returns nil and CommentsElement() returns error.
	brokenXP, err := opc.NewXmlPart(
		"/word/comments.xml", opc.CTWmlComments,
		[]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`),
		target.wmlPkg.OpcPackage,
	)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}
	if brokenXP.Element() != nil {
		t.Fatal("precondition failed: expected nil root element")
	}
	brokenCP := parts.NewCommentsPart(brokenXP)
	target.Part().Rels().GetOrAdd(opc.RTComments, brokenCP)

	source := rcSourceWithParagraph(t, "inserted text")

	_, err = target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err == nil {
		t.Fatal("expected error from broken comments part")
	}

	// Body should be rolled back to its pre-call state.
	texts := rcBodyTexts(t, target)
	foundOriginal := false
	foundTag := false
	for _, txt := range texts {
		if strings.Contains(txt, "original text") {
			foundOriginal = true
		}
		if strings.Contains(txt, "[<TAG>]") {
			foundTag = true
		}
		if strings.Contains(txt, "inserted text") {
			t.Error("body should have been rolled back — 'inserted text' must not be present")
		}
	}
	if !foundOriginal {
		t.Error("expected 'original text' to survive rollback")
	}
	if !foundTag {
		t.Error("expected '[<TAG>]' to be restored after rollback")
	}

	// Document must remain saveable after rollback.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save after rollback: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Integration tests: KeepSourceFormatting — expand to direct attributes
// ---------------------------------------------------------------------------

func TestDocument_ReplaceWithContent_KeepSourceFormatting_ExpandsDirect(t *testing.T) {
	// Setup: target and source both have "CustomTitle" style but with
	// DIFFERENT definitions. Under KeepSourceFormatting, the source
	// style formatting should be expanded into direct attributes on
	// the inserted paragraph.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	// Add "CustomTitle" style to target (left-aligned, sz=24).
	tgtStyles, err := target.part.Styles()
	if err != nil {
		t.Fatalf("target Styles: %v", err)
	}
	tgtStyleXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="left"/></w:pPr>` +
		`<w:rPr><w:sz w:val="24"/></w:rPr>` +
		`</w:style>`
	tgtStyleEl, _ := oxml.ParseXml([]byte(tgtStyleXml))
	tgtStyles.RawElement().AddChild(tgtStyleEl)

	// Source: create doc, add "CustomTitle" style (center, bold, sz=28),
	// and inject a paragraph referencing it via raw XML.
	source := mustNewDoc(t)
	srcStyles, err := source.part.Styles()
	if err != nil {
		t.Fatalf("source Styles: %v", err)
	}
	srcStyleXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>` +
		`</w:style>`
	srcStyleEl, _ := oxml.ParseXml([]byte(srcStyleXml))
	srcStyles.RawElement().AddChild(srcStyleEl)

	// Inject paragraph with pStyle=CustomTitle into source body.
	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>Styled heading</w:t></w:r></w:p>`,
	))
	// Insert before sectPr.
	children := srcBody.ChildElements()
	inserted := false
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			inserted = true
			break
		}
	}
	if !inserted {
		srcBody.AddChild(styledP)
	}

	// Replace with KeepSourceFormatting.
	count, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Find the inserted paragraph and verify expansion happened.
	body, err := target.getBody()
	if err != nil {
		t.Fatalf("getBody: %v", err)
	}
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		// Look for expanded jc=center from source style.
		jc := findChild(pPr, "w", "jc")
		if jc != nil && jc.SelectAttrValue("w:val", "") == "center" {
			// The source style's jc=center was expanded as a direct attribute.
			// Also check that rPr with b was expanded.
			rPr := findChild(pPr, "w", "rPr")
			if rPr == nil {
				t.Error("expected rPr with expanded properties")
				return
			}
			if findChild(rPr, "w", "b") == nil {
				t.Error("expected bold (b) expanded from source style")
			}
			// pStyle should be remapped to default (Normal), not CustomTitle.
			pStyle := findChild(pPr, "w", "pStyle")
			if pStyle != nil {
				val := pStyle.SelectAttrValue("w:val", "")
				if val == "CustomTitle" {
					t.Error("pStyle should have been remapped away from CustomTitle")
				}
			}
			return
		}
	}
	t.Error("expected to find paragraph with expanded jc=center from source style")
}

func TestDocument_ReplaceWithContent_KeepSourceFormatting_RoundTrip(t *testing.T) {
	// End-to-end test: expand + save + reopen must produce valid document.
	target := mustNewDoc(t)
	target.AddParagraph("[<X>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:b/></w:rPr></w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	// Inject styled paragraph.
	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>heading text</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	_, err := target.ReplaceWithContent("[<X>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Save and reopen.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	reopened, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Verify text survived.
	texts := rcBodyTexts(t, reopened)
	found := false
	for _, txt := range texts {
		if strings.Contains(txt, "heading text") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'heading text' in reopened document")
	}
}

func TestDocument_ReplaceWithContent_UseDestination_NoExpansion(t *testing.T) {
	// Verify that UseDestinationStyles (default) does NOT expand.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	// Inject styled paragraph.
	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>title</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: UseDestinationStyles, // explicit default
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Under UseDestinationStyles, pStyle stays as CustomTitle (no expansion),
	// and no jc=center should appear (target's left-aligned definition wins).
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		pStyle := findChild(pPr, "w", "pStyle")
		if pStyle == nil {
			continue
		}
		val := pStyle.SelectAttrValue("w:val", "")
		if val == "CustomTitle" {
			// No expansion — jc should NOT be present as direct attr.
			jc := findChild(pPr, "w", "jc")
			if jc != nil && jc.SelectAttrValue("w:val", "") == "center" {
				t.Error("UseDestinationStyles should NOT expand jc=center from source")
			}
			return
		}
	}
}

func TestDocument_SnapshotBody_RoundTrip(t *testing.T) {
	// Verify that snapshot/restore produces a valid document that can be saved.
	doc := mustNewDoc(t)
	doc.AddParagraph("paragraph one")
	doc.AddParagraph("paragraph two")

	snap, err := doc.snapshotBody()
	if err != nil {
		t.Fatalf("snapshotBody: %v", err)
	}

	// Mutate.
	doc.AddParagraph("paragraph three")

	// Restore.
	doc.restoreBody(snap)

	// Save and reopen — document must be valid.
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save after restore: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes after restore: %v", err)
	}

	texts := rcBodyTexts(t, doc2)
	foundOne := false
	foundTwo := false
	for _, txt := range texts {
		if strings.Contains(txt, "paragraph one") {
			foundOne = true
		}
		if strings.Contains(txt, "paragraph two") {
			foundTwo = true
		}
		if strings.Contains(txt, "paragraph three") {
			t.Error("'paragraph three' should not exist after restore")
		}
	}
	if !foundOne {
		t.Error("expected 'paragraph one' in round-tripped document")
	}
	if !foundTwo {
		t.Error("expected 'paragraph two' in round-tripped document")
	}
}

func TestDocument_SnapshotBody_PreservesTablesAndSectPr(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("text before table")
	doc.AddTable(2, 3)
	doc.AddParagraph("text after table")

	snap, err := doc.snapshotBody()
	if err != nil {
		t.Fatalf("snapshotBody: %v", err)
	}

	// Clear everything.
	body, err := doc.getBody()
	if err != nil {
		t.Fatalf("getBody: %v", err)
	}
	body.ClearContent()

	// Restore.
	doc.restoreBody(snap)

	// Check tables preserved.
	paras, tables := rcCountBodyElements(t, doc)
	if tables < 1 {
		t.Errorf("expected at least 1 table after restore, got %d", tables)
	}
	if paras < 2 {
		t.Errorf("expected at least 2 paragraphs after restore, got %d", paras)
	}

	// Verify document is still saveable (sectPr preserved).
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save after restore: %v", err)
	}
}

// --------------------------------------------------------------------------
// ForceCopyStyles integration tests (Step 5)
// --------------------------------------------------------------------------

func TestDocument_ReplaceWithContent_ForceCopyStyles_SuffixGenerated(t *testing.T) {
	t.Parallel()
	// When ForceCopyStyles is true and source/target have conflicting style,
	// the source style must be copied with _0 suffix and all pStyle refs
	// remapped to the new ID.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:b/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	// Inject styled paragraph into source.
	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>Styled heading</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	count, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source:  source,
		Format:  KeepSourceFormatting,
		Options: ImportFormatOptions{ForceCopyStyles: true},
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// 1. Verify CustomTitle_0 exists in target styles.
	tgtStyles, _ = target.part.Styles()
	copied := tgtStyles.GetByID("CustomTitle_0")
	if copied == nil {
		t.Fatal("expected CustomTitle_0 style in target, not found")
	}

	// 2. Verify semiHidden and unhideWhenUsed are set.
	raw := copied.RawElement()
	if findChild(raw, "w", "semiHidden") == nil {
		t.Error("expected semiHidden on CustomTitle_0")
	}
	if findChild(raw, "w", "unhideWhenUsed") == nil {
		t.Error("expected unhideWhenUsed on CustomTitle_0")
	}

	// 3. Verify display name has " (imported)" suffix.
	nameEl := findChild(raw, "w", "name")
	if nameEl == nil {
		t.Fatal("name element not found")
	}
	if got := nameEl.SelectAttrValue("w:val", ""); got != "Custom Title (imported)" {
		t.Errorf("display name = %q, want %q", got, "Custom Title (imported)")
	}

	// 4. Verify the inserted paragraph's pStyle was remapped to CustomTitle_0.
	body, _ := target.getBody()
	found := false
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		pStyle := findChild(pPr, "w", "pStyle")
		if pStyle == nil {
			continue
		}
		if pStyle.SelectAttrValue("w:val", "") == "CustomTitle_0" {
			found = true
			break
		}
	}
	if !found {
		t.Error("expected pStyle remapped to CustomTitle_0 in inserted paragraph")
	}

	// 5. Verify original target CustomTitle is preserved unchanged.
	original := tgtStyles.GetByID("CustomTitle")
	if original == nil {
		t.Fatal("original CustomTitle should still exist in target")
	}
	origPPr := findChild(original.RawElement(), "w", "pPr")
	if origPPr != nil {
		jc := findChild(origPPr, "w", "jc")
		if jc == nil || jc.SelectAttrValue("w:val", "") != "left" {
			t.Error("original CustomTitle should still have jc=left")
		}
	}
}

func TestDocument_ReplaceWithContent_ForceCopyStyles_RoundTrip(t *testing.T) {
	t.Parallel()
	// ForceCopyStyles + save + reopen must produce valid document.
	target := mustNewDoc(t)
	target.AddParagraph("[<X>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:b/></w:rPr></w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>force copy text</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	_, err := target.ReplaceWithContent("[<X>]", ContentData{
		Source:  source,
		Format:  KeepSourceFormatting,
		Options: ImportFormatOptions{ForceCopyStyles: true},
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Save and reopen.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	reopened, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Verify text survived.
	texts := rcBodyTexts(t, reopened)
	found := false
	for _, txt := range texts {
		if strings.Contains(txt, "force copy text") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'force copy text' in reopened document")
	}

	// Verify renamed style survived round-trip.
	reopenedStyles, _ := reopened.part.Styles()
	if reopenedStyles.GetByID("CustomTitle_0") == nil {
		t.Error("expected CustomTitle_0 style to survive round-trip")
	}
}

func TestDocument_ReplaceWithContent_ForceCopyStyles_NoConflict_NoCopy(t *testing.T) {
	t.Parallel()
	// When ForceCopyStyles is set but the source style doesn't conflict
	// (not present in target), it should be copied under its original ID
	// WITHOUT semiHidden/unhideWhenUsed.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="UniqueStyle">` +
		`<w:name w:val="Unique Style"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="UniqueStyle"/></w:pPr>` +
			`<w:r><w:t>unique text</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source:  source,
		Format:  KeepSourceFormatting,
		Options: ImportFormatOptions{ForceCopyStyles: true},
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// UniqueStyle should be copied under original ID (no conflict → no rename).
	tgtStyles, _ := target.part.Styles()
	copied := tgtStyles.GetByID("UniqueStyle")
	if copied == nil {
		t.Fatal("expected UniqueStyle in target")
	}

	// No semiHidden because it was NOT renamed.
	raw := copied.RawElement()
	if findChild(raw, "w", "semiHidden") != nil {
		t.Error("semiHidden should NOT be present for non-conflicting style")
	}
}

func TestDocument_ReplaceWithContent_KeepDifferentStyles_ForceCopy(t *testing.T) {
	t.Parallel()
	// KeepDifferentStyles + ForceCopyStyles: different formatting → rename.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>different</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source:  source,
		Format:  KeepDifferentStyles,
		Options: ImportFormatOptions{ForceCopyStyles: true},
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Styles differ → should be renamed to CustomTitle_0.
	tgtStyles, _ = target.part.Styles()
	copied := tgtStyles.GetByID("CustomTitle_0")
	if copied == nil {
		t.Fatal("expected CustomTitle_0 for KeepDifferentStyles + ForceCopyStyles with different formatting")
	}
	if findChild(copied.RawElement(), "w", "semiHidden") == nil {
		t.Error("expected semiHidden on renamed style")
	}
}

func TestDocument_ReplaceWithContent_KeepDifferentStyles_SameFormatting_NoRename(t *testing.T) {
	t.Parallel()
	// KeepDifferentStyles + ForceCopyStyles: identical formatting → use target, no copy.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	// Both styles have identical formatting (jc=center).
	sharedXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CustomTitle">` +
		`<w:name w:val="Custom Title"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(sharedXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcEl, _ := oxml.ParseXml([]byte(sharedXml))
	srcStyles.RawElement().AddChild(srcEl)

	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="CustomTitle"/></w:pPr>` +
			`<w:r><w:t>same style</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, styledP)
			break
		}
	}

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source:  source,
		Format:  KeepDifferentStyles,
		Options: ImportFormatOptions{ForceCopyStyles: true},
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Identical formatting → no rename, pStyle stays CustomTitle.
	tgtStyles, _ = target.part.Styles()
	if tgtStyles.GetByID("CustomTitle_0") != nil {
		t.Error("expected NO rename when formatting is identical (KeepDifferentStyles)")
	}

	// Verify pStyle stayed as CustomTitle in the inserted paragraph.
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		pStyle := findChild(pPr, "w", "pStyle")
		if pStyle != nil && pStyle.SelectAttrValue("w:val", "") == "CustomTitle" {
			// Good — style was kept as-is.
			return
		}
	}
	t.Error("expected pStyle=CustomTitle in inserted paragraph (same formatting → use target)")
}

// --------------------------------------------------------------------------
// ImportFormatMode — additional coverage (Step 8.2)
// --------------------------------------------------------------------------

func TestRWC_UseDestination_MissingStyleCopied(t *testing.T) {
	t.Parallel()
	// When the source style does NOT exist in target, UseDestinationStyles
	// deep-copies it (all 3 modes agree on this behavior).
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="OnlyInSource">` +
		`<w:name w:val="Only In Source"/>` +
		`<w:pPr><w:jc w:val="right"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "OnlyInSource", "styled text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: UseDestinationStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Style should have been deep-copied to target.
	tgtStyles, _ := target.part.Styles()
	if tgtStyles.GetByID("OnlyInSource") == nil {
		t.Error("expected OnlyInSource style to be deep-copied into target")
	}
}

func TestRWC_KeepSource_Conflict_DirectPreservesExisting(t *testing.T) {
	t.Parallel()
	// Existing direct attributes on a paragraph must NOT be overwritten
	// when expanding source style formatting to direct attributes.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="Conflict">` +
		`<w:name w:val="Conflict"/>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="Conflict">` +
		`<w:name w:val="Conflict"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:sz w:val="28"/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	// Inject paragraph with existing direct jc=right (should win over style's center).
	srcBody := source.element.Body().RawElement()
	styledP, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="Conflict"/><w:jc w:val="right"/></w:pPr>` +
			`<w:r><w:t>direct wins</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, styledP)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Find inserted paragraph — jc must be "right" (direct), not "center" (style).
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		jc := findChild(pPr, "w", "jc")
		if jc == nil {
			continue
		}
		if jc.SelectAttrValue("w:val", "") == "right" {
			// Good — sz=28 should also have been expanded from style.
			rPr := findChild(pPr, "w", "rPr")
			if rPr == nil {
				t.Error("expected rPr with sz=28 expanded from style")
			} else if sz := findChild(rPr, "w", "sz"); sz == nil || sz.SelectAttrValue("w:val", "") != "28" {
				t.Error("expected sz=28 expanded from source style")
			}
			return
		}
	}
	t.Error("expected paragraph with jc=right (direct attribute preserved)")
}

func TestRWC_KeepSource_ForceRename_ChainRenamed(t *testing.T) {
	t.Parallel()
	// When ForceCopyStyles is true and a source style with basedOn chain
	// conflicts, both the child and parent styles should be renamed.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	// Add both "BaseStyle" and "DerivedStyle" to target to cause conflict.
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="BaseStyle">` +
			`<w:name w:val="Base Style"/>` +
			`<w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="DerivedStyle">` +
			`<w:name w:val="Derived Style"/>` +
			`<w:basedOn w:val="BaseStyle"/>` +
			`<w:rPr><w:i/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		tgtStyles.RawElement().AddChild(el)
	}

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	// Source versions with different formatting.
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="BaseStyle">` +
			`<w:name w:val="Base Style"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="DerivedStyle">` +
			`<w:name w:val="Derived Style"/>` +
			`<w:basedOn w:val="BaseStyle"/>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		srcStyles.RawElement().AddChild(el)
	}
	rcInjectStyledParagraph(t, source, "DerivedStyle", "chain text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source:  source,
		Format:  KeepSourceFormatting,
		Options: ImportFormatOptions{ForceCopyStyles: true},
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Both styles should be renamed (_0 suffix).
	tgtStyles, _ = target.part.Styles()
	if tgtStyles.GetByID("DerivedStyle_0") == nil {
		t.Error("expected DerivedStyle_0 in target")
	}
	if tgtStyles.GetByID("BaseStyle_0") == nil {
		t.Error("expected BaseStyle_0 in target")
	}

	// Verify both renamed styles have semiHidden marker.
	for _, sid := range []string{"DerivedStyle_0", "BaseStyle_0"} {
		s := tgtStyles.GetByID(sid)
		if s != nil && findChild(s.RawElement(), "w", "semiHidden") == nil {
			t.Errorf("expected semiHidden on %s", sid)
		}
	}
}

func TestRWC_KeepDifferent_DifferentAlwaysCopied(t *testing.T) {
	t.Parallel()
	// KeepDifferentStyles: different formatting → always copy with suffix.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="DiffStyle">` +
		`<w:name w:val="Diff Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="DiffStyle">` +
		`<w:name w:val="Diff Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:b/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "DiffStyle", "copied text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// DiffStyle_0 must be created in target styles.
	tgtStyles, _ = target.part.Styles()
	copied := tgtStyles.GetByID("DiffStyle_0")
	if copied == nil {
		t.Fatal("expected DiffStyle_0 in target styles")
	}

	// Copied style must have semiHidden marker.
	if findChild(copied.RawElement(), "w", "semiHidden") == nil {
		t.Error("expected semiHidden on DiffStyle_0")
	}

	// Paragraph must reference DiffStyle_0 (not default or original).
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		ps := findChild(pPr, "w", "pStyle")
		if ps != nil && ps.SelectAttrValue("w:val", "") == "DiffStyle_0" {
			// jc must NOT be injected as direct attr.
			jc := findChild(pPr, "w", "jc")
			if jc != nil {
				t.Error("jc should not be injected as direct attribute when style is copied")
			}
			return
		}
	}
	t.Error("expected paragraph referencing DiffStyle_0")
}

func TestRWC_KeepDifferent_SameReusesTarget(t *testing.T) {
	t.Parallel()
	// Same formatting in src and tgt → no copy, paragraph uses target styleId.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	xml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="SameStyle">` +
		`<w:name w:val="Same Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	el, _ := oxml.ParseXml([]byte(xml))
	tgtStyles.RawElement().AddChild(el)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="SameStyle">` +
		`<w:name w:val="Same Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "SameStyle", "same style text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// SameStyle_0 must NOT exist.
	tgtStyles, _ = target.part.Styles()
	if tgtStyles.GetByID("SameStyle_0") != nil {
		t.Error("SameStyle_0 should not exist — styles are identical")
	}

	// Paragraph must reference "SameStyle" (original target ID).
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		ps := findChild(pPr, "w", "pStyle")
		if ps != nil && ps.SelectAttrValue("w:val", "") == "SameStyle" {
			return
		}
	}
	t.Error("expected paragraph referencing SameStyle (target ID)")
}

func TestRWC_KeepDifferent_MissingCopiedAsIs(t *testing.T) {
	t.Parallel()
	// Style absent in target → copied under original ID (no suffix).
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="UniqueStyle">` +
		`<w:name w:val="Unique Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "UniqueStyle", "unique text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// UniqueStyle must exist in target under original ID.
	tgtStyles, _ := target.part.Styles()
	if tgtStyles.GetByID("UniqueStyle") == nil {
		t.Error("expected UniqueStyle copied as-is to target")
	}
	// No suffixed version.
	if tgtStyles.GetByID("UniqueStyle_0") != nil {
		t.Error("UniqueStyle_0 should not exist — no conflict")
	}
}

func TestRWC_KeepDifferent_ForceCopyIgnored(t *testing.T) {
	t.Parallel()
	// ForceCopyStyles=true + different styles → result identical to false.
	makeDoc := func(forceCopy bool) *Document {
		target := mustNewDoc(t)
		target.AddParagraph("[<TAG>]")
		tgtStyles, _ := target.part.Styles()
		tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="FCSty">` +
			`<w:name w:val="FC Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`
		tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
		tgtStyles.RawElement().AddChild(tgtEl)

		source := mustNewDoc(t)
		srcStyles, _ := source.part.Styles()
		srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="FCSty">` +
			`<w:name w:val="FC Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`
		srcEl, _ := oxml.ParseXml([]byte(srcXml))
		srcStyles.RawElement().AddChild(srcEl)
		rcInjectStyledParagraph(t, source, "FCSty", "fc text")

		_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
			Source:  source,
			Format:  KeepDifferentStyles,
			Options: ImportFormatOptions{ForceCopyStyles: forceCopy},
		})
		if err != nil {
			t.Fatalf("ReplaceWithContent(force=%v): %v", forceCopy, err)
		}
		return target
	}

	docFalse := makeDoc(false)
	docTrue := makeDoc(true)

	// Both must have FCSty_0.
	s1, _ := docFalse.part.Styles()
	s2, _ := docTrue.part.Styles()
	if s1.GetByID("FCSty_0") == nil {
		t.Error("ForceCopyStyles=false: expected FCSty_0")
	}
	if s2.GetByID("FCSty_0") == nil {
		t.Error("ForceCopyStyles=true: expected FCSty_0")
	}
}

func TestRWC_KeepDifferent_ChainBothDifferent(t *testing.T) {
	t.Parallel()
	// Child basedOn Parent, both different → both copied with suffix,
	// basedOn in Child_0 remapped to Parent_0.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ParStyle">` +
			`<w:name w:val="Par Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ChiStyle">` +
			`<w:name w:val="Chi Style"/><w:basedOn w:val="ParStyle"/>` +
			`<w:rPr><w:i/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		tgtStyles.RawElement().AddChild(el)
	}

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ParStyle">` +
			`<w:name w:val="Par Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ChiStyle">` +
			`<w:name w:val="Chi Style"/><w:basedOn w:val="ParStyle"/>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		srcStyles.RawElement().AddChild(el)
	}
	rcInjectStyledParagraph(t, source, "ChiStyle", "chain text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	tgtStyles, _ = target.part.Styles()
	if tgtStyles.GetByID("ParStyle_0") == nil {
		t.Error("expected ParStyle_0 in target")
	}
	if tgtStyles.GetByID("ChiStyle_0") == nil {
		t.Error("expected ChiStyle_0 in target")
	}

	// Verify basedOn in ChiStyle_0 points to ParStyle_0.
	chi0 := tgtStyles.GetByID("ChiStyle_0")
	if chi0 != nil {
		bo := findChild(chi0.RawElement(), "w", "basedOn")
		if bo == nil {
			t.Error("ChiStyle_0 missing basedOn")
		} else if bo.SelectAttrValue("w:val", "") != "ParStyle_0" {
			t.Errorf("ChiStyle_0 basedOn = %q, want ParStyle_0",
				bo.SelectAttrValue("w:val", ""))
		}
	}
}

func TestRWC_KeepDifferent_ChainParentInTarget(t *testing.T) {
	t.Parallel()
	// Style different, basedOn points to target style (identical) →
	// style copied with suffix, basedOn points to original target style.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="BaseOk">` +
			`<w:name w:val="Base Ok"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ChildDiff">` +
			`<w:name w:val="Child Diff"/><w:basedOn w:val="BaseOk"/>` +
			`<w:rPr><w:i/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		tgtStyles.RawElement().AddChild(el)
	}

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	for _, s := range []string{
		// Same BaseOk in source (identical to target).
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="BaseOk">` +
			`<w:name w:val="Base Ok"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`,
		// Different ChildDiff.
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ChildDiff">` +
			`<w:name w:val="Child Diff"/><w:basedOn w:val="BaseOk"/>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		srcStyles.RawElement().AddChild(el)
	}
	rcInjectStyledParagraph(t, source, "ChildDiff", "child diff text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	tgtStyles, _ = target.part.Styles()
	// BaseOk should NOT be copied (same formatting).
	if tgtStyles.GetByID("BaseOk_0") != nil {
		t.Error("BaseOk_0 should not exist — parent is identical")
	}
	// ChildDiff should be copied.
	child0 := tgtStyles.GetByID("ChildDiff_0")
	if child0 == nil {
		t.Fatal("expected ChildDiff_0 in target")
	}
	// basedOn should point to original BaseOk.
	bo := findChild(child0.RawElement(), "w", "basedOn")
	if bo == nil {
		t.Error("ChildDiff_0 missing basedOn")
	} else if bo.SelectAttrValue("w:val", "") != "BaseOk" {
		t.Errorf("ChildDiff_0 basedOn = %q, want BaseOk",
			bo.SelectAttrValue("w:val", ""))
	}
}

func TestRWC_KeepDifferent_CharacterStyle(t *testing.T) {
	t.Parallel()
	// Different character style → copied with suffix, rPr preserved.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="character" w:styleId="CharSty">` +
		`<w:name w:val="Char Style"/><w:rPr><w:i/></w:rPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="character" w:styleId="CharSty">` +
		`<w:name w:val="Char Style"/><w:rPr><w:b/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	// Inject a paragraph with a run referencing CharSty.
	srcBody := source.element.Body().RawElement()
	pEl, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:r><w:rPr><w:rStyle w:val="CharSty"/></w:rPr><w:t>char test</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, pEl)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	tgtStyles, _ = target.part.Styles()
	copied := tgtStyles.GetByID("CharSty_0")
	if copied == nil {
		t.Fatal("expected CharSty_0 in target")
	}
	// Verify rPr with bold preserved.
	rPr := findChild(copied.RawElement(), "w", "rPr")
	if rPr == nil {
		t.Error("expected rPr on copied character style")
	} else if findChild(rPr, "w", "b") == nil {
		t.Error("expected bold in copied character style rPr")
	}
}

func TestRWC_KeepDifferent_MultipleDifferent(t *testing.T) {
	t.Parallel()
	// Three styles with different formatting → suffixes _0, _1, _2.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()

	for _, name := range []string{"StyleA", "StyleB", "StyleC"} {
		// Target versions: jc=left.
		tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="` + name + `">` +
			`<w:name w:val="` + name + `"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`
		tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
		tgtStyles.RawElement().AddChild(tgtEl)

		// Source versions: jc=center.
		srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="` + name + `">` +
			`<w:name w:val="` + name + `"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`
		srcEl, _ := oxml.ParseXml([]byte(srcXml))
		srcStyles.RawElement().AddChild(srcEl)
	}
	// Add paragraphs referencing all three.
	for _, name := range []string{"StyleA", "StyleB", "StyleC"} {
		rcInjectStyledParagraph(t, source, name, name+" text")
	}

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	tgtStyles, _ = target.part.Styles()
	for _, name := range []string{"StyleA_0", "StyleB_0", "StyleC_0"} {
		if tgtStyles.GetByID(name) == nil {
			t.Errorf("expected %s in target styles", name)
		}
	}
}

func TestRWC_KeepDifferent_RootStyleCompensateAll(t *testing.T) {
	t.Parallel()
	// Root style (no basedOn), different formatting, src/tgt docDefaults
	// differ (src sz=24, tgt sz=20) → copy with suffix + compensateAll
	// delta injected into pPr/rPr.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	// Set target docDefaults: sz=20.
	ddXml := `<w:docDefaults xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>` +
		`</w:docDefaults>`
	ddEl, _ := oxml.ParseXml([]byte(ddXml))
	tgtStyles.RawElement().InsertChildAt(0, ddEl)
	// Target style: no own rPr.
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="RootSty">` +
		`<w:name w:val="Root Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	// Source docDefaults: sz=24.
	srcDdXml := `<w:docDefaults xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>` +
		`</w:docDefaults>`
	srcDdEl, _ := oxml.ParseXml([]byte(srcDdXml))
	srcStyles.RawElement().InsertChildAt(0, srcDdEl)
	// Source style: different pPr (jc=center), no basedOn.
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="RootSty">` +
		`<w:name w:val="Root Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "RootSty", "root compensate text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	tgtStyles, _ = target.part.Styles()
	copied := tgtStyles.GetByID("RootSty_0")
	if copied == nil {
		t.Fatal("expected RootSty_0 in target")
	}
	// Verify compensateAll: delta sz=24 injected into rPr.
	rPr := findChild(copied.RawElement(), "w", "rPr")
	if rPr == nil {
		t.Fatal("expected rPr with docDefaults compensation on RootSty_0")
	}
	szEl := findChild(rPr, "w", "sz")
	if szEl == nil {
		t.Fatal("expected sz element in RootSty_0 rPr (docDefaults delta)")
	}
	if v := szEl.SelectAttrValue("w:val", ""); v != "24" {
		t.Errorf("sz val = %q, want 24 (source docDefaults compensation)", v)
	}
}

func TestRWC_KeepDifferent_DocDefaultsCompensation(t *testing.T) {
	t.Parallel()
	// Different style + src/tgt docDefaults differ → copy with suffix
	// + Pass 2 fixupCopiedStyles applies compensation. Verify rPr sz delta.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	ddXml := `<w:docDefaults xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>` +
		`<w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault>` +
		`</w:docDefaults>`
	ddEl, _ := oxml.ParseXml([]byte(ddXml))
	tgtStyles.RawElement().InsertChildAt(0, ddEl)
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="DDSty">` +
		`<w:name w:val="DD Style"/><w:rPr><w:i/></w:rPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcDdXml := `<w:docDefaults xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>` +
		`<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>` +
		`</w:docDefaults>`
	srcDdEl, _ := oxml.ParseXml([]byte(srcDdXml))
	srcStyles.RawElement().InsertChildAt(0, srcDdEl)
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="DDSty">` +
		`<w:name w:val="DD Style"/><w:rPr><w:b/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "DDSty", "docdefaults text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	tgtStyles, _ = target.part.Styles()
	copied := tgtStyles.GetByID("DDSty_0")
	if copied == nil {
		t.Fatal("expected DDSty_0 in target")
	}
	// rPr should contain sz delta (24) and original bold.
	rPr := findChild(copied.RawElement(), "w", "rPr")
	if rPr == nil {
		t.Fatal("expected rPr on DDSty_0")
	}
	szEl := findChild(rPr, "w", "sz")
	if szEl == nil {
		t.Fatal("expected sz in DDSty_0 rPr (docDefaults compensation)")
	}
	if v := szEl.SelectAttrValue("w:val", ""); v != "24" {
		t.Errorf("sz val = %q, want 24", v)
	}
	if findChild(rPr, "w", "b") == nil {
		t.Error("expected bold preserved in DDSty_0 rPr")
	}
	// pPr should contain spacing delta.
	pPr := findChild(copied.RawElement(), "w", "pPr")
	if pPr != nil {
		sp := findChild(pPr, "w", "spacing")
		if sp != nil {
			if v := sp.SelectAttrValue("w:after", ""); v != "160" {
				t.Errorf("spacing after = %q, want 160", v)
			}
		}
	}
}

func TestRWC_KeepDifferent_RoundTrip(t *testing.T) {
	t.Parallel()
	// Full round-trip: insert + Save + reopen + verify.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="RTStyle">` +
		`<w:name w:val="RT Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="RTStyle">` +
		`<w:name w:val="RT Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "RTStyle", "round trip text")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepDifferentStyles,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Save and reopen.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	reopened, err := Open(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	// Verify copied style survived round-trip.
	reStyles, _ := reopened.part.Styles()
	if reStyles.GetByID("RTStyle_0") == nil {
		t.Error("expected RTStyle_0 after round-trip")
	}

	// Verify paragraph references.
	body, _ := reopened.getBody()
	found := false
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		ps := findChild(pPr, "w", "pStyle")
		if ps != nil && ps.SelectAttrValue("w:val", "") == "RTStyle_0" {
			found = true
			break
		}
	}
	if !found {
		t.Error("expected paragraph referencing RTStyle_0 after round-trip")
	}
}

func TestRWC_BackwardCompat_ZeroValue(t *testing.T) {
	t.Parallel()
	// ContentData{Source: doc} with zero-value Format/Options must behave
	// exactly like UseDestinationStyles (the backward-compatible default).
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="MyStyle">` +
		`<w:name w:val="My Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="MyStyle">` +
		`<w:name w:val="My Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)
	rcInjectStyledParagraph(t, source, "MyStyle", "zero value mode")

	// Zero-value ContentData — no Format set.
	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// UseDestinationStyles: pStyle stays MyStyle, no jc=center expansion.
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		jc := findChild(pPr, "w", "jc")
		if jc != nil && jc.SelectAttrValue("w:val", "") == "center" {
			t.Error("zero-value Format should NOT expand jc=center (UseDestinationStyles)")
		}
	}
}

// --------------------------------------------------------------------------
// Expand to Direct Attributes — additional coverage (Step 8.3)
// --------------------------------------------------------------------------

func TestRWC_Expand_RunStyle_Expanded(t *testing.T) {
	t.Parallel()
	// Character style (rStyle) on a run should be expanded to direct rPr
	// when KeepSourceFormatting is active and the style conflicts.
	// Uses unique styleId to avoid collision with built-in Emphasis.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="character" w:styleId="RunCharCustom">` +
		`<w:name w:val="Run Char Custom"/><w:rPr><w:i/></w:rPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="character" w:styleId="RunCharCustom">` +
		`<w:name w:val="Run Char Custom"/><w:rPr><w:b/><w:i/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	// Inject run with rStyle=RunCharCustom into source body.
	srcBody := source.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:r><w:rPr><w:rStyle w:val="RunCharCustom"/></w:rPr>` +
			`<w:t>emphasized text</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Run should have bold expanded from source RunCharCustom style.
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		for _, r := range child.ChildElements() {
			if r.Space != "w" || r.Tag != "r" {
				continue
			}
			rPr := findChild(r, "w", "rPr")
			if rPr != nil && findChild(rPr, "w", "b") != nil {
				return // OK — bold was expanded from character style
			}
		}
	}
	t.Error("expected bold expanded from source character style on run")
}

func TestRWC_Expand_MixedStyles_BothExpanded(t *testing.T) {
	t.Parallel()
	// Both paragraph style and character style conflict: both should be
	// expanded to direct attributes on the same paragraph/run.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ParaStyle">` +
			`<w:name w:val="Para Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="character" w:styleId="CharStyle">` +
			`<w:name w:val="Char Style"/><w:rPr><w:i/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		tgtStyles.RawElement().AddChild(el)
	}

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="ParaStyle">` +
			`<w:name w:val="Para Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="character" w:styleId="CharStyle">` +
			`<w:name w:val="Char Style"/><w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		srcStyles.RawElement().AddChild(el)
	}

	srcBody := source.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="ParaStyle"/></w:pPr>` +
			`<w:r><w:rPr><w:rStyle w:val="CharStyle"/></w:rPr>` +
			`<w:t>both expanded</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Verify: jc=center from paragraph style AND bold from character style.
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		jc := findChild(pPr, "w", "jc")
		if jc == nil || jc.SelectAttrValue("w:val", "") != "center" {
			continue
		}
		// Paragraph style expanded. Now check run.
		for _, r := range child.ChildElements() {
			if r.Space != "w" || r.Tag != "r" {
				continue
			}
			rPr := findChild(r, "w", "rPr")
			if rPr != nil && findChild(rPr, "w", "b") != nil {
				return // Both expanded
			}
		}
	}
	t.Error("expected both paragraph (jc=center) and character (bold) styles expanded")
}

func TestRWC_Expand_ExistingDirectWins(t *testing.T) {
	t.Parallel()
	// Direct attributes on a run must NOT be overwritten by expanded style.
	// Uses unique styleId to avoid collision with built-in Strong.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="character" w:styleId="DirectWinsChar">` +
		`<w:name w:val="Direct Wins Char"/><w:rPr><w:i/></w:rPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="character" w:styleId="DirectWinsChar">` +
		`<w:name w:val="Direct Wins Char"/><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	// Run with direct sz=24 (should NOT be overwritten by style's sz=28).
	srcBody := source.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:r><w:rPr><w:rStyle w:val="DirectWinsChar"/><w:sz w:val="24"/></w:rPr>` +
			`<w:t>direct wins</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		for _, r := range child.ChildElements() {
			if r.Space != "w" || r.Tag != "r" {
				continue
			}
			rPr := findChild(r, "w", "rPr")
			if rPr == nil {
				continue
			}
			sz := findChild(rPr, "w", "sz")
			if sz == nil {
				continue
			}
			if sz.SelectAttrValue("w:val", "") == "24" {
				// Direct sz=24 preserved. color=FF0000 should also be expanded.
				color := findChild(rPr, "w", "color")
				if color == nil || color.SelectAttrValue("w:val", "") != "FF0000" {
					t.Error("expected color=FF0000 expanded from style")
				}
				return
			}
			if sz.SelectAttrValue("w:val", "") == "28" {
				t.Error("direct sz=24 was overwritten by style sz=28")
				return
			}
		}
	}
	t.Error("expected run with direct sz=24 preserved")
}

func TestRWC_Expand_DeepMerge_RFonts(t *testing.T) {
	t.Parallel()
	// <w:rFonts> attributes should merge at attribute level: if direct has
	// w:ascii and style has w:hAnsi, the result should contain both.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="FontMerge">` +
		`<w:name w:val="Font Merge"/>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="FontMerge">` +
		`<w:name w:val="Font Merge"/>` +
		`<w:rPr><w:rFonts w:ascii="Times" w:hAnsi="Times"/></w:rPr>` +
		`</w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	// Paragraph with existing rFonts w:ascii="Arial" (should win).
	srcBody := source.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="FontMerge"/>` +
			`<w:rPr><w:rFonts w:ascii="Arial"/></w:rPr></w:pPr>` +
			`<w:r><w:t>font merge test</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Verify rFonts has both ascii=Arial (direct) and hAnsi=Times (from style).
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		rPr := findChild(pPr, "w", "rPr")
		if rPr == nil {
			continue
		}
		rFonts := findChild(rPr, "w", "rFonts")
		if rFonts == nil {
			continue
		}
		ascii := rFonts.SelectAttrValue("w:ascii", "")
		hAnsi := rFonts.SelectAttrValue("w:hAnsi", "")
		if ascii == "Arial" && hAnsi == "Times" {
			return // deep merge worked
		}
	}
	t.Error("expected rFonts with ascii=Arial (direct) and hAnsi=Times (style)")
}

func TestRWC_Expand_BasedOnChain_3Levels(t *testing.T) {
	t.Parallel()
	// 3-level basedOn chain: Grandchild → Child → Base.
	// Expansion should resolve all three levels.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="Grandchild">` +
		`<w:name w:val="Grandchild"/>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="Base">` +
			`<w:name w:val="Base"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`<w:rPr><w:sz w:val="20"/></w:rPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="Child">` +
			`<w:name w:val="Child"/>` +
			`<w:basedOn w:val="Base"/>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="Grandchild">` +
			`<w:name w:val="Grandchild"/>` +
			`<w:basedOn w:val="Child"/>` +
			`<w:rPr><w:i/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		srcStyles.RawElement().AddChild(el)
	}
	rcInjectStyledParagraph(t, source, "Grandchild", "3-level chain")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Verify resolved properties: jc=center (Base), bold (Child), italic (Grandchild).
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		pPr := findChild(child, "w", "pPr")
		if pPr == nil {
			continue
		}
		jc := findChild(pPr, "w", "jc")
		if jc == nil || jc.SelectAttrValue("w:val", "") != "center" {
			continue
		}
		rPr := findChild(pPr, "w", "rPr")
		if rPr == nil {
			continue
		}
		hasB := findChild(rPr, "w", "b") != nil
		hasI := findChild(rPr, "w", "i") != nil
		hasSz := findChild(rPr, "w", "sz") != nil
		if hasB && hasI && hasSz {
			return // all 3 levels resolved
		}
		t.Errorf("missing properties: bold=%v, italic=%v, sz=%v", hasB, hasI, hasSz)
		return
	}
	t.Error("expected paragraph with 3-level basedOn chain resolved")
}

func TestRWC_Expand_CyclicBasedOn_NoPanic(t *testing.T) {
	t.Parallel()
	// Cyclic basedOn: A → B → A. Must not infinite loop or panic.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	tgtStyles, _ := target.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="CyclicA">` +
		`<w:name w:val="Cyclic A"/>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)

	source := mustNewDoc(t)
	srcStyles, _ := source.part.Styles()
	for _, s := range []string{
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="CyclicA">` +
			`<w:name w:val="Cyclic A"/>` +
			`<w:basedOn w:val="CyclicB"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
		`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:type="paragraph" w:styleId="CyclicB">` +
			`<w:name w:val="Cyclic B"/>` +
			`<w:basedOn w:val="CyclicA"/>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	} {
		el, _ := oxml.ParseXml([]byte(s))
		srcStyles.RawElement().AddChild(el)
	}
	rcInjectStyledParagraph(t, source, "CyclicA", "cyclic test")

	// Must complete without panic.
	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		Format: KeepSourceFormatting,
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v (should not error on cycle)", err)
	}
}

// --------------------------------------------------------------------------
// Edge Cases (Step 8.6)
// --------------------------------------------------------------------------

func TestRWC_SDT_Preserved(t *testing.T) {
	t.Parallel()
	// SDT (Structured Document Tag) elements must survive copy.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	srcBody := source.element.Body().RawElement()
	sdt, _ := oxml.ParseXml([]byte(
		`<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:sdtPr><w:alias w:val="TestSDT"/></w:sdtPr>` +
			`<w:sdtContent><w:p><w:r><w:t>SDT content</w:t></w:r></w:p></w:sdtContent>` +
			`</w:sdt>`,
	))
	rcInsertBeforeSectPr(srcBody, sdt)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Verify SDT element exists in target body.
	body, _ := target.getBody()
	for _, child := range body.Element().ChildElements() {
		if child.Space == "w" && child.Tag == "sdt" {
			sdtPr := findChild(child, "w", "sdtPr")
			if sdtPr != nil {
				alias := findChild(sdtPr, "w", "alias")
				if alias != nil && alias.SelectAttrValue("w:val", "") == "TestSDT" {
					return // OK
				}
			}
		}
	}
	t.Error("expected SDT element with alias=TestSDT in target body")
}

func TestRWC_FieldCode_Preserved(t *testing.T) {
	t.Parallel()
	// Field codes (w:fldChar, w:instrText) must survive copy.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	srcBody := source.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
			`<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
			`<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
			`<w:r><w:t>1</w:t></w:r>` +
			`<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
			`</w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Find fldChar elements in target.
	body, _ := target.getBody()
	fldCharCount := 0
	instrCount := 0
	for _, child := range body.Element().ChildElements() {
		if child.Space != "w" || child.Tag != "p" {
			continue
		}
		for _, r := range child.ChildElements() {
			if r.Space != "w" || r.Tag != "r" {
				continue
			}
			for _, inner := range r.ChildElements() {
				if inner.Space == "w" && inner.Tag == "fldChar" {
					fldCharCount++
				}
				if inner.Space == "w" && inner.Tag == "instrText" {
					instrCount++
				}
			}
		}
	}
	if fldCharCount < 3 {
		t.Errorf("expected at least 3 fldChar elements, got %d", fldCharCount)
	}
	if instrCount < 1 {
		t.Errorf("expected at least 1 instrText element, got %d", instrCount)
	}
}

func TestRWC_100Placeholders_SameTag(t *testing.T) {
	t.Parallel()
	// Scaling test: 100 occurrences of the same tag.
	target := mustNewDoc(t)
	for i := 0; i < 100; i++ {
		target.AddParagraph("[<X>]")
	}

	source := rcSourceWithParagraph(t, "bulk")

	count, err := target.ReplaceWithContent("[<X>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}
	if count != 100 {
		t.Errorf("count = %d, want 100", count)
	}

	n := 0
	for _, txt := range rcBodyTexts(t, target) {
		if strings.Contains(txt, "bulk") {
			n++
		}
	}
	if n < 100 {
		t.Errorf("expected 'bulk' at least 100 times, got %d", n)
	}
}

func TestRWC_DeeplyNestedTables_5Levels(t *testing.T) {
	t.Parallel()
	// 5-level nested table structure in source must survive copy.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	srcBody := source.element.Body().RawElement()

	// Build 5-level nested table: each table has a single cell containing
	// another table, down to level 5 which has text.
	inner := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:t>deepest cell</w:t></w:r></w:p>`
	for i := 0; i < 5; i++ {
		inner = `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:tr><w:tc>` + inner + `</w:tc></w:tr></w:tbl>`
	}
	tbl, _ := oxml.ParseXml([]byte(inner))
	rcInsertBeforeSectPr(srcBody, tbl)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Find "deepest cell" text recursively (rcBodyTexts only scans top-level paragraphs).
	body, _ := target.getBody()
	found := rcFindTextRecursive(body.Element(), "deepest cell")
	if !found {
		t.Error("expected 'deepest cell' text from 5-level nested table")
	}
}

func TestRWC_UnicodeTag_Cyrillic(t *testing.T) {
	t.Parallel()
	target := mustNewDoc(t)
	target.AddParagraph("Начало [<СОДЕРЖАНИЕ>] Конец")

	source := rcSourceWithParagraph(t, "Кириллица вставлена")

	count, err := target.ReplaceWithContent("[<СОДЕРЖАНИЕ>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	allText := ""
	for _, txt := range rcBodyTexts(t, target) {
		allText += txt
	}
	if !strings.Contains(allText, "Кириллица вставлена") {
		t.Error("expected inserted cyrillic text")
	}
	if !strings.Contains(allText, "Начало") {
		t.Error("surrounding cyrillic text lost")
	}
}

func TestRWC_KeepSourceNumbering_SeparateList(t *testing.T) {
	t.Parallel()
	// When KeepSourceNumbering=true, source list definitions must be imported
	// as separate lists (not merged into target).
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	// Add a numbered list to target (decimal, numId=1).
	rcSetupNumbering(t, target, 1, 0, "decimal")

	// Count baseline abstractNums before replacement.
	tgtNP, err := target.part.GetOrAddNumberingPart()
	if err != nil {
		t.Fatalf("GetOrAddNumberingPart: %v", err)
	}
	tgtNumbering, err := tgtNP.Numbering()
	if err != nil {
		t.Fatalf("Numbering: %v", err)
	}
	absNumsBefore := len(tgtNumbering.AllAbstractNums())

	// Create source with its own numbered list (decimal, numId=1).
	source := mustNewDoc(t)
	rcSetupNumbering(t, source, 1, 0, "decimal")
	rcInjectNumberedParagraph(t, source, 1, 0, "Source item")

	_, replaceErr := target.ReplaceWithContent("[<TAG>]", ContentData{
		Source:  source,
		Options: ImportFormatOptions{KeepSourceNumbering: true},
	})
	if replaceErr != nil {
		t.Fatalf("ReplaceWithContent: %v", replaceErr)
	}

	// Target should have the inserted text.
	found := false
	for _, txt := range rcBodyTexts(t, target) {
		if strings.Contains(txt, "Source item") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'Source item' in target")
	}

	// Target numbering should have a new abstractNum (imported separately).
	absNumsAfter := len(tgtNumbering.AllAbstractNums())
	if absNumsAfter <= absNumsBefore {
		t.Errorf("expected new abstractNum (KeepSourceNumbering=true), before=%d after=%d",
			absNumsBefore, absNumsAfter)
	}
}

func TestRWC_KeepSourceNumbering_False_Merged(t *testing.T) {
	t.Parallel()
	// When KeepSourceNumbering=false (default), source list with matching
	// numFmt should be merged into existing target list.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	// Add a numbered list to target (decimal).
	rcSetupNumbering(t, target, 1, 0, "decimal")

	// Count baseline abstractNums before replacement.
	tgtNP, err := target.part.GetOrAddNumberingPart()
	if err != nil {
		t.Fatalf("GetOrAddNumberingPart: %v", err)
	}
	tgtNumbering, err := tgtNP.Numbering()
	if err != nil {
		t.Fatalf("Numbering: %v", err)
	}
	absNumsBefore := len(tgtNumbering.AllAbstractNums())

	// Create source with its own decimal list.
	source := mustNewDoc(t)
	rcSetupNumbering(t, source, 1, 0, "decimal")
	rcInjectNumberedParagraph(t, source, 1, 0, "Merged item")

	_, err = target.ReplaceWithContent("[<TAG>]", ContentData{
		Source: source,
		// KeepSourceNumbering defaults to false → merge
	})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	found := false
	for _, txt := range rcBodyTexts(t, target) {
		if strings.Contains(txt, "Merged item") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'Merged item' in target")
	}

	// With merge, no new abstractNum should have been created (matching
	// numFmt reuses existing target definition).
	absNumsAfter := len(tgtNumbering.AllAbstractNums())
	if absNumsAfter != absNumsBefore {
		t.Errorf("expected no new abstractNums (merged), got delta=%d", absNumsAfter-absNumsBefore)
	}
}

// rcSetupNumberingMultiLevel creates a multi-level numbering definition
// in the document. levels is a slice of [numFmt, lvlText] pairs.
// Returns the actual numId assigned (auto-incremented by the numbering part).
func rcSetupNumberingMultiLevel(
	t *testing.T, doc *Document,
	levels [][2]string,
) int {
	t.Helper()
	np, err := doc.part.GetOrAddNumberingPart()
	if err != nil {
		t.Fatalf("GetOrAddNumberingPart: %v", err)
	}
	numbering, err := np.Numbering()
	if err != nil {
		t.Fatalf("Numbering: %v", err)
	}

	absId := numbering.NextAbstractNumId()
	lvlsXml := ""
	for i, lv := range levels {
		lvlsXml += `<w:lvl w:ilvl="` + strconv.Itoa(i) + `">` +
			`<w:numFmt w:val="` + lv[0] + `"/>` +
			`<w:lvlText w:val="` + lv[1] + `"/>` +
			`</w:lvl>`
	}
	absNum, _ := oxml.ParseXml([]byte(
		`<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:abstractNumId="` + strconv.Itoa(absId) + `">` +
			`<w:nsid w:val="AABB0011"/>` + lvlsXml +
			`</w:abstractNum>`,
	))
	numbering.InsertAbstractNum(absNum)
	num, err := numbering.AddNumWithAbstractNumId(absId)
	if err != nil {
		t.Fatalf("AddNumWithAbstractNumId: %v", err)
	}
	numId, err := num.NumId()
	if err != nil {
		t.Fatalf("NumId: %v", err)
	}
	return numId
}

func TestRWC_NumMerge_MultiLevel_Compatible(t *testing.T) {
	t.Parallel()
	// Source and target with 3-level list, all match →
	// absNums count unchanged (merged).
	// Uses taiwaneseCounting to avoid accidental match with default template.
	levels := [][2]string{
		{"taiwaneseCounting", "第%1"},
		{"ideographTraditional", "(%2)"},
		{"ideographZodiac", "[%3]"},
	}
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	rcSetupNumberingMultiLevel(t, target, levels)

	tgtNP, _ := target.part.GetOrAddNumberingPart()
	tgtNum, _ := tgtNP.Numbering()
	before := len(tgtNum.AllAbstractNums())

	source := mustNewDoc(t)
	srcNumId := rcSetupNumberingMultiLevel(t, source, levels)
	rcInjectNumberedParagraph(t, source, srcNumId, 0, "multi level item")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	after := len(tgtNum.AllAbstractNums())
	if after != before {
		t.Errorf("compatible multi-level → expected merge (no new absNum), before=%d after=%d",
			before, after)
	}
}

func TestRWC_NumMerge_MultiLevel_Incompatible(t *testing.T) {
	t.Parallel()
	// Target: 3 levels (taiwaneseCounting, upperLetter, upperRoman).
	// Source: 3 levels (taiwaneseCounting, lowerLetter, lowerRoman).
	// Level 0 matches, levels 1,2 differ → absNums increases (separate).
	// Uses taiwaneseCounting to avoid accidental match with default template.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	rcSetupNumberingMultiLevel(t, target, [][2]string{
		{"taiwaneseCounting", "第%1"},
		{"upperLetter", "%2)"},
		{"upperRoman", "%3."},
	})

	tgtNP, _ := target.part.GetOrAddNumberingPart()
	tgtNum, _ := tgtNP.Numbering()
	before := len(tgtNum.AllAbstractNums())

	source := mustNewDoc(t)
	srcNumId := rcSetupNumberingMultiLevel(t, source, [][2]string{
		{"taiwaneseCounting", "第%1"},
		{"lowerLetter", "%2)"},
		{"lowerRoman", "%3."},
	})
	rcInjectNumberedParagraph(t, source, srcNumId, 0, "incompatible item")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	after := len(tgtNum.AllAbstractNums())
	if after <= before {
		t.Errorf("incompatible multi-level → expected separate absNum, before=%d after=%d",
			before, after)
	}
}

func TestRWC_NumMerge_DifferentLvlText_Separate(t *testing.T) {
	t.Parallel()
	// Both taiwaneseCounting + lvl 0, but "第%1" vs "第%1。" →
	// absNums increases (not merged despite same numFmt).
	// Uses taiwaneseCounting to avoid accidental match with default template.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")
	rcSetupNumberingMultiLevel(t, target, [][2]string{{"taiwaneseCounting", "第%1"}})

	tgtNP, _ := target.part.GetOrAddNumberingPart()
	tgtNum, _ := tgtNP.Numbering()
	before := len(tgtNum.AllAbstractNums())

	source := mustNewDoc(t)
	srcNumId := rcSetupNumberingMultiLevel(t, source, [][2]string{{"taiwaneseCounting", "第%1。"}})
	rcInjectNumberedParagraph(t, source, srcNumId, 0, "different text item")

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	after := len(tgtNum.AllAbstractNums())
	if after <= before {
		t.Errorf("different lvlText → expected separate absNum, before=%d after=%d",
			before, after)
	}
}

func TestRWC_DeepImport_RoundTrip(t *testing.T) {
	t.Parallel()
	// Integration test: deep import with save → reopen.
	// Use an OLE-like part (generic internal relationship) to exercise
	// the deep import path end-to-end through ReplaceWithContent.
	target := mustNewDoc(t)
	target.AddParagraph("[<TAG>]")

	source := mustNewDoc(t)
	source.AddParagraph("with deep part")

	// Add a generic internal part to source (simulates chart/diagram).
	genericBlob := []byte(`<?xml version="1.0"?><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`)
	genericPart, _ := opc.NewXmlPart(
		"/word/charts/chart1.xml",
		"application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
		genericBlob,
		source.wmlPkg.OpcPackage,
	)
	source.wmlPkg.OpcPackage.AddPart(genericPart)
	rel := source.part.Rels().GetOrAdd(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
		genericPart,
	)

	// Inject a drawing referencing the chart via r:id.
	srcBody := source.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<w:r><w:drawing><c:chart r:id="` + rel.RID + `" ` +
			`xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/></w:drawing></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)

	_, err := target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	if err != nil {
		t.Fatalf("ReplaceWithContent: %v", err)
	}

	// Save and reopen.
	var buf bytes.Buffer
	if err := target.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	reopened, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Verify text survived round-trip.
	found := false
	for _, txt := range rcBodyTexts(t, reopened) {
		if strings.Contains(txt, "with deep part") {
			found = true
		}
	}
	if !found {
		t.Error("expected 'with deep part' in reopened document")
	}
}

// --------------------------------------------------------------------------
// Test helpers (Step 8)
// --------------------------------------------------------------------------

// rcInjectStyledParagraph injects a paragraph with pStyle into the source body.
func rcInjectStyledParagraph(t *testing.T, doc *Document, styleId, text string) {
	t.Helper()
	srcBody := doc.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="` + styleId + `"/></w:pPr>` +
			`<w:r><w:t>` + text + `</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)
}

// rcInsertBeforeSectPr inserts an element before the first w:sectPr in body.
func rcInsertBeforeSectPr(body, el *etree.Element) {
	children := body.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			body.InsertChildAt(i, el)
			return
		}
	}
	body.AddChild(el)
}

// rcSetupNumbering creates a numbering definition in the document with
// the given numId, abstractNumId, and numFmt (e.g., "decimal", "bullet").
func rcSetupNumbering(t *testing.T, doc *Document, numId, abstractNumId int, numFmt string) {
	t.Helper()
	np, err := doc.part.GetOrAddNumberingPart()
	if err != nil {
		t.Fatalf("GetOrAddNumberingPart: %v", err)
	}
	numbering, err := np.Numbering()
	if err != nil {
		t.Fatalf("Numbering: %v", err)
	}

	// Add abstractNum with a single level.
	absNum, _ := oxml.ParseXml([]byte(
		`<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`w:abstractNumId="` + strconv.Itoa(abstractNumId) + `">` +
			`<w:nsid w:val="AABB0011"/>` +
			`<w:lvl w:ilvl="0"><w:numFmt w:val="` + numFmt + `"/></w:lvl>` +
			`</w:abstractNum>`,
	))
	numbering.InsertAbstractNum(absNum)

	// Add num referencing the abstractNum.
	numbering.AddNumWithAbstractNumId(abstractNumId)
}

// rcFindTextRecursive searches the entire element tree for a <w:t> element
// containing the given substring. Used for deeply nested structures (tables
// within tables) where rcBodyTexts doesn't reach.
func rcFindTextRecursive(el *etree.Element, substr string) bool {
	stack := []*etree.Element{el}
	for len(stack) > 0 {
		node := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if node.Space == "w" && node.Tag == "t" {
			if strings.Contains(node.Text(), substr) {
				return true
			}
		}
		stack = append(stack, node.ChildElements()...)
	}
	return false
}

// rcInjectNumberedParagraph injects a paragraph with numPr into the source body.
func rcInjectNumberedParagraph(t *testing.T, doc *Document, numId, ilvl int, text string) {
	t.Helper()
	srcBody := doc.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:numPr>` +
			`<w:ilvl w:val="` + strconv.Itoa(ilvl) + `"/>` +
			`<w:numId w:val="` + strconv.Itoa(numId) + `"/>` +
			`</w:numPr></w:pPr>` +
			`<w:r><w:t>` + text + `</w:t></w:r></w:p>`,
	))
	rcInsertBeforeSectPr(srcBody, p)
}
