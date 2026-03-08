package docx

import (
	"bytes"
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
