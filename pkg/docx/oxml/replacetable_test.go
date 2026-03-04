package oxml

import (
	"strings"
	"testing"

	"github.com/beevik/etree"
)

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

// mkP builds a <w:p> from raw XML. Panics on parse error.
func mkP(xml string) *etree.Element {
	doc := etree.NewDocument()
	doc.ReadSettings.Permissive = true
	if err := doc.ReadFromString(xml); err != nil {
		panic("mkP: " + err.Error())
	}
	return doc.Root()
}

// fragSummary returns a compact representation of each fragment:
// "P(text)" for paragraphs (text = concatenated <w:t> contents), "X" for placeholders.
func fragSummary(frags []Fragment) []string {
	var result []string
	for _, f := range frags {
		if f.IsPlaceholder() {
			result = append(result, "X")
		} else {
			result = append(result, "P("+collectPText(f.Element())+")")
		}
	}
	return result
}

// collectPText concatenates all <w:t> text inside a <w:p>, including runs
// inside hyperlinks.
func collectPText(p *etree.Element) string {
	var sb strings.Builder
	for _, child := range p.ChildElements() {
		switch {
		case child.Space == "w" && child.Tag == "r":
			collectRunText(child, &sb)
		case child.Space == "w" && child.Tag == "hyperlink":
			for _, r := range child.ChildElements() {
				if r.Space == "w" && r.Tag == "r" {
					collectRunText(r, &sb)
				}
			}
		}
	}
	return sb.String()
}

func collectRunText(rEl *etree.Element, sb *strings.Builder) {
	for _, c := range rEl.ChildElements() {
		if c.Space == "w" && c.Tag == "t" {
			sb.WriteString(c.Text())
		}
	}
}

// hasPPr reports whether a <w:p> has a <w:pPr> child.
func hasPPr(p *etree.Element) bool {
	for _, c := range p.ChildElements() {
		if c.Space == "w" && c.Tag == "pPr" {
			return true
		}
	}
	return false
}

// pPrHasChild checks whether <w:pPr> of a <w:p> contains a child with the given tag.
func pPrHasChild(p *etree.Element, space, tag string) bool {
	for _, c := range p.ChildElements() {
		if c.Space == "w" && c.Tag == "pPr" {
			for _, gc := range c.ChildElements() {
				if gc.Space == space && gc.Tag == tag {
					return true
				}
			}
		}
	}
	return false
}

// firstRunHasRPrChild checks whether the first <w:r> in a <w:p> has a <w:rPr>
// child containing an element with the given tag.
func firstRunHasRPrChild(p *etree.Element, space, tag string) bool {
	for _, c := range p.ChildElements() {
		if c.Space == "w" && c.Tag == "r" {
			for _, rc := range c.ChildElements() {
				if rc.Space == "w" && rc.Tag == "rPr" {
					for _, rpc := range rc.ChildElements() {
						if rpc.Space == space && rpc.Tag == tag {
							return true
						}
					}
				}
			}
			return false // only check first run
		}
	}
	return false
}

// countChildrenByTag counts direct child elements with the given space:tag.
func countChildrenByTag(el *etree.Element, space, tag string) int {
	n := 0
	for _, c := range el.ChildElements() {
		if c.Space == space && c.Tag == tag {
			n++
		}
	}
	return n
}

// hasChildWithTag reports whether el has a direct child with given space:tag.
func hasChildWithTag(el *etree.Element, space, tag string) bool {
	return countChildrenByTag(el, space, tag) > 0
}

// assertFragSummary is a helper that compares fragment summaries.
func assertFragSummary(t *testing.T, frags []Fragment, want []string) {
	t.Helper()
	got := fragSummary(frags)
	if len(got) != len(want) {
		t.Fatalf("fragment count: got %d %v, want %d %v", len(got), got, len(want), want)
	}
	for i := range got {
		if got[i] != want[i] {
			t.Errorf("fragment[%d]: got %q, want %q", i, got[i], want[i])
		}
	}
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

func TestSplitParagraph_NoMatch(t *testing.T) {
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Hello world</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[<TAG>]")
	if frags != nil {
		t.Fatalf("expected nil, got %d fragments", len(frags))
	}
}

func TestSplitParagraph_EmptyOld(t *testing.T) {
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Hello</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "")
	if frags != nil {
		t.Fatalf("expected nil for empty old, got %d fragments", len(frags))
	}
}

func TestSplitParagraph_TagIsEntireText(t *testing.T) {
	// Case 1: tag is the entire paragraph text → only placeholder, no text paragraphs.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>[TAG]</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	// Expect: placeholder only (empty segments produce no paragraphs).
	assertFragSummary(t, frags, []string{"X"})
}

func TestSplitParagraph_TextBefore(t *testing.T) {
	// Case 2: text only before the tag.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Before [TAG]</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X"})
}

func TestSplitParagraph_TextAfter(t *testing.T) {
	// Case 3: text only after the tag.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>[TAG] after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"X", "P( after)"})
}

func TestSplitParagraph_TextBothSides(t *testing.T) {
	// Case 4: text before and after.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Before [TAG] after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})
}

func TestSplitParagraph_CrossRunTag(t *testing.T) {
	// Case 5: tag split across two runs.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Text [TA</w:t></w:r>
		<w:r><w:t>G] rest</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Text )", "X", "P( rest)"})
}

func TestSplitParagraph_FormattingPreserved(t *testing.T) {
	// rPr must be cloned into both split runs.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r>
			<w:rPr><w:b/></w:rPr>
			<w:t>AAA[TAG]BBB</w:t>
		</w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(AAA)", "X", "P(BBB)"})

	// Both paragraphs' runs should have bold formatting.
	if !firstRunHasRPrChild(frags[0].Element(), "w", "b") {
		t.Error("before-paragraph run missing <w:b/>")
	}
	if !firstRunHasRPrChild(frags[2].Element(), "w", "b") {
		t.Error("after-paragraph run missing <w:b/>")
	}
}

func TestSplitParagraph_ParagraphStylePreserved(t *testing.T) {
	// pPr (without sectPr) should be cloned into each paragraph fragment.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
		<w:r><w:t>Before [TAG] after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})

	// Both paragraph fragments should have pPr with pStyle.
	for _, idx := range []int{0, 2} {
		if !pPrHasChild(frags[idx].Element(), "w", "pStyle") {
			t.Errorf("fragment[%d] missing pStyle in pPr", idx)
		}
	}
}

func TestSplitParagraph_SectPrOnLastParagraph(t *testing.T) {
	// When sectPr is present and there's text after the tag,
	// sectPr should end up on the last paragraph fragment.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr>
			<w:pStyle w:val="Normal"/>
			<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>
		</w:pPr>
		<w:r><w:t>Before [TAG] after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})

	// sectPr should NOT be on the before paragraph.
	if pPrHasChild(frags[0].Element(), "w", "sectPr") {
		t.Error("before-paragraph should NOT have sectPr")
	}

	// sectPr SHOULD be on the after paragraph.
	if !pPrHasChild(frags[2].Element(), "w", "sectPr") {
		t.Error("after-paragraph should have sectPr")
	}

	// pStyle should still be on both.
	if !pPrHasChild(frags[0].Element(), "w", "pStyle") {
		t.Error("before-paragraph missing pStyle")
	}
	if !pPrHasChild(frags[2].Element(), "w", "pStyle") {
		t.Error("after-paragraph missing pStyle")
	}
}

func TestSplitParagraph_SectPrNoTextAfter(t *testing.T) {
	// Tag at end of paragraph with sectPr → empty trailing paragraph for sectPr.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:pPr>
		<w:r><w:t>Before [TAG]</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")

	// Expect: P("Before "), X, and then sectPr must be somewhere.
	// Since no text after tag, the last segment is empty → no paragraph.
	// But sectPr requires a trailing paragraph.
	// So we get: P("Before "), X, P("") ← empty paragraph with sectPr.
	if len(frags) < 3 {
		t.Fatalf("expected at least 3 fragments, got %d: %v", len(frags), fragSummary(frags))
	}

	// Last fragment must be a paragraph with sectPr.
	lastFrag := frags[len(frags)-1]
	if !lastFrag.IsParagraph() {
		t.Fatal("last fragment should be a paragraph (for sectPr)")
	}
	if !pPrHasChild(lastFrag.Element(), "w", "sectPr") {
		t.Error("last paragraph should have sectPr")
	}
}

func TestSplitParagraph_SectPrTagIsEntireText(t *testing.T) {
	// Tag is entire text AND sectPr is present → placeholder + empty paragraph with sectPr.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:pPr>
		<w:r><w:t>[TAG]</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")

	// Expect: X, P("") with sectPr.
	if len(frags) < 2 {
		t.Fatalf("expected at least 2 fragments, got %d: %v", len(frags), fragSummary(frags))
	}

	if !frags[0].IsPlaceholder() {
		t.Error("first fragment should be placeholder")
	}

	lastFrag := frags[len(frags)-1]
	if !lastFrag.IsParagraph() {
		t.Fatal("last fragment should be a paragraph (for sectPr)")
	}
	if !pPrHasChild(lastFrag.Element(), "w", "sectPr") {
		t.Error("last paragraph should have sectPr")
	}
}

func TestSplitParagraph_HyperlinkTagInside(t *testing.T) {
	// Tag is entirely inside a hyperlink.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:r><w:t>Before </w:t></w:r>
		<w:hyperlink r:id="rId1">
			<w:r><w:t>[TAG]</w:t></w:r>
		</w:hyperlink>
		<w:r><w:t> after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})
}

func TestSplitParagraph_TagSpansHyperlinkBoundary(t *testing.T) {
	// Tag starts before hyperlink and ends inside it (rare but possible).
	// "text [TA" is in a regular run, "G] link" is inside a hyperlink.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:r><w:t>text [TA</w:t></w:r>
		<w:hyperlink r:id="rId1">
			<w:r><w:t>G] link</w:t></w:r>
		</w:hyperlink>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(text )", "X", "P( link)"})

	// The " link" text should be inside a hyperlink in the after-paragraph.
	afterP := frags[2].Element()
	if !hasChildWithTag(afterP, "w", "hyperlink") {
		t.Error("after-paragraph should contain a hyperlink")
	}
}

func TestSplitParagraph_FixedAtoms(t *testing.T) {
	// Tab inside the tag should be removed.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>A[T</w:t></w:r>
		<w:r><w:tab/><w:t>AG]B</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[T\tAG]")
	assertFragSummary(t, frags, []string{"P(A)", "X", "P(B)"})
}

func TestSplitParagraph_EmptyRunsOmitted(t *testing.T) {
	// When the tag consumes all text of a run, no empty run should remain.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Before </w:t></w:r>
		<w:r><w:rPr><w:b/></w:rPr><w:t>[TAG]</w:t></w:r>
		<w:r><w:t> after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})

	// Before-paragraph should have exactly 1 run.
	beforeP := frags[0].Element()
	if n := countChildrenByTag(beforeP, "w", "r"); n != 1 {
		t.Errorf("before-paragraph: expected 1 run, got %d", n)
	}
}

func TestSplitParagraph_DrawingPositional(t *testing.T) {
	// Drawing before the tag → stays in before-paragraph.
	// Drawing after the tag → stays in after-paragraph.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r>
			<w:drawing><w:inline/></w:drawing>
			<w:t>Before [TAG] after</w:t>
		</w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})

	// The drawing should be in the before-paragraph (it precedes all text atoms).
	beforeP := frags[0].Element()
	beforeRun := findFirstRun(beforeP)
	if beforeRun == nil {
		t.Fatal("before-paragraph has no run")
	}
	if !hasChildWithTag(beforeRun, "w", "drawing") {
		t.Error("drawing should be in before-paragraph's run")
	}

	// After-paragraph should NOT have a drawing.
	afterP := frags[2].Element()
	afterRun := findFirstRun(afterP)
	if afterRun != nil && hasChildWithTag(afterRun, "w", "drawing") {
		t.Error("after-paragraph's run should NOT have drawing")
	}
}

func TestSplitParagraph_BookmarkPositional(t *testing.T) {
	// Bookmark before the run → goes to the same segment as the first run.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:bookmarkStart w:id="0" w:name="bm1"/>
		<w:r><w:t>Before [TAG] after</w:t></w:r>
		<w:bookmarkEnd w:id="0"/>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})

	// bookmarkStart should be in the before-paragraph (assigned before any run → seg 0).
	beforeP := frags[0].Element()
	if !hasChildWithTag(beforeP, "w", "bookmarkStart") {
		t.Error("bookmarkStart should be in before-paragraph")
	}

	// bookmarkEnd comes after the run → should go to after-paragraph's segment.
	afterP := frags[2].Element()
	if !hasChildWithTag(afterP, "w", "bookmarkEnd") {
		t.Error("bookmarkEnd should be in after-paragraph")
	}
}

func TestSplitParagraph_MultipleTagsSinglePass(t *testing.T) {
	// Two tags in one paragraph → 5 fragments.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>A [TAG] B [TAG] C</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(A )", "X", "P( B )", "X", "P( C)"})
}

func TestSplitParagraph_MultipleTagsNoTextBetween(t *testing.T) {
	// Two adjacent tags with no text between → middle segment is empty.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>A[TAG][TAG]B</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	// Segments: "A", "", "B". Empty middle produces no paragraph.
	assertFragSummary(t, frags, []string{"P(A)", "X", "X", "P(B)"})
}

func TestSplitParagraph_MultiRunPreservesRunOrder(t *testing.T) {
	// Multiple runs, tag doesn't cross boundary — runs preserved in order.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:rPr><w:b/></w:rPr><w:t>Bold </w:t></w:r>
		<w:r><w:t>[TAG]</w:t></w:r>
		<w:r><w:rPr><w:i/></w:rPr><w:t> italic</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Bold )", "X", "P( italic)"})

	// Before-paragraph run should be bold.
	if !firstRunHasRPrChild(frags[0].Element(), "w", "b") {
		t.Error("before-paragraph should have bold run")
	}

	// After-paragraph run should be italic.
	if !firstRunHasRPrChild(frags[2].Element(), "w", "i") {
		t.Error("after-paragraph should have italic run")
	}
}

func TestSplitParagraph_PreserveSpaceOnTrimmedText(t *testing.T) {
	// Trimmed text with leading/trailing spaces should have xml:space="preserve".
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t xml:space="preserve">Hello [TAG] world</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Hello )", "X", "P( world)"})

	// "Hello " has trailing space → should preserve.
	beforeT := findFirstT(frags[0].Element())
	if beforeT == nil {
		t.Fatal("no <w:t> in before-paragraph")
	}
	if v := etreeAttrVal(beforeT, "xml", "space"); v != "preserve" {
		t.Errorf("before <w:t> xml:space = %q, want 'preserve'", v)
	}

	// " world" has leading space → should preserve.
	afterT := findFirstT(frags[2].Element())
	if afterT == nil {
		t.Fatal("no <w:t> in after-paragraph")
	}
	if v := etreeAttrVal(afterT, "xml", "space"); v != "preserve" {
		t.Errorf("after <w:t> xml:space = %q, want 'preserve'", v)
	}
}

func TestSplitParagraph_FixedAtomOutsideTag(t *testing.T) {
	// Tab BEFORE the tag should be preserved as <w:tab/>, not turned into <w:t>.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:tab/><w:t>[TAG]after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	// Seg 0 contains only the tab ("\t"), seg 1 contains "after".
	if len(frags) < 2 {
		t.Fatalf("expected at least 2 fragments, got %d: %v", len(frags), fragSummary(frags))
	}

	// First paragraph fragment should contain a <w:tab/> (not <w:t>).
	var firstParaFrag Fragment
	for _, f := range frags {
		if f.IsParagraph() {
			firstParaFrag = f
			break
		}
	}
	r := findFirstRun(firstParaFrag.Element())
	if r == nil {
		t.Fatal("first paragraph has no run")
	}
	if !hasChildWithTag(r, "w", "tab") {
		t.Error("tab before tag should be preserved as <w:tab/>, not <w:t>")
	}
}

func TestSplitParagraph_DrawingInsideTagGapDropped(t *testing.T) {
	// Drawing between two text atoms that are both intersected by the tag
	// should be dropped (§5.1.2: children between intersected atoms → dropped).
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r>
			<w:t>A[TA</w:t>
			<w:drawing><w:inline/></w:drawing>
			<w:t>G]B</w:t>
		</w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(A)", "X", "P(B)"})

	// "A" paragraph's run should NOT have the drawing (it was inside the tag gap).
	beforeRun := findFirstRun(frags[0].Element())
	if beforeRun != nil && hasChildWithTag(beforeRun, "w", "drawing") {
		t.Error("drawing inside tag gap should be dropped, but found in before-paragraph")
	}

	// "B" paragraph's run should NOT have the drawing either.
	afterRun := findFirstRun(frags[2].Element())
	if afterRun != nil && hasChildWithTag(afterRun, "w", "drawing") {
		t.Error("drawing inside tag gap should be dropped, but found in after-paragraph")
	}
}

func TestSplitParagraph_TextlessRunPreserved(t *testing.T) {
	// A run with only a drawing (no text atoms) should not be lost.
	// It should be cloned to the same segment as the nearest preceding run.
	p := mkP(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Before </w:t></w:r>
		<w:r><w:drawing><w:inline/></w:drawing></w:r>
		<w:r><w:t>[TAG] after</w:t></w:r>
	</w:p>`)

	frags := SplitParagraphAtTags(p, "[TAG]")
	assertFragSummary(t, frags, []string{"P(Before )", "X", "P( after)"})

	// The drawing run should be in the before-paragraph (lastSeg=0 after first run).
	beforeP := frags[0].Element()
	foundDrawing := false
	for _, child := range beforeP.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			if hasChildWithTag(child, "w", "drawing") {
				foundDrawing = true
			}
		}
	}
	if !foundDrawing {
		t.Error("text-less run with drawing should be preserved in before-paragraph")
	}
}

// ---------------------------------------------------------------------------
// Additional helpers used by tests
// ---------------------------------------------------------------------------

// findFirstRun returns the first <w:r> child of a <w:p>, or nil.
func findFirstRun(p *etree.Element) *etree.Element {
	for _, c := range p.ChildElements() {
		if c.Space == "w" && c.Tag == "r" {
			return c
		}
	}
	return nil
}

// findFirstT returns the first <w:t> element under the first <w:r> in a <w:p>.
func findFirstT(p *etree.Element) *etree.Element {
	r := findFirstRun(p)
	if r == nil {
		return nil
	}
	for _, c := range r.ChildElements() {
		if c.Space == "w" && c.Tag == "t" {
			return c
		}
	}
	return nil
}
