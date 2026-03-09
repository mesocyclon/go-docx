package oxml

import (
	"testing"

	"github.com/beevik/etree"
)

func TestCT_R_AddTWithText_PreservesSpace(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	t1 := r.AddTWithText(" hello ")
	// Check xml:space="preserve" is set
	val := t1.e.SelectAttrValue("xml:space", "")
	if val != "preserve" {
		t.Errorf("expected xml:space=preserve for text with spaces, got %q", val)
	}

	r2El := OxmlElement("w:r")
	r2 := &CT_R{Element{e: r2El}}
	t2 := r2.AddTWithText("hello")
	val2 := t2.e.SelectAttrValue("xml:space", "")
	if val2 != "" {
		t.Errorf("expected no xml:space for trimmed text, got %q", val2)
	}
}

func TestCT_R_RunText(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.AddTWithText("Hello")
	r.AddTab()
	r.AddTWithText("World")

	got := r.RunText()
	if got != "Hello\tWorld" {
		t.Errorf("RunText() = %q, want %q", got, "Hello\tWorld")
	}
}

func TestCT_R_RunTextWithBr(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.AddTWithText("Line1")
	r.AddBr() // default type = textWrapping → "\n"
	r.AddTWithText("Line2")

	got := r.RunText()
	if got != "Line1\nLine2" {
		t.Errorf("RunText() = %q, want %q", got, "Line1\nLine2")
	}
}

func TestCT_R_ClearContent(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.GetOrAddRPr()
	r.AddTWithText("text")
	r.AddBr()

	r.ClearContent()

	if r.RPr() == nil {
		t.Error("rPr should be preserved after ClearContent")
	}
	if len(r.TList()) != 0 || len(r.BrList()) != 0 {
		t.Error("content should be removed after ClearContent")
	}
}

func TestCT_R_Style_RoundTrip(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	if s, err := r.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style for new run")
	}

	s := "Emphasis"
	if err := r.SetStyle(&s); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	got, err := r.Style()
	if err != nil {
		t.Fatalf("Style: %v", err)
	}
	if got == nil || *got != "Emphasis" {
		t.Errorf("expected Emphasis style, got %v", got)
	}

	if err := r.SetStyle(nil); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	if s, err := r.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style after removing")
	}
}

func TestCT_R_SetRunText(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.GetOrAddRPr() // should be preserved
	r.SetRunText("Hello\tWorld\nNew")

	// Check rPr still exists
	if r.RPr() == nil {
		t.Error("rPr should be preserved after SetRunText")
	}

	got := r.RunText()
	if got != "Hello\tWorld\nNew" {
		t.Errorf("after SetRunText, RunText() = %q, want %q", got, "Hello\tWorld\nNew")
	}
}

// --- CT_Br tests ---

func TestCT_Br_TextEquivalent(t *testing.T) {
	// Default (textWrapping)
	br1 := &CT_Br{Element{e: OxmlElement("w:br")}}
	if br1.TextEquivalent() != "\n" {
		t.Error("expected newline for default break type")
	}

	// Page break
	br2 := &CT_Br{Element{e: OxmlElement("w:br")}}
	if err := br2.SetType("page"); err != nil {
		t.Fatalf("SetType: %v", err)
	}
	if br2.TextEquivalent() != "" {
		t.Error("expected empty string for page break")
	}
}

func TestAddDrawingWithInline(t *testing.T) {
	t.Parallel()
	r := &CT_R{Element{e: OxmlElement("w:r")}}

	inlineEl := OxmlElement("wp:inline")
	inline := &CT_Inline{Element{e: inlineEl}}

	drawing := r.AddDrawingWithInline(inline)
	if drawing == nil {
		t.Fatal("AddDrawingWithInline returned nil")
	}
	// Drawing should contain the inline as child
	children := drawing.e.ChildElements()
	found := false
	for _, c := range children {
		if c == inlineEl {
			found = true
		}
	}
	if !found {
		t.Error("drawing should contain the inline element")
	}
}

func TestLastRenderedPageBreaks(t *testing.T) {
	t.Parallel()
	r := &CT_R{Element{e: OxmlElement("w:r")}}
	// Add some content with lastRenderedPageBreak
	lrpb := r.e.CreateElement("lastRenderedPageBreak")
	lrpb.Space = "w"
	lrpb.Tag = "lastRenderedPageBreak"

	breaks := r.LastRenderedPageBreaks()
	if len(breaks) != 1 {
		t.Errorf("expected 1 lastRenderedPageBreak, got %d", len(breaks))
	}
}

func TestLastRenderedPageBreaks_Empty(t *testing.T) {
	t.Parallel()
	r := &CT_R{Element{e: OxmlElement("w:r")}}
	breaks := r.LastRenderedPageBreaks()
	if len(breaks) != 0 {
		t.Errorf("expected 0 breaks in empty run, got %d", len(breaks))
	}
}

func TestCT_Cr_TextEquivalent(t *testing.T) {
	t.Parallel()
	cr := &CT_Cr{Element{e: OxmlElement("w:cr")}}
	if cr.TextEquivalent() != "\n" {
		t.Errorf("CT_Cr.TextEquivalent() = %q, want \\n", cr.TextEquivalent())
	}
}

func TestCT_NoBreakHyphen_TextEquivalent(t *testing.T) {
	t.Parallel()
	nbh := &CT_NoBreakHyphen{Element{e: OxmlElement("w:noBreakHyphen")}}
	if nbh.TextEquivalent() != "-" {
		t.Errorf("CT_NoBreakHyphen.TextEquivalent() = %q, want -", nbh.TextEquivalent())
	}
}

func TestCT_PTab_TextEquivalent(t *testing.T) {
	t.Parallel()
	pt := &CT_PTab{Element{e: OxmlElement("w:ptab")}}
	if pt.TextEquivalent() != "\t" {
		t.Errorf("CT_PTab.TextEquivalent() = %q, want \\t", pt.TextEquivalent())
	}
}

func TestCT_Br_TextEquivalent_PageBreak(t *testing.T) {
	t.Parallel()
	// page break → ""
	br := &CT_Br{Element{e: OxmlElement("w:br")}}
	br.e.CreateAttr("w:type", "page")
	if br.TextEquivalent() != "" {
		t.Errorf("CT_Br.TextEquivalent(page) = %q, want empty", br.TextEquivalent())
	}
}

func TestCT_Text_ContentText(t *testing.T) {
	t.Parallel()
	te := OxmlElement("w:t")
	te.SetText("Hello World")
	ct := &CT_Text{Element{e: te}}
	if ct.ContentText() != "Hello World" {
		t.Errorf("ContentText() = %q, want Hello World", ct.ContentText())
	}
}

func TestCT_Text_SetPreserveSpace(t *testing.T) {
	t.Parallel()
	te := OxmlElement("w:t")
	ct := &CT_Text{Element{e: te}}
	ct.SetPreserveSpace()
	val := te.SelectAttrValue("xml:space", "")
	if val != "preserve" {
		t.Errorf("xml:space = %q, want preserve", val)
	}
}

func TestInnerContentItems(t *testing.T) {
	t.Parallel()
	r := &CT_R{Element{e: OxmlElement("w:r")}}

	// Add mixed content: text, drawing, lastRenderedPageBreak, tab, cr, noBreakHyphen
	te := r.e.CreateElement("t")
	te.Space = "w"
	te.SetText("Hello")

	drawing := r.e.CreateElement("drawing")
	drawing.Space = "w"

	te2 := r.e.CreateElement("t")
	te2.Space = "w"
	te2.SetText(" World")

	lrpb := r.e.CreateElement("lastRenderedPageBreak")
	lrpb.Space = "w"

	tab := r.e.CreateElement("tab")
	tab.Space = "w"

	cr := r.e.CreateElement("cr")
	cr.Space = "w"

	nbh := r.e.CreateElement("noBreakHyphen")
	nbh.Space = "w"

	ptab := r.e.CreateElement("ptab")
	ptab.Space = "w"

	items := r.InnerContentItems()
	// Expected: "Hello", *CT_Drawing, " World", *CT_LastRenderedPageBreak, "\t\n-\t"
	if len(items) != 5 {
		t.Fatalf("expected 5 items, got %d", len(items))
	}
	if s, ok := items[0].(string); !ok || s != "Hello" {
		t.Errorf("item[0] = %v, want Hello", items[0])
	}
	if _, ok := items[1].(*CT_Drawing); !ok {
		t.Errorf("item[1] should be *CT_Drawing, got %T", items[1])
	}
	if s, ok := items[2].(string); !ok || s != " World" {
		t.Errorf("item[2] = %v, want ' World'", items[2])
	}
	if _, ok := items[3].(*CT_LastRenderedPageBreak); !ok {
		t.Errorf("item[3] should be *CT_LastRenderedPageBreak, got %T", items[3])
	}
	if s, ok := items[4].(string); !ok || s != "\t\n-\t" {
		t.Errorf("item[4] = %v, want '\\t\\n-\\t'", items[4])
	}
}

func TestInnerContentItemsIgnoreNonWNamespace(t *testing.T) {
	t.Parallel()
	r := &CT_R{Element{e: OxmlElement("w:r")}}
	// Non-w namespace child should be ignored
	foreign := r.e.CreateElement("foo")
	foreign.Space = "mc"
	foreign.SetText("ignored")

	items := r.InnerContentItems()
	if len(items) != 0 {
		t.Errorf("expected 0 items for non-w children, got %d", len(items))
	}
}

func TestInsertCommentRangeStartAbove(t *testing.T) {
	t.Parallel()
	p := OxmlElement("w:p")
	r1 := p.CreateElement("r")
	r1.Space = "w"
	r2 := p.CreateElement("r")
	r2.Space = "w"

	run := &CT_R{Element{e: r1}}
	run.InsertCommentRangeStartAbove(42)

	children := p.ChildElements()
	if len(children) != 3 {
		t.Fatalf("expected 3 children, got %d", len(children))
	}
	if children[0].Tag != "commentRangeStart" {
		t.Errorf("first child should be commentRangeStart, got %q", children[0].Tag)
	}
	id := children[0].SelectAttrValue("w:id", "")
	if id != "42" {
		t.Errorf("commentRangeStart id = %q, want 42", id)
	}
}

func TestInsertCommentRangeStartAbove_NoParent(t *testing.T) {
	t.Parallel()
	// Detached run — should not panic
	r := &CT_R{Element{e: OxmlElement("w:r")}}
	r.InsertCommentRangeStartAbove(1) // no-op, no parent
}

func TestInsertCommentRangeEndAndReferenceBelow(t *testing.T) {
	t.Parallel()
	p := OxmlElement("w:p")
	r1 := p.CreateElement("r")
	r1.Space = "w"

	run := &CT_R{Element{e: r1}}
	run.InsertCommentRangeEndAndReferenceBelow(7)

	children := p.ChildElements()
	if len(children) != 3 {
		t.Fatalf("expected 3 children, got %d", len(children))
	}
	// children[0] = original run
	// children[1] = commentRangeEnd
	// children[2] = reference run
	if children[1].Tag != "commentRangeEnd" {
		t.Errorf("second child = %q, want commentRangeEnd", children[1].Tag)
	}
	if children[1].SelectAttrValue("w:id", "") != "7" {
		t.Error("commentRangeEnd should have id=7")
	}
	if children[2].Tag != "r" {
		t.Errorf("third child = %q, want r", children[2].Tag)
	}
	// Check reference run has rStyle and commentReference
	rPr := children[2].FindElement("w:rPr/w:rStyle")
	if rPr == nil {
		t.Fatal("reference run should have rPr/rStyle")
	}
	if rPr.SelectAttrValue("w:val", "") != "CommentReference" {
		t.Error("rStyle val should be CommentReference")
	}
	cr := children[2].FindElement("w:commentReference")
	if cr == nil {
		t.Fatal("reference run should have commentReference")
	}
	if cr.SelectAttrValue("w:id", "") != "7" {
		t.Error("commentReference should have id=7")
	}
}

func TestInsertCommentRangeEndAndReferenceBelow_NoParent(t *testing.T) {
	t.Parallel()
	r := &CT_R{Element{e: OxmlElement("w:r")}}
	r.InsertCommentRangeEndAndReferenceBelow(1) // no-op
}

func TestChildIndex(t *testing.T) {
	t.Parallel()
	parent := etree.NewElement("p")
	c1 := parent.CreateElement("a")
	c2 := parent.CreateElement("b")

	if idx := childIndex(parent, c1); idx != 0 {
		t.Errorf("childIndex(c1) = %d, want 0", idx)
	}
	if idx := childIndex(parent, c2); idx != 1 {
		t.Errorf("childIndex(c2) = %d, want 1", idx)
	}

	orphan := etree.NewElement("orphan")
	if idx := childIndex(parent, orphan); idx != -1 {
		t.Errorf("childIndex(orphan) = %d, want -1", idx)
	}
}
