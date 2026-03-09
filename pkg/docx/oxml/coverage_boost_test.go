package oxml

import (
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ===========================================================================
// babelfish.go — UI2Internal / Internal2UI
// ===========================================================================

func TestUI2Internal(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input, want string
	}{
		{"Heading 1", "heading 1"},
		{"Heading 9", "heading 9"},
		{"Caption", "caption"},
		{"Header", "header"},
		{"Footer", "footer"},
		{"Normal", "Normal"},           // unmapped — returned as-is
		{"My Custom Style", "My Custom Style"}, // unmapped
	}
	for _, tt := range tests {
		if got := UI2Internal(tt.input); got != tt.want {
			t.Errorf("UI2Internal(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestInternal2UI(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input, want string
	}{
		{"heading 1", "Heading 1"},
		{"caption", "Caption"},
		{"header", "Header"},
		{"footer", "Footer"},
		{"Normal", "Normal"}, // unmapped
	}
	for _, tt := range tests {
		if got := Internal2UI(tt.input); got != tt.want {
			t.Errorf("Internal2UI(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

// ===========================================================================
// ns.go — LookupNsURI / LookupPrefix
// ===========================================================================

func TestLookupNsURI(t *testing.T) {
	t.Parallel()

	t.Run("known prefix", func(t *testing.T) {
		uri, ok := LookupNsURI("w")
		if !ok || uri != NsWml {
			t.Errorf("LookupNsURI(w) = (%q, %v), want (%q, true)", uri, ok, NsWml)
		}
	})
	t.Run("unknown prefix", func(t *testing.T) {
		_, ok := LookupNsURI("zzz")
		if ok {
			t.Error("LookupNsURI(zzz) should return false")
		}
	})
}

func TestLookupPrefix(t *testing.T) {
	t.Parallel()

	t.Run("known URI", func(t *testing.T) {
		pfx, ok := LookupPrefix(NsWml)
		if !ok || pfx != "w" {
			t.Errorf("LookupPrefix(NsWml) = (%q, %v), want (w, true)", pfx, ok)
		}
	})
	t.Run("unknown URI", func(t *testing.T) {
		_, ok := LookupPrefix("http://example.com/unknown")
		if ok {
			t.Error("LookupPrefix(unknown) should return false")
		}
	})
}

// ===========================================================================
// element.go — WrapElement, XPath, resolveAttrName (Clark notation branch)
// ===========================================================================

func TestWrapElement(t *testing.T) {
	t.Parallel()
	raw := etree.NewElement("p")
	raw.Space = "w"
	el := WrapElement(raw)
	if el.RawElement() != raw {
		t.Error("WrapElement should wrap the given etree element")
	}
}

func TestElementXPath(t *testing.T) {
	t.Parallel()
	p := etree.NewElement("p")
	p.Space = "w"
	r := p.CreateElement("r")
	r.Space = "w"
	el := &Element{e: p}

	results := el.XPath("r")
	if len(results) != 1 {
		t.Errorf("XPath(r) returned %d results, want 1", len(results))
	}
}

func TestResolveAttrNameClarkNotation(t *testing.T) {
	t.Parallel()

	// Clark-notation attribute access via GetAttr/SetAttr
	e := etree.NewElement("p")
	e.Space = "w"
	el := &Element{e: e}

	// Set with Clark notation → should resolve to prefix form
	el.SetAttr("w:val", "test")
	val, ok := el.GetAttr("{"+NsWml+"}val")
	if !ok || val != "test" {
		t.Errorf("GetAttr with Clark notation = (%q, %v), want (test, true)", val, ok)
	}
}

func TestResolveAttrNameUnknownClark(t *testing.T) {
	t.Parallel()
	e := etree.NewElement("p")
	e.Space = "w"
	e.CreateAttr("val", "hello")
	el := &Element{e: e}

	// Unknown Clark notation URI — should still resolve to local part
	val, ok := el.GetAttr("{http://unknown.example.com}val")
	if !ok || val != "hello" {
		t.Errorf("GetAttr with unknown Clark = (%q, %v), want (hello, true)", val, ok)
	}
}

func TestGetAttrNamespacedMismatch(t *testing.T) {
	t.Parallel()
	e := etree.NewElement("p")
	e.Space = "w"
	// Attr with specific namespace
	e.CreateAttr("r:id", "rId1")
	el := &Element{e: e}

	// Request with different namespace — should not match
	_, ok := el.GetAttr("w:id")
	if ok {
		t.Error("GetAttr should not match attr from different namespace")
	}

	// Request with correct namespace
	val, ok := el.GetAttr("r:id")
	if !ok || val != "rId1" {
		t.Errorf("GetAttr(r:id) = (%q, %v), want (rId1, true)", val, ok)
	}
}

func TestSetAttrWithNamespace(t *testing.T) {
	t.Parallel()
	e := etree.NewElement("p")
	el := &Element{e: e}
	el.SetAttr("w:val", "hello")
	val, ok := el.GetAttr("w:val")
	if !ok || val != "hello" {
		t.Errorf("SetAttr(w:val) round-trip failed: (%q, %v)", val, ok)
	}
}

func TestRemoveAttrWithNamespace(t *testing.T) {
	t.Parallel()
	e := etree.NewElement("p")
	el := &Element{e: e}
	el.SetAttr("w:val", "hello")
	el.RemoveAttr("w:val")
	_, ok := el.GetAttr("w:val")
	if ok {
		t.Error("RemoveAttr(w:val) should have removed the attribute")
	}
}

func TestClarkFromEtreeNoSpace(t *testing.T) {
	t.Parallel()
	e := etree.NewElement("custom")
	// No Space set — should just return tag
	tag := clarkFromEtree(e)
	if tag != "custom" {
		t.Errorf("clarkFromEtree with no space = %q, want custom", tag)
	}
}

func TestClarkFromEtreeUnknownSpace(t *testing.T) {
	t.Parallel()
	e := etree.NewElement("foo")
	e.Space = "zzz" // not in nsmap
	tag := clarkFromEtree(e)
	// Should use Space as URI directly
	if tag != "{zzz}foo" {
		t.Errorf("clarkFromEtree with unknown space = %q, want {zzz}foo", tag)
	}
}

func TestInsertBeforeRefNotFound(t *testing.T) {
	t.Parallel()
	parent := etree.NewElement("body")
	newChild := etree.NewElement("p")
	refChild := etree.NewElement("missing") // not in parent

	insertBefore(parent, newChild, refChild)
	// Should append
	children := parent.ChildElements()
	if len(children) != 1 || children[0] != newChild {
		t.Error("insertBefore should append when refChild not found")
	}
}

// ===========================================================================
// styles_custom.go — GetStyleIDByName, DocDefaultsRPr, DocDefaultsPPr
// ===========================================================================

func TestGetStyleIDByName(t *testing.T) {
	t.Parallel()
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}

	// Add Normal as default paragraph style
	normal := styles.AddStyle()
	xmlType, _ := enum.WdStyleTypeParagraph.ToXml()
	_ = normal.SetStyleId("Normal")
	_ = normal.SetNameVal("Normal")
	_ = normal.SetType(xmlType)
	_ = normal.SetDefault(true)

	// Add Heading 1 as non-default
	h1 := styles.AddStyle()
	_ = h1.SetStyleId("Heading1")
	_ = h1.SetNameVal("heading 1")
	_ = h1.SetType(xmlType)

	t.Run("returns nil for default style", func(t *testing.T) {
		t.Parallel()
		result, err := styles.GetStyleIDByName("Normal", enum.WdStyleTypeParagraph)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if result != nil {
			t.Errorf("expected nil for default style, got %q", *result)
		}
	})

	t.Run("returns style ID for non-default", func(t *testing.T) {
		t.Parallel()
		// "Heading 1" → BabelFish → "heading 1" → found by name
		result, err := styles.GetStyleIDByName("Heading 1", enum.WdStyleTypeParagraph)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if result == nil || *result != "Heading1" {
			t.Errorf("expected Heading1, got %v", result)
		}
	})

	t.Run("fallback to ID lookup", func(t *testing.T) {
		t.Parallel()
		// "Heading1" is not a name — but it is an ID
		result, err := styles.GetStyleIDByName("Heading1", enum.WdStyleTypeParagraph)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if result == nil || *result != "Heading1" {
			t.Errorf("expected Heading1, got %v", result)
		}
	})

	t.Run("not found returns error", func(t *testing.T) {
		t.Parallel()
		_, err := styles.GetStyleIDByName("NoSuchStyle", enum.WdStyleTypeParagraph)
		if err == nil {
			t.Error("expected error for unknown style")
		}
	})

	t.Run("type mismatch returns error", func(t *testing.T) {
		t.Parallel()
		_, err := styles.GetStyleIDByName("Heading 1", enum.WdStyleTypeCharacter)
		if err == nil {
			t.Error("expected error for type mismatch")
		}
	})
}

func TestDocDefaultsRPr(t *testing.T) {
	t.Parallel()
	stylesXML := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:docDefaults>
			<w:rPrDefault>
				<w:rPr><w:sz w:val="24"/></w:rPr>
			</w:rPrDefault>
		</w:docDefaults>
	</w:styles>`
	doc := etree.NewDocument()
	if err := doc.ReadFromString(stylesXML); err != nil {
		t.Fatalf("parse: %v", err)
	}
	styles := &CT_Styles{Element{e: doc.Root()}}

	rPr := styles.DocDefaultsRPr()
	if rPr == nil {
		t.Fatal("expected non-nil DocDefaultsRPr")
	}
	sz := rPr.FindElement("w:sz")
	if sz == nil {
		t.Fatal("expected w:sz child")
	}
}

func TestDocDefaultsPPr(t *testing.T) {
	t.Parallel()
	stylesXML := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:docDefaults>
			<w:pPrDefault>
				<w:pPr><w:spacing w:after="200"/></w:pPr>
			</w:pPrDefault>
		</w:docDefaults>
	</w:styles>`
	doc := etree.NewDocument()
	if err := doc.ReadFromString(stylesXML); err != nil {
		t.Fatalf("parse: %v", err)
	}
	styles := &CT_Styles{Element{e: doc.Root()}}

	pPr := styles.DocDefaultsPPr()
	if pPr == nil {
		t.Fatal("expected non-nil DocDefaultsPPr")
	}
}

func TestDocDefaultsRPr_Nil(t *testing.T) {
	t.Parallel()
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	if styles.DocDefaultsRPr() != nil {
		t.Error("expected nil when no docDefaults")
	}
	if styles.DocDefaultsPPr() != nil {
		t.Error("expected nil when no docDefaults")
	}
}

// ===========================================================================
// text_run_custom.go — AddDrawingWithInline, InnerContentItems, ContentText,
// SetPreserveSpace, InsertCommentRangeStartAbove, InsertCommentRangeEndAndReferenceBelow,
// childIndex, LastRenderedPageBreaks, TextEquivalent methods
// ===========================================================================

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

// ===========================================================================
// text_parfmt_custom.go — SpacingLine, SetSpacingLineRule, IndRight, KeepNext, WidowControl
// ===========================================================================

func newPPr() *CT_PPr {
	return &CT_PPr{Element{e: OxmlElement("w:pPr")}}
}

func TestSpacingLine(t *testing.T) {
	t.Parallel()
	pPr := newPPr()

	// nil when no spacing
	v, err := pPr.SpacingLine()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("expected nil when no spacing")
	}

	// Set and read back
	val := 360
	if err := pPr.SetSpacingLine(&val); err != nil {
		t.Fatal(err)
	}
	got, err := pPr.SpacingLine()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 360 {
		t.Errorf("SpacingLine = %v, want 360", got)
	}
}

func TestSpacingLineRule(t *testing.T) {
	t.Parallel()
	pPr := newPPr()

	// nil when no spacing
	lr, err := pPr.SpacingLineRule()
	if err != nil {
		t.Fatal(err)
	}
	if lr != nil {
		t.Error("expected nil")
	}

	// Set line only → defaults to MULTIPLE
	val := 240
	if err := pPr.SetSpacingLine(&val); err != nil {
		t.Fatal(err)
	}
	lr, err = pPr.SpacingLineRule()
	if err != nil {
		t.Fatal(err)
	}
	if lr == nil || *lr != enum.WdLineSpacingMultiple {
		t.Errorf("SpacingLineRule with line but no rule = %v, want MULTIPLE", lr)
	}

	// Set explicit lineRule
	rule := enum.WdLineSpacingExactly
	if err := pPr.SetSpacingLineRule(&rule); err != nil {
		t.Fatal(err)
	}
	lr, err = pPr.SpacingLineRule()
	if err != nil {
		t.Fatal(err)
	}
	if lr == nil || *lr != enum.WdLineSpacingExactly {
		t.Errorf("SpacingLineRule = %v, want EXACTLY", lr)
	}

	// nil lineRule → removes
	if err := pPr.SetSpacingLineRule(nil); err != nil {
		t.Fatal(err)
	}
}

func TestSetSpacingLineRule_NilNoSpacing(t *testing.T) {
	t.Parallel()
	pPr := newPPr()
	// Setting nil when no spacing element — should be no-op
	if err := pPr.SetSpacingLineRule(nil); err != nil {
		t.Fatal(err)
	}
}

func TestIndRight(t *testing.T) {
	t.Parallel()
	pPr := newPPr()

	v, err := pPr.IndRight()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("expected nil")
	}

	val := 720
	if err := pPr.SetIndRight(&val); err != nil {
		t.Fatal(err)
	}
	got, err := pPr.IndRight()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 720 {
		t.Errorf("IndRight = %v, want 720", got)
	}

	// Clear
	if err := pPr.SetIndRight(nil); err != nil {
		t.Fatal(err)
	}
}

func TestSetIndRight_NilNoInd(t *testing.T) {
	t.Parallel()
	pPr := newPPr()
	if err := pPr.SetIndRight(nil); err != nil {
		t.Fatal(err)
	}
}

func TestKeepNextVal(t *testing.T) {
	t.Parallel()
	pPr := newPPr()

	if pPr.KeepNextVal() != nil {
		t.Error("expected nil by default")
	}

	tr := true
	if err := pPr.SetKeepNextVal(&tr); err != nil {
		t.Fatal(err)
	}
	v := pPr.KeepNextVal()
	if v == nil || !*v {
		t.Error("expected true")
	}

	if err := pPr.SetKeepNextVal(nil); err != nil {
		t.Fatal(err)
	}
	if pPr.KeepNextVal() != nil {
		t.Error("expected nil after clearing")
	}
}

func TestWidowControlVal(t *testing.T) {
	t.Parallel()
	pPr := newPPr()

	if pPr.WidowControlVal() != nil {
		t.Error("expected nil by default")
	}

	tr := true
	if err := pPr.SetWidowControlVal(&tr); err != nil {
		t.Fatal(err)
	}
	v := pPr.WidowControlVal()
	if v == nil || !*v {
		t.Error("expected true")
	}

	if err := pPr.SetWidowControlVal(nil); err != nil {
		t.Fatal(err)
	}
	if pPr.WidowControlVal() != nil {
		t.Error("expected nil after clearing")
	}
}

func TestSetPageBreakBeforeVal_SetFalse(t *testing.T) {
	t.Parallel()
	pPr := newPPr()

	f := false
	if err := pPr.SetPageBreakBeforeVal(&f); err != nil {
		t.Fatal(err)
	}
	// Should have element with val=false
	v := pPr.PageBreakBeforeVal()
	if v == nil || *v {
		t.Error("expected false")
	}
}

// ===========================================================================
// text_font_custom.go — ComplexScript, CsBold, CsItalic, Rtl
// ===========================================================================

func newRPr() *CT_RPr {
	return &CT_RPr{Element{e: OxmlElement("w:rPr")}}
}

func TestComplexScriptVal(t *testing.T) {
	t.Parallel()
	rPr := newRPr()

	if rPr.ComplexScriptVal() != nil {
		t.Error("expected nil by default")
	}

	tr := true
	if err := rPr.SetComplexScriptVal(&tr); err != nil {
		t.Fatal(err)
	}
	v := rPr.ComplexScriptVal()
	if v == nil || !*v {
		t.Error("expected true")
	}

	if err := rPr.SetComplexScriptVal(nil); err != nil {
		t.Fatal(err)
	}
	if rPr.ComplexScriptVal() != nil {
		t.Error("expected nil after clearing")
	}
}

func TestCsBoldVal(t *testing.T) {
	t.Parallel()
	rPr := newRPr()

	if rPr.CsBoldVal() != nil {
		t.Error("expected nil")
	}

	tr := true
	if err := rPr.SetCsBoldVal(&tr); err != nil {
		t.Fatal(err)
	}
	if v := rPr.CsBoldVal(); v == nil || !*v {
		t.Error("expected true")
	}

	if err := rPr.SetCsBoldVal(nil); err != nil {
		t.Fatal(err)
	}
	if rPr.CsBoldVal() != nil {
		t.Error("expected nil")
	}
}

func TestCsItalicVal(t *testing.T) {
	t.Parallel()
	rPr := newRPr()

	tr := true
	if err := rPr.SetCsItalicVal(&tr); err != nil {
		t.Fatal(err)
	}
	if v := rPr.CsItalicVal(); v == nil || !*v {
		t.Error("expected true")
	}

	if err := rPr.SetCsItalicVal(nil); err != nil {
		t.Fatal(err)
	}
	if rPr.CsItalicVal() != nil {
		t.Error("expected nil")
	}
}

func TestRtlVal(t *testing.T) {
	t.Parallel()
	rPr := newRPr()

	tr := true
	if err := rPr.SetRtlVal(&tr); err != nil {
		t.Fatal(err)
	}
	if v := rPr.RtlVal(); v == nil || !*v {
		t.Error("expected true")
	}

	f := false
	if err := rPr.SetRtlVal(&f); err != nil {
		t.Fatal(err)
	}
	if v := rPr.RtlVal(); v == nil || *v {
		t.Error("expected false")
	}

	if err := rPr.SetRtlVal(nil); err != nil {
		t.Fatal(err)
	}
	if rPr.RtlVal() != nil {
		t.Error("expected nil")
	}
}

// ===========================================================================
// document_custom.go — SetSectPr
// ===========================================================================

func TestBodySetSectPr(t *testing.T) {
	t.Parallel()
	bodyEl := OxmlElement("w:body")
	body := &CT_Body{Element{e: bodyEl}}

	sectPrEl := OxmlElement("w:sectPr")
	sectPr := &CT_SectPr{Element{e: sectPrEl}}
	sectPr.e.CreateAttr("w:rsidR", "test123")

	body.SetSectPr(sectPr)

	// Verify sectPr is present
	got := body.SectPr()
	if got == nil {
		t.Fatal("expected sectPr after SetSectPr")
	}

	// Replace with a new one
	sectPr2El := OxmlElement("w:sectPr")
	sectPr2 := &CT_SectPr{Element{e: sectPr2El}}
	sectPr2.e.CreateAttr("w:rsidR", "new456")

	body.SetSectPr(sectPr2)
	got2 := body.SectPr()
	if got2 == nil {
		t.Fatal("expected new sectPr")
	}
}

// ===========================================================================
// numbering_custom.go — AddLvlOverrideWithIlvl, AddStartOverrideWithVal
// ===========================================================================

func TestAddLvlOverrideWithIlvl(t *testing.T) {
	t.Parallel()
	numEl := OxmlElement("w:num")
	numEl.CreateAttr("w:numId", "1")
	num := &CT_Num{Element{e: numEl}}

	lvl, err := num.AddLvlOverrideWithIlvl(3)
	if err != nil {
		t.Fatal(err)
	}
	ilvl, err := lvl.Ilvl()
	if err != nil {
		t.Fatalf("Ilvl: %v", err)
	}
	if ilvl != 3 {
		t.Errorf("ilvl = %d, want 3", ilvl)
	}
}

func TestAddStartOverrideWithVal(t *testing.T) {
	t.Parallel()
	numLvlEl := OxmlElement("w:lvlOverride")
	nl := &CT_NumLvl{Element{e: numLvlEl}}

	so, err := nl.AddStartOverrideWithVal(5)
	if err != nil {
		t.Fatal(err)
	}
	val, err := so.Val()
	if err != nil {
		t.Fatalf("Val: %v", err)
	}
	if val != 5 {
		t.Errorf("startOverride val = %d, want 5", val)
	}
}

// ===========================================================================
// coreprops_custom.go — IdentifierText, SetIdentifierText
// ===========================================================================

func TestIdentifierText(t *testing.T) {
	t.Parallel()
	cp, err := NewCoreProperties()
	if err != nil {
		t.Fatal(err)
	}

	if cp.IdentifierText() != "" {
		t.Error("expected empty string by default")
	}

	if err := cp.SetIdentifierText("urn:isbn:123"); err != nil {
		t.Fatal(err)
	}
	if got := cp.IdentifierText(); got != "urn:isbn:123" {
		t.Errorf("IdentifierText() = %q, want urn:isbn:123", got)
	}
}

// ===========================================================================
// content_items.go — BlockItem / InlineItem marker methods
// ===========================================================================

func TestBlockItemInterface(t *testing.T) {
	t.Parallel()
	var bi BlockItem

	bi = &CT_P{Element{e: OxmlElement("w:p")}}
	bi.isBlockItem() // compile-time + runtime verification

	bi = &CT_Tbl{Element{e: OxmlElement("w:tbl")}}
	bi.isBlockItem()
	_ = bi
}

func TestInlineItemInterface(t *testing.T) {
	t.Parallel()
	var ii InlineItem

	ii = &CT_R{Element{e: OxmlElement("w:r")}}
	ii.isInlineItem()

	ii = &CT_Hyperlink{Element{e: OxmlElement("w:hyperlink")}}
	ii.isInlineItem()
	_ = ii
}

// ===========================================================================
// text_paragraph_custom.go — LastRenderedPageBreaks
// ===========================================================================

func TestParagraphLastRenderedPageBreaks(t *testing.T) {
	t.Parallel()
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	// Add a run with a lastRenderedPageBreak
	rEl := pEl.CreateElement("r")
	rEl.Space = "w"
	lrpb := rEl.CreateElement("lastRenderedPageBreak")
	lrpb.Space = "w"

	breaks := p.LastRenderedPageBreaks()
	if len(breaks) != 1 {
		t.Errorf("expected 1 break, got %d", len(breaks))
	}
}

func TestParagraphLastRenderedPageBreaks_InHyperlink(t *testing.T) {
	t.Parallel()
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	// Add a hyperlink with a run containing lastRenderedPageBreak
	hlEl := pEl.CreateElement("hyperlink")
	hlEl.Space = "w"
	rEl := hlEl.CreateElement("r")
	rEl.Space = "w"
	lrpb := rEl.CreateElement("lastRenderedPageBreak")
	lrpb.Space = "w"

	breaks := p.LastRenderedPageBreaks()
	if len(breaks) != 1 {
		t.Errorf("expected 1 break in hyperlink, got %d", len(breaks))
	}
}

// ===========================================================================
// text_hyperlink_custom.go — HyperlinkLastRenderedPageBreaks
// ===========================================================================

func TestHyperlinkLastRenderedPageBreaks(t *testing.T) {
	t.Parallel()
	hlEl := OxmlElement("w:hyperlink")
	hl := &CT_Hyperlink{Element{e: hlEl}}

	rEl := hlEl.CreateElement("r")
	rEl.Space = "w"
	lrpb := rEl.CreateElement("lastRenderedPageBreak")
	lrpb.Space = "w"

	breaks := hl.HyperlinkLastRenderedPageBreaks()
	if len(breaks) != 1 {
		t.Errorf("expected 1 break, got %d", len(breaks))
	}
}

func TestHyperlinkLastRenderedPageBreaks_Empty(t *testing.T) {
	t.Parallel()
	hlEl := OxmlElement("w:hyperlink")
	hl := &CT_Hyperlink{Element{e: hlEl}}

	breaks := hl.HyperlinkLastRenderedPageBreaks()
	if len(breaks) != 0 {
		t.Errorf("expected 0 breaks, got %d", len(breaks))
	}
}

// ===========================================================================
// text_pagebreak_custom.go — isRunInnerContent
// ===========================================================================

func TestIsRunInnerContent(t *testing.T) {
	t.Parallel()
	// These are the tags recognized by isRunInnerContent
	runContent := []string{"t", "br", "cr", "tab", "noBreakHyphen", "drawing", "ptab"}

	for _, tag := range runContent {
		e := etree.NewElement(tag)
		e.Space = "w"
		if !isRunInnerContent(e) {
			t.Errorf("isRunInnerContent(%q) should be true", tag)
		}
	}

	// Non-content elements
	nonContent := []string{"rPr", "lastRenderedPageBreak", "fldChar", "instrText"}
	for _, tag := range nonContent {
		e := etree.NewElement(tag)
		e.Space = "w"
		if isRunInnerContent(e) {
			t.Errorf("isRunInnerContent(%q) should be false", tag)
		}
	}

	// Non-w namespace
	foreign := etree.NewElement("t")
	foreign.Space = "mc"
	if isRunInnerContent(foreign) {
		t.Error("isRunInnerContent(mc:t) should be false")
	}
}

// ===========================================================================
// Additional styles_custom.go coverage — AddStyleOfType builtin flag,
// SetBasedOnVal empty clear
// ===========================================================================

func TestSetBasedOnVal_Clear(t *testing.T) {
	t.Parallel()
	styles := &CT_Styles{Element{e: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	_ = s.SetBasedOnVal("Normal")
	if err := s.SetBasedOnVal(""); err != nil {
		t.Fatal(err)
	}
	v, err := s.BasedOnVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "" {
		t.Errorf("expected empty after clear, got %q", v)
	}
}

// ===========================================================================
// Additional numbering_custom.go coverage
// ===========================================================================

func TestNewDecimalNumber_InvalidTag(t *testing.T) {
	t.Parallel()
	_, err := NewDecimalNumber("bad", 1)
	if err == nil {
		t.Error("expected error for invalid tag")
	}
}

func TestNewCtString_InvalidTag(t *testing.T) {
	t.Parallel()
	_, err := NewCtString("bad", "val")
	if err == nil {
		t.Error("expected error for invalid tag")
	}
}

func TestSetNumIdVal(t *testing.T) {
	t.Parallel()
	npEl := OxmlElement("w:numPr")
	np := &CT_NumPr{Element{e: npEl}}

	// Initially nil
	v, err := np.NumIdVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("expected nil")
	}

	if err := np.SetNumIdVal(5); err != nil {
		t.Fatal(err)
	}
	got, err := np.NumIdVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 5 {
		t.Errorf("NumIdVal() = %v, want 5", got)
	}
}

func TestSetIlvlVal(t *testing.T) {
	t.Parallel()
	npEl := OxmlElement("w:numPr")
	np := &CT_NumPr{Element{e: npEl}}

	v, err := np.IlvlVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("expected nil")
	}

	if err := np.SetIlvlVal(2); err != nil {
		t.Fatal(err)
	}
	got, err := np.IlvlVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 2 {
		t.Errorf("IlvlVal() = %v, want 2", got)
	}
}

// ===========================================================================
// Additional section_custom.go coverage
// ===========================================================================

func TestSectPrSetPageHeight(t *testing.T) {
	t.Parallel()
	sectPrEl := OxmlElement("w:sectPr")
	sp := &CT_SectPr{Element{e: sectPrEl}}

	h := 15840
	if err := sp.SetPageHeight(&h); err != nil {
		t.Fatal(err)
	}
	got, err := sp.PageHeight()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 15840 {
		t.Errorf("PageHeight = %v, want 15840", got)
	}
}

func TestSectPrGetFooterRef(t *testing.T) {
	t.Parallel()
	sectPrEl := OxmlElement("w:sectPr")
	sp := &CT_SectPr{Element{e: sectPrEl}}

	// Add footer ref
	_, err := sp.AddFooterRef(enum.WdHeaderFooterIndexPrimary, "rId1")
	if err != nil {
		t.Fatal(err)
	}
	ref, err := sp.GetFooterRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if ref == nil {
		t.Error("expected non-nil footer ref")
	}
}

// ===========================================================================
// resolveNspTag — tag without prefix
// ===========================================================================

func TestResolveNspTag_NoPrefix(t *testing.T) {
	t.Parallel()
	prefix, local := resolveNspTag("body")
	if prefix != "" || local != "body" {
		t.Errorf("resolveNspTag(body) = (%q, %q), want ('', 'body')", prefix, local)
	}
}

// ===========================================================================
// NSPTag constructors — error paths
// ===========================================================================

func TestNewNSPTag_Panic(t *testing.T) {
	t.Parallel()
	defer func() {
		if r := recover(); r == nil {
			t.Error("NewNSPTag with invalid tag should panic")
		}
	}()
	NewNSPTag("invalidnocolon")
}

func TestNSPTagFromClark_Panic(t *testing.T) {
	t.Parallel()
	defer func() {
		if r := recover(); r == nil {
			t.Error("NSPTagFromClark with invalid clark should panic")
		}
	}()
	NSPTagFromClark("notclark")
}

func TestParseNSPTagFromClark_Errors(t *testing.T) {
	t.Parallel()

	t.Run("empty string", func(t *testing.T) {
		_, err := ParseNSPTagFromClark("")
		if err == nil {
			t.Error("expected error")
		}
	})
	t.Run("no closing brace", func(t *testing.T) {
		_, err := ParseNSPTagFromClark("{http://example.com")
		if err == nil {
			t.Error("expected error")
		}
	})
	t.Run("unknown URI", func(t *testing.T) {
		_, err := ParseNSPTagFromClark("{http://unknown.example.com}tag")
		if err == nil {
			t.Error("expected error")
		}
	})
}

func TestNSPTag_Methods(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:p")

	if tag.Prefix() != "w" {
		t.Errorf("Prefix() = %q, want w", tag.Prefix())
	}
	if tag.LocalPart() != "p" {
		t.Errorf("LocalPart() = %q, want p", tag.LocalPart())
	}
	if tag.NsURI() != NsWml {
		t.Errorf("NsURI() = %q, want %q", tag.NsURI(), NsWml)
	}
	if tag.String() != "w:p" {
		t.Errorf("String() = %q, want w:p", tag.String())
	}
	nsm := tag.NsMap()
	if nsm["w"] != NsWml {
		t.Errorf("NsMap() = %v", nsm)
	}
}
