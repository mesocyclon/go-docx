package oxml

import (
	"strings"
	"testing"
)

// -----------------------------------------------------------------------
// replacetext_test.go — unit tests for cross-run text replacement
// -----------------------------------------------------------------------

// helper: build a <w:p> with programmatic API, return CT_P.
func buildP(fn func(p *CT_P)) *CT_P {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}
	fn(p)
	return p
}

// helper: verify paragraph text equals expected.
func assertText(t *testing.T, p *CT_P, expected string) {
	t.Helper()
	got := p.ParagraphText()
	if got != expected {
		t.Errorf("ParagraphText() = %q, want %q", got, expected)
	}
}

// helper: count child elements with given tag in an element tree (non-recursive, direct children of runs).
func countTagsInRuns(p *CT_P, tag string) int {
	count := 0
	for _, child := range p.e.ChildElements() {
		if child.Space == "w" && (child.Tag == "r" || child.Tag == "hyperlink") {
			var runs []*CT_R
			if child.Tag == "r" {
				runs = append(runs, &CT_R{Element{e: child}})
			} else {
				for _, gc := range child.ChildElements() {
					if gc.Space == "w" && gc.Tag == "r" {
						runs = append(runs, &CT_R{Element{e: gc}})
					}
				}
			}
			for _, r := range runs {
				for _, rc := range r.e.ChildElements() {
					if rc.Space == "w" && rc.Tag == tag {
						count++
					}
				}
			}
		}
	}
	return count
}

// --- Test: empty old → 0 ---

func TestReplaceText_EmptyOld(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Hello")
	})
	n := p.ReplaceText("", "X")
	if n != 0 {
		t.Errorf("expected 0 replacements, got %d", n)
	}
	assertText(t, p, "Hello")
}

// --- Test: old == new → 0, no modification ---

func TestReplaceText_OldEqualsNew(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Hello")
	})
	n := p.ReplaceText("Hello", "Hello")
	if n != 0 {
		t.Errorf("expected 0 replacements, got %d", n)
	}
	assertText(t, p, "Hello")
}

// --- Test: no runs → 0 ---

func TestReplaceText_NoRuns(t *testing.T) {
	p := buildP(func(p *CT_P) {
		// empty paragraph, just pPr
		p.GetOrAddPPr()
	})
	n := p.ReplaceText("X", "Y")
	if n != 0 {
		t.Errorf("expected 0 replacements, got %d", n)
	}
}

// --- Test: simple replacement in single run ---

func TestReplaceText_SingleRun(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Hello World")
	})
	n := p.ReplaceText("World", "Go")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Hello Go")
}

// --- Test: replacement spanning 2 runs ---

func TestReplaceText_CrossTwoRuns(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Hel")
		p.AddR().AddTWithText("lo World")
	})
	n := p.ReplaceText("Hello", "Hi")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Hi World")
}

// --- Test: replacement spanning 3+ runs (middle ones emptied) ---

func TestReplaceText_CrossThreeRuns(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("AB")
		p.AddR().AddTWithText("CD")
		p.AddR().AddTWithText("EF")
	})
	n := p.ReplaceText("BCDE", "X")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "AXF")
}

// --- Test: replacement crossing run ↔ hyperlink boundary ---

func TestReplaceText_CrossRunHyperlink(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("click ")
		h := p.AddHyperlink()
		h.AddR().AddTWithText("here please")
	})
	n := p.ReplaceText("click here", "go there")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "go there please")

	// Verify hyperlink element still exists.
	if len(p.HyperlinkList()) == 0 {
		t.Error("hyperlink element should still exist")
	}
}

// --- Test: replacement including <w:tab> → tab removed ---

func TestReplaceText_IncludesTab(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("A")
		r.AddTab()
		r.AddTWithText("B")
	})
	// Text is "A\tB"
	assertText(t, p, "A\tB")

	n := p.ReplaceText("A\tB", "XY")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "XY")

	// <w:tab> should be removed.
	if cnt := countTagsInRuns(p, "tab"); cnt != 0 {
		t.Errorf("expected 0 <w:tab> elements, got %d", cnt)
	}
}

// --- Test: replacement including <w:br> (textWrapping) → br removed ---

func TestReplaceText_IncludesBr(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("line1")
		r.AddBr() // textWrapping by default
		r.AddTWithText("line2")
	})
	assertText(t, p, "line1\nline2")

	n := p.ReplaceText("line1\nline2", "single")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "single")

	if cnt := countTagsInRuns(p, "br"); cnt != 0 {
		t.Errorf("expected 0 <w:br> elements, got %d", cnt)
	}
}

// --- Test: multiple occurrences ---

func TestReplaceText_MultipleOccurrences(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("aXbXc")
	})
	n := p.ReplaceText("X", "Y")
	if n != 2 {
		t.Errorf("expected 2 replacements, got %d", n)
	}
	assertText(t, p, "aYbYc")
}

// --- Test: replace with empty string (deletion) ---

func TestReplaceText_DeleteText(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Hello World")
	})
	n := p.ReplaceText(" World", "")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Hello")
}

// --- Test: replacement increasing length ---

func TestReplaceText_IncreaseLength(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("A")
		p.AddR().AddTWithText("B")
	})
	n := p.ReplaceText("AB", "XYZXYZ")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "XYZXYZ")
}

// --- Test: <w:drawing> inside run is not affected ---

func TestReplaceText_DrawingPreserved(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("before")
		r.AddDrawing() // adds empty <w:drawing>
		r.AddTWithText(" after")
	})
	n := p.ReplaceText("before", "REPLACED")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "REPLACED after")

	// Drawing must still be present.
	if cnt := countTagsInRuns(p, "drawing"); cnt != 1 {
		t.Errorf("expected 1 <w:drawing>, got %d", cnt)
	}
}

// --- Test: <w:rPr> is preserved ---

func TestReplaceText_FormattingPreserved(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r1 := p.AddR()
		rPr1 := r1.GetOrAddRPr()
		bEl := OxmlElement("w:b")
		rPr1.e.AddChild(bEl)
		r1.AddTWithText("bold")

		r2 := p.AddR()
		rPr2 := r2.GetOrAddRPr()
		iEl := OxmlElement("w:i")
		rPr2.e.AddChild(iEl)
		r2.AddTWithText("italic")
	})

	n := p.ReplaceText("bolditalic", "REPLACED")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "REPLACED")

	// Both rPr elements should still exist.
	runs := p.RList()
	if len(runs) < 2 {
		t.Fatalf("expected at least 2 runs, got %d", len(runs))
	}

	rPr1 := runs[0].RPr()
	if rPr1 == nil {
		t.Error("run 1 should still have rPr")
	} else if rPr1.FindChild("w:b") == nil {
		t.Error("run 1 should still have <w:b>")
	}

	rPr2 := runs[1].RPr()
	if rPr2 == nil {
		t.Error("run 2 should still have rPr")
	} else if rPr2.FindChild("w:i") == nil {
		t.Error("run 2 should still have <w:i>")
	}
}

// --- Test: xml:space="preserve" updated ---

func TestReplaceText_PreserveSpace(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("NoSpaces")
	})
	// Replace with text that has leading space.
	p.ReplaceText("NoSpaces", " leading")

	tElems := p.RList()[0].TList()
	if len(tElems) == 0 {
		t.Fatal("expected at least one <w:t>")
	}
	spaceAttr := ""
	for _, attr := range tElems[0].e.Attr {
		if attr.Key == "space" {
			spaceAttr = attr.Value
		}
	}
	if spaceAttr != "preserve" {
		t.Errorf("expected xml:space='preserve', got %q", spaceAttr)
	}
}

// --- Test: <w:commentRangeStart/End> not affected ---

func TestReplaceText_CommentRangePreserved(t *testing.T) {
	p := buildP(func(p *CT_P) {
		// Manually add commentRangeStart before run.
		crs := OxmlElement("w:commentRangeStart")
		crs.CreateAttr("w:id", "1")
		p.e.AddChild(crs)

		r := p.AddR()
		r.AddTWithText("commented text")

		cre := OxmlElement("w:commentRangeEnd")
		cre.CreateAttr("w:id", "1")
		p.e.AddChild(cre)
	})

	n := p.ReplaceText("commented", "replaced")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "replaced text")

	// Comment markers should still exist.
	var foundStart, foundEnd bool
	for _, child := range p.e.ChildElements() {
		if child.Space == "w" && child.Tag == "commentRangeStart" {
			foundStart = true
		}
		if child.Space == "w" && child.Tag == "commentRangeEnd" {
			foundEnd = true
		}
	}
	if !foundStart {
		t.Error("commentRangeStart should still exist")
	}
	if !foundEnd {
		t.Error("commentRangeEnd should still exist")
	}
}

// --- Test: <w:commentReference> in run not affected ---

func TestReplaceText_CommentRefPreserved(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("text")
		// Manually add commentReference.
		cr := OxmlElement("w:commentReference")
		cr.CreateAttr("w:id", "1")
		r.e.AddChild(cr)
	})

	n := p.ReplaceText("text", "new")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "new")

	// commentReference should still exist in the run.
	found := false
	for _, child := range p.RList()[0].e.ChildElements() {
		if child.Space == "w" && child.Tag == "commentReference" {
			found = true
		}
	}
	if !found {
		t.Error("commentReference should still exist")
	}
}

// --- Test: Cyrillic / multibyte UTF-8 ---

func TestReplaceText_CyrillicUTF8(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Привет мир")
	})
	n := p.ReplaceText("мир", "Go")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Привет Go")
}

// --- Test: Cyrillic across run boundary ---

func TestReplaceText_CyrillicCrossRun(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Прив")
		p.AddR().AddTWithText("ет мир")
	})
	n := p.ReplaceText("Привет", "Здравствуй")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Здравствуй мир")
}

// --- Test: not found → 0, XML unchanged ---

func TestReplaceText_NotFound(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("Hello World")
	})
	n := p.ReplaceText("Missing", "X")
	if n != 0 {
		t.Errorf("expected 0 replacements, got %d", n)
	}
	assertText(t, p, "Hello World")
}

// --- Test: multi-run formatting preserved across replacement boundary ---

func TestReplaceText_CrossRunFormattingPreserved(t *testing.T) {
	p := buildP(func(p *CT_P) {
		// Run 1: bold "Hel"
		r1 := p.AddR()
		rPr1 := r1.GetOrAddRPr()
		rPr1.e.AddChild(OxmlElement("w:b"))
		r1.AddTWithText("Hel")

		// Run 2: italic "lo World"
		r2 := p.AddR()
		rPr2 := r2.GetOrAddRPr()
		rPr2.e.AddChild(OxmlElement("w:i"))
		r2.AddTWithText("lo World")
	})

	n := p.ReplaceText("Hello", "Hi")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Hi World")

	runs := p.RList()
	// Run 1 should still be bold.
	if rPr := runs[0].RPr(); rPr == nil || rPr.FindChild("w:b") == nil {
		t.Error("run 1 should still be bold")
	}
	// Run 2 should still be italic.
	if rPr := runs[1].RPr(); rPr == nil || rPr.FindChild("w:i") == nil {
		t.Error("run 2 should still be italic")
	}
}

// --- Test: page break <w:br type="page"> does NOT produce text ---

func TestReplaceText_PageBreakIgnored(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("before")
		br := r.NewDetachedBr()
		_ = br.SetType("page")
		r.AttachBr(br)
		r.AddTWithText("after")
	})
	// Page break should not appear in text.
	assertText(t, p, "beforeafter")

	n := p.ReplaceText("beforeafter", "X")
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "X")
}

// --- Test: multiple occurrences across multiple runs ---

func TestReplaceText_MultipleOccurrencesCrossRun(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("AXB")
		p.AddR().AddTWithText("XC")
	})
	n := p.ReplaceText("X", "Y")
	if n != 2 {
		t.Errorf("expected 2 replacements, got %d", n)
	}
	assertText(t, p, "AYBYC")
}

// --- Test: replacement at very start and end ---

func TestReplaceText_AtStartAndEnd(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("XXhelloXX")
	})
	n := p.ReplaceText("XX", "")
	if n != 2 {
		t.Errorf("expected 2 replacements, got %d", n)
	}
	assertText(t, p, "hello")
}

// --- Test: collectTextAtoms and findOccurrences directly ---

func TestCollectTextAtoms_Basic(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("Hello")
		r.AddTab()
		r.AddTWithText("World")
	})

	atoms, fullText := collectTextAtoms(p.e)
	if fullText != "Hello\tWorld" {
		t.Errorf("fullText = %q, want %q", fullText, "Hello\tWorld")
	}
	if len(atoms) != 3 {
		t.Fatalf("expected 3 atoms, got %d", len(atoms))
	}

	// Atom 0: <w:t>"Hello", pos 0, editable
	if atoms[0].text != "Hello" || atoms[0].startPos != 0 || !atoms[0].editable {
		t.Errorf("atom 0: %+v", atoms[0])
	}
	// Atom 1: <w:tab>, pos 5, fixed
	if atoms[1].text != "\t" || atoms[1].startPos != 5 || atoms[1].editable {
		t.Errorf("atom 1: %+v", atoms[1])
	}
	// Atom 2: <w:t>"World", pos 6, editable
	if atoms[2].text != "World" || atoms[2].startPos != 6 || !atoms[2].editable {
		t.Errorf("atom 2: %+v", atoms[2])
	}
}

func TestFindOccurrences(t *testing.T) {
	tests := []struct {
		text     string
		old      string
		expected []int
	}{
		{"aXbXc", "X", []int{1, 3}},
		{"hello", "X", nil},
		{"AAAA", "AA", []int{0, 2}}, // non-overlapping
		{"", "X", nil},
	}
	for _, tt := range tests {
		got := findOccurrences(tt.text, tt.old)
		if len(got) != len(tt.expected) {
			t.Errorf("findOccurrences(%q, %q) = %v, want %v", tt.text, tt.old, got, tt.expected)
			continue
		}
		for i := range got {
			if got[i] != tt.expected[i] {
				t.Errorf("findOccurrences(%q, %q)[%d] = %d, want %d", tt.text, tt.old, i, got[i], tt.expected[i])
			}
		}
	}
}

// --- Test: whole paragraph replacement ---

func TestReplaceText_WholeParagraph(t *testing.T) {
	p := buildP(func(p *CT_P) {
		p.AddR().AddTWithText("replace me entirely")
	})
	n := p.ReplaceText("replace me entirely", "done")
	if n != 1 {
		t.Errorf("expected 1, got %d", n)
	}
	assertText(t, p, "done")
}

// --- Test: hyperlink r:id attribute preserved ---

func TestReplaceText_HyperlinkRIdPreserved(t *testing.T) {
	p := buildP(func(p *CT_P) {
		h := p.AddHyperlink()
		_ = h.SetRId("rId5")
		h.AddR().AddTWithText("link text")
	})

	p.ReplaceText("link", "new")
	assertText(t, p, "new text")

	hls := p.HyperlinkList()
	if len(hls) != 1 {
		t.Fatalf("expected 1 hyperlink, got %d", len(hls))
	}
	if rId := hls[0].RId(); rId != "rId5" {
		t.Errorf("hyperlink rId = %q, want %q", rId, "rId5")
	}
}

// --- Test: replacement of only fixed atoms (edge case) ---

func TestReplaceText_OnlyFixedAtoms(t *testing.T) {
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		r.AddTWithText("A")
		r.AddTab()
		r.AddTab()
		r.AddTWithText("B")
	})
	// Text is "A\t\tB", replace "\t\t" with "X"
	n := p.ReplaceText("\t\t", "X")
	if n != 1 {
		t.Errorf("expected 1, got %d", n)
	}
	got := p.ParagraphText()
	if !strings.Contains(got, "A") || !strings.Contains(got, "B") || !strings.Contains(got, "X") {
		t.Errorf("unexpected text: %q", got)
	}
}

// --- Regression: findRunForAtom must return the correct run, not a neighbor ---

// TestReplaceText_FixedAtomsCrossRun_CorrectRun verifies that when a match
// spans only fixed atoms across two differently-formatted runs, the
// replacement <w:t> is inserted into the first matched run (preserving its
// formatting), not into an unrelated neighbor run.
func TestReplaceText_FixedAtomsCrossRun_CorrectRun(t *testing.T) {
	// Structure:
	//   Run 1 (bold):   <w:tab/>           → "\t"
	//   Run 2 (normal): <w:tab/><w:t>B</w:t> → "\tB"
	// Full text: "\t\tB"
	// Replace "\t\t" → "X"
	//
	// Before fix: findRunForAtom could return Run 2 (nearest alive parent),
	// so "X" would appear in the normal run instead of the bold run.
	boolTrue := true
	p := buildP(func(p *CT_P) {
		r1 := p.AddR()
		_ = r1.GetOrAddRPr().SetBoldVal(&boolTrue) // Run 1: bold
		r1.AddTab()

		r2 := p.AddR()
		r2.AddTab()
		r2.AddTWithText("B")
	})

	n := p.ReplaceText("\t\t", "X")
	if n != 1 {
		t.Fatalf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "XB")

	// The replacement "X" must be in Run 1 (the bold run), not Run 2.
	runs := p.RList()
	if len(runs) < 1 {
		t.Fatal("expected at least 1 run")
	}
	r1 := runs[0]
	// Check that Run 1 has a <w:t> child with "X".
	found := false
	for _, child := range r1.RawElement().ChildElements() {
		if child.Space == "w" && child.Tag == "t" && child.Text() == "X" {
			found = true
			break
		}
	}
	if !found {
		t.Error("replacement text 'X' should be in the first (bold) run, but was not found there")
	}
	// Verify bold rPr is still present in that run.
	rPr := r1.RPr()
	if rPr == nil {
		t.Error("first run should still have rPr (bold)")
	}
}

// TestReplaceText_FixedAtomsNotLeakedToHyperlink verifies that replacing
// fixed atoms adjacent to a hyperlink run does not insert the replacement
// text inside the hyperlink.
func TestReplaceText_FixedAtomsNotLeakedToHyperlink(t *testing.T) {
	// Structure:
	//   Run 1:      <w:tab/><w:tab/>   → "\t\t"
	//   Hyperlink:  Run 2: <w:t>link</w:t>  → "link"
	// Full text: "\t\tlink"
	// Replace "\t\t" → "X"
	p := buildP(func(p *CT_P) {
		r1 := p.AddR()
		r1.AddTab()
		r1.AddTab()

		h := p.AddHyperlink()
		h.AddR().AddTWithText("link")
	})

	n := p.ReplaceText("\t\t", "X")
	if n != 1 {
		t.Fatalf("expected 1 replacement, got %d", n)
	}
	assertText(t, p, "Xlink")

	// Verify "X" is NOT inside the hyperlink.
	for _, hl := range p.HyperlinkList() {
		text := hl.HyperlinkText()
		if strings.Contains(text, "X") {
			t.Errorf("replacement 'X' leaked into hyperlink, hyperlink text = %q", text)
		}
	}
}

// TestReplaceText_FixedAtoms_InsertAfterRPr verifies that the replacement
// <w:t> is inserted right after <w:rPr> (not appended at the end of the run).
func TestReplaceText_FixedAtoms_InsertAfterRPr(t *testing.T) {
	boolTrue := true
	p := buildP(func(p *CT_P) {
		r := p.AddR()
		_ = r.GetOrAddRPr().SetBoldVal(&boolTrue)
		r.AddTab()
		r.AddTab()
	})

	n := p.ReplaceText("\t\t", "X")
	if n != 1 {
		t.Fatalf("expected 1 replacement, got %d", n)
	}

	// The run's children should be: rPr, then t (not t appended at end).
	r := p.RList()[0]
	children := r.RawElement().ChildElements()
	if len(children) < 2 {
		t.Fatalf("expected at least 2 children in run, got %d", len(children))
	}
	if children[0].Tag != "rPr" {
		t.Errorf("first child should be rPr, got %s:%s", children[0].Space, children[0].Tag)
	}
	if children[1].Tag != "t" || children[1].Text() != "X" {
		t.Errorf("second child should be <w:t>X</w:t>, got <%s:%s>%s",
			children[1].Space, children[1].Tag, children[1].Text())
	}
}
