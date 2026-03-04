package docx

import (
	"bytes"
	"strings"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// ---------------------------------------------------------------------------
// Phase 2: buildTableElement tests
// ---------------------------------------------------------------------------

func TestBuildTableElement_ZeroRows(t *testing.T) {
	td := TableData{Rows: [][]string{}}
	_, err := buildTableElement(td, 9360)
	if err == nil {
		t.Fatal("expected error for zero rows")
	}
	if !strings.Contains(err.Error(), "no rows") {
		t.Errorf("unexpected error: %v", err)
	}
}

func TestBuildTableElement_ZeroCols(t *testing.T) {
	td := TableData{Rows: [][]string{nil}}
	_, err := buildTableElement(td, 9360)
	if err == nil {
		t.Fatal("expected error for zero cols")
	}
	if !strings.Contains(err.Error(), "no cells") {
		t.Errorf("unexpected error: %v", err)
	}
}

func TestBuildTableElement_JaggedRows(t *testing.T) {
	td := TableData{Rows: [][]string{
		{"a", "b", "c"},
		{"d", "e"}, // 2 cells instead of 3
	}}
	_, err := buildTableElement(td, 9360)
	if err == nil {
		t.Fatal("expected error for jagged rows")
	}
	if !strings.Contains(err.Error(), "row 1 has 2 cells, expected 3") {
		t.Errorf("unexpected error: %v", err)
	}
}

func TestBuildTableElement_Simple(t *testing.T) {
	td := TableData{Rows: [][]string{
		{"A1", "B1"},
		{"A2", "B2"},
	}}
	el, err := buildTableElement(td, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if el.Tag != "tbl" || el.Space != "w" {
		t.Fatalf("expected <w:tbl>, got <%s:%s>", el.Space, el.Tag)
	}

	// Verify structure: 2 <w:tr>, each with 2 <w:tc>.
	trs := rtChildrenByTag(el, "w", "tr")
	if len(trs) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(trs))
	}
	for r, tr := range trs {
		tcs := rtChildrenByTag(tr, "w", "tc")
		if len(tcs) != 2 {
			t.Fatalf("row %d: expected 2 cells, got %d", r, len(tcs))
		}
	}

	// Verify text content.
	rtAssertCellText(t, el, 0, 0, "A1")
	rtAssertCellText(t, el, 0, 1, "B1")
	rtAssertCellText(t, el, 1, 0, "A2")
	rtAssertCellText(t, el, 1, 1, "B2")
}

func TestBuildTableElement_EmptyCells(t *testing.T) {
	td := TableData{Rows: [][]string{
		{"text", ""},
	}}
	el, err := buildTableElement(td, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Cell 0: should have text.
	rtAssertCellText(t, el, 0, 0, "text")

	// Cell 1: empty text → no <w:r> added, but <w:p> exists.
	trs := rtChildrenByTag(el, "w", "tr")
	tcs := rtChildrenByTag(trs[0], "w", "tc")
	ps := rtChildrenByTag(tcs[1], "w", "p")
	if len(ps) != 1 {
		t.Fatalf("empty cell should have 1 <w:p>, got %d", len(ps))
	}
	rs := rtChildrenByTag(ps[0], "w", "r")
	if len(rs) != 0 {
		t.Errorf("empty cell paragraph should have 0 runs, got %d", len(rs))
	}
}

func TestBuildTableElement_WithStyle(t *testing.T) {
	td := TableData{
		Rows:  [][]string{{"a"}},
		Style: StyleName("Table Grid"),
	}
	el, err := buildTableElement(td, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Verify tblStyle was set via the CT_Tbl API.
	tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(el)}
	styleVal, err := tbl.TblStyleVal()
	if err != nil {
		t.Fatalf("TblStyleVal error: %v", err)
	}
	if styleVal != "Table Grid" {
		t.Errorf("tblStyle = %q, want %q", styleVal, "Table Grid")
	}
}

func TestBuildTableElement_NilStyle(t *testing.T) {
	td := TableData{
		Rows: [][]string{{"a"}},
		// Style is zero-value (nil)
	}
	el, err := buildTableElement(td, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(el)}
	styleVal, _ := tbl.TblStyleVal()
	if styleVal != "" {
		t.Errorf("expected no style, got %q", styleVal)
	}
}

func TestBuildTableElement_SingleCell(t *testing.T) {
	td := TableData{Rows: [][]string{{"only"}}}
	el, err := buildTableElement(td, 5000)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	trs := rtChildrenByTag(el, "w", "tr")
	if len(trs) != 1 {
		t.Fatalf("expected 1 row, got %d", len(trs))
	}
	tcs := rtChildrenByTag(trs[0], "w", "tc")
	if len(tcs) != 1 {
		t.Fatalf("expected 1 cell, got %d", len(tcs))
	}
	rtAssertCellText(t, el, 0, 0, "only")
}

// ---------------------------------------------------------------------------
// Phase 4: replaceTagWithElements engine tests
// ---------------------------------------------------------------------------

// TestReplaceTagWithElements_BuildFnCalledPerPlaceholder verifies that buildFn
// is called once per placeholder, and each call returns a fresh element
// (pointer identity check).
func TestReplaceTagWithElements_BuildFnCalledPerPlaceholder(t *testing.T) {
	// Build a <w:body> with a paragraph containing two tags.
	body := oxml.OxmlElement("w:body")
	p := body.CreateElement("p")
	p.Space = "w"
	r := p.CreateElement("r")
	r.Space = "w"
	wt := r.CreateElement("t")
	wt.Space = "w"
	wt.SetText("AAA[<TAG>]BBB[<TAG>]CCC")

	bic := newBlockItemContainer(body, nil)

	var receivedElements []*etree.Element
	callCount := 0
	buildFn := func(widthTwips int) ([]*etree.Element, error) {
		callCount++
		el := oxml.OxmlElement("w:tbl")
		receivedElements = append(receivedElements, el)
		return []*etree.Element{el}, nil
	}

	count, err := bic.replaceTagWithElements("[<TAG>]", buildFn, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
	if callCount != 2 {
		t.Errorf("buildFn called %d times, want 2", callCount)
	}
	// Pointer identity: each call returned a different element.
	if len(receivedElements) == 2 && receivedElements[0] == receivedElements[1] {
		t.Error("buildFn returned the same element twice (pointer identity)")
	}

	// Verify the body now contains: <w:p>AAA</w:p> <w:tbl/> <w:p>BBB</w:p> <w:tbl/> <w:p>CCC</w:p>
	children := body.ChildElements()
	if len(children) != 5 {
		t.Fatalf("expected 5 children in body, got %d", len(children))
	}
	wantTags := []string{"p", "tbl", "p", "tbl", "p"}
	for i, child := range children {
		if child.Tag != wantTags[i] {
			t.Errorf("child[%d].Tag = %q, want %q", i, child.Tag, wantTags[i])
		}
	}
}

// TestReplaceTagWithElements_BuildFnReceivesCellWidth verifies that when the
// engine recurses into a table cell, buildFn receives the cell width (not the
// parent body width).
func TestReplaceTagWithElements_BuildFnReceivesCellWidth(t *testing.T) {
	// Build:
	// <w:body>
	//   <w:tbl>
	//     <w:tr>
	//       <w:tc>
	//         <w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>
	//         <w:p><w:r><w:t>[<TAG>]</w:t></w:r></w:p>
	//       </w:tc>
	//     </w:tr>
	//   </w:tbl>
	// </w:body>
	body := oxml.OxmlElement("w:body")
	tblE := body.CreateElement("tbl")
	tblE.Space = "w"
	trE := tblE.CreateElement("tr")
	trE.Space = "w"
	tcE := trE.CreateElement("tc")
	tcE.Space = "w"
	tcPrE := tcE.CreateElement("tcPr")
	tcPrE.Space = "w"
	tcW := tcPrE.CreateElement("tcW")
	tcW.Space = "w"
	tcW.CreateAttr("w:type", "dxa")
	tcW.CreateAttr("w:w", "3000")
	pE := tcE.CreateElement("p")
	pE.Space = "w"
	rE := pE.CreateElement("r")
	rE.Space = "w"
	wtE := rE.CreateElement("t")
	wtE.Space = "w"
	wtE.SetText("[<TAG>]")

	bic := newBlockItemContainer(body, nil)

	var receivedWidth int
	buildFn := func(widthTwips int) ([]*etree.Element, error) {
		receivedWidth = widthTwips
		// Return a minimal <w:tbl> so the replacement proceeds.
		return []*etree.Element{oxml.OxmlElement("w:tbl")}, nil
	}

	count, err := bic.replaceTagWithElements("[<TAG>]", buildFn, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}
	if receivedWidth != 3000 {
		t.Errorf("buildFn received widthTwips=%d, want 3000", receivedWidth)
	}
}

// TestReplaceTagWithElements_EnsureTcHasParagraph verifies that after the
// engine replaces the only <w:p> in a <w:tc>, an empty <w:p> is added to
// satisfy the OOXML schema invariant.
func TestReplaceTagWithElements_EnsureTcHasParagraph(t *testing.T) {
	// Build:
	// <w:body>
	//   <w:tbl>
	//     <w:tr>
	//       <w:tc>
	//         <w:tcPr><w:tcW w:type="dxa" w:w="5000"/></w:tcPr>
	//         <w:p><w:r><w:t>[<TAG>]</w:t></w:r></w:p>
	//       </w:tc>
	//     </w:tr>
	//   </w:tbl>
	// </w:body>
	body := oxml.OxmlElement("w:body")
	tblE := body.CreateElement("tbl")
	tblE.Space = "w"
	trE := tblE.CreateElement("tr")
	trE.Space = "w"
	tcE := trE.CreateElement("tc")
	tcE.Space = "w"
	tcPrE := tcE.CreateElement("tcPr")
	tcPrE.Space = "w"
	tcW := tcPrE.CreateElement("tcW")
	tcW.Space = "w"
	tcW.CreateAttr("w:type", "dxa")
	tcW.CreateAttr("w:w", "5000")
	pE := tcE.CreateElement("p")
	pE.Space = "w"
	rE := pE.CreateElement("r")
	rE.Space = "w"
	wtE := rE.CreateElement("t")
	wtE.Space = "w"
	wtE.SetText("[<TAG>]")

	bic := newBlockItemContainer(body, nil)

	buildFn := func(widthTwips int) ([]*etree.Element, error) {
		// Return a nested table element.
		return []*etree.Element{oxml.OxmlElement("w:tbl")}, nil
	}

	count, err := bic.replaceTagWithElements("[<TAG>]", buildFn, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// The cell should now contain: <w:tcPr>, <w:tbl>, <w:p> (trailing empty).
	tcChildren := tcE.ChildElements()
	hasParagraph := false
	hasNestedTbl := false
	for _, child := range tcChildren {
		if child.Space == "w" && child.Tag == "p" {
			hasParagraph = true
		}
		if child.Space == "w" && child.Tag == "tbl" {
			hasNestedTbl = true
		}
	}
	if !hasNestedTbl {
		t.Error("expected nested <w:tbl> in cell after replacement")
	}
	if !hasParagraph {
		t.Error("expected trailing <w:p> in cell (OOXML invariant)")
	}
}

// TestSpliceElement_PreservesCharData verifies that CharData (whitespace text
// nodes) between elements is not lost during splice operations.
func TestSpliceElement_PreservesCharData(t *testing.T) {
	// Build a body with CharData between elements:
	// <w:body>\n  <w:p>target</w:p>\n  <w:p>after</w:p>\n</w:body>
	body := oxml.OxmlElement("w:body")

	body.CreateCharData("\n  ")
	target := body.CreateElement("p")
	target.Space = "w"
	rTarget := target.CreateElement("r")
	rTarget.Space = "w"
	tTarget := rTarget.CreateElement("t")
	tTarget.Space = "w"
	tTarget.SetText("target")

	body.CreateCharData("\n  ")
	after := body.CreateElement("p")
	after.Space = "w"
	rAfter := after.CreateElement("r")
	rAfter.Space = "w"
	tAfter := rAfter.CreateElement("t")
	tAfter.Space = "w"
	tAfter.SetText("after")

	body.CreateCharData("\n")

	bic := newBlockItemContainer(body, nil)

	// Replace `target` element with two new elements.
	new1 := oxml.OxmlElement("w:p")
	new2 := oxml.OxmlElement("w:tbl")
	bic.spliceElement(target, []*etree.Element{new1, new2})

	// Count CharData tokens in body.Child.
	charDataCount := 0
	for _, tok := range body.Child {
		if _, ok := tok.(*etree.CharData); ok {
			charDataCount++
		}
	}
	// Original had 3 CharData tokens (\n  , \n  , \n).
	// After splice, all 3 should still be present.
	if charDataCount != 3 {
		t.Errorf("CharData count = %d, want 3 (CharData was lost)", charDataCount)
	}

	// Verify element order: new1, new2, after.
	elems := body.ChildElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elems))
	}
	if elems[0] != new1 {
		t.Error("element[0] should be new1")
	}
	if elems[1] != new2 {
		t.Error("element[1] should be new2")
	}
	if elems[2] != after {
		t.Error("element[2] should be 'after' paragraph")
	}
}

// TestReplaceWithTable_Simple verifies the replaceWithTable wrapper with a
// basic case: single tag replaced by a 2x2 table.
func TestReplaceWithTable_Simple(t *testing.T) {
	body := oxml.OxmlElement("w:body")
	p := body.CreateElement("p")
	p.Space = "w"
	r := p.CreateElement("r")
	r.Space = "w"
	wt := r.CreateElement("t")
	wt.Space = "w"
	wt.SetText("Before [<T>] After")

	bic := newBlockItemContainer(body, nil)
	td := TableData{Rows: [][]string{{"A1", "B1"}, {"A2", "B2"}}}

	count, err := bic.replaceWithTable("[<T>]", td, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Body should have: <w:p>Before </w:p> <w:tbl/> <w:p> After</w:p>
	elems := body.ChildElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 children, got %d", len(elems))
	}
	if elems[0].Tag != "p" {
		t.Errorf("child[0].Tag = %q, want 'p'", elems[0].Tag)
	}
	if elems[1].Tag != "tbl" {
		t.Errorf("child[1].Tag = %q, want 'tbl'", elems[1].Tag)
	}
	if elems[2].Tag != "p" {
		t.Errorf("child[2].Tag = %q, want 'p'", elems[2].Tag)
	}

	// Verify the inserted table has 2 rows and 2 cols.
	trs := rtChildrenByTag(elems[1], "w", "tr")
	if len(trs) != 2 {
		t.Fatalf("expected 2 rows in table, got %d", len(trs))
	}
	for i, tr := range trs {
		tcs := rtChildrenByTag(tr, "w", "tc")
		if len(tcs) != 2 {
			t.Errorf("row %d: expected 2 cells, got %d", i, len(tcs))
		}
	}

	// Verify cell content.
	rtAssertCellText(t, elems[1], 0, 0, "A1")
	rtAssertCellText(t, elems[1], 1, 1, "B2")
}

// TestReplaceWithTable_DefensiveCopy verifies that the caller can safely
// modify TableData.Rows after calling replaceWithTable.
func TestReplaceWithTable_DefensiveCopy(t *testing.T) {
	body := oxml.OxmlElement("w:body")
	p1 := body.CreateElement("p")
	p1.Space = "w"
	r1 := p1.CreateElement("r")
	r1.Space = "w"
	wt1 := r1.CreateElement("t")
	wt1.Space = "w"
	wt1.SetText("[<T>]")
	p2 := body.CreateElement("p")
	p2.Space = "w"
	r2 := p2.CreateElement("r")
	r2.Space = "w"
	wt2 := r2.CreateElement("t")
	wt2.Space = "w"
	wt2.SetText("[<T>]")

	bic := newBlockItemContainer(body, nil)
	td := TableData{Rows: [][]string{{"original"}}}

	count, err := bic.replaceWithTable("[<T>]", td, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}

	// Mutate the original — should not affect already-inserted tables.
	td.Rows[0][0] = "MUTATED"

	// Both inserted tables should have "original", not "MUTATED".
	for _, el := range body.ChildElements() {
		if el.Space == "w" && el.Tag == "tbl" {
			rtAssertCellText(t, el, 0, 0, "original")
		}
	}
}

// TestReplaceWithTable_ErrorPropagation verifies that buildTableElement errors
// (e.g. jagged rows) are propagated through replaceWithTable.
func TestReplaceWithTable_ErrorPropagation(t *testing.T) {
	body := oxml.OxmlElement("w:body")
	p := body.CreateElement("p")
	p.Space = "w"
	r := p.CreateElement("r")
	r.Space = "w"
	wt := r.CreateElement("t")
	wt.Space = "w"
	wt.SetText("[<T>]")

	bic := newBlockItemContainer(body, nil)
	td := TableData{Rows: [][]string{{"a", "b"}, {"c"}}} // jagged

	_, err := bic.replaceWithTable("[<T>]", td, 9360)
	if err == nil {
		t.Fatal("expected error for jagged rows")
	}
	if !strings.Contains(err.Error(), "row 1 has 1 cells, expected 2") {
		t.Errorf("unexpected error message: %v", err)
	}
}

// TestReplaceTagWithElements_NoMatch verifies that the engine returns 0 and
// does not modify the container when the tag is not found.
func TestReplaceTagWithElements_NoMatch(t *testing.T) {
	body := oxml.OxmlElement("w:body")
	p := body.CreateElement("p")
	p.Space = "w"
	r := p.CreateElement("r")
	r.Space = "w"
	wt := r.CreateElement("t")
	wt.Space = "w"
	wt.SetText("no tags here")

	bic := newBlockItemContainer(body, nil)
	callCount := 0
	buildFn := func(w int) ([]*etree.Element, error) {
		callCount++
		return []*etree.Element{oxml.OxmlElement("w:tbl")}, nil
	}

	count, err := bic.replaceTagWithElements("[<X>]", buildFn, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("count = %d, want 0", count)
	}
	if callCount != 0 {
		t.Errorf("buildFn called %d times, want 0", callCount)
	}

	// Body should still have exactly 1 paragraph.
	elems := body.ChildElements()
	if len(elems) != 1 || elems[0].Tag != "p" {
		t.Errorf("body should be unchanged, got %d children", len(elems))
	}
}

// TestReplaceTagWithElements_RecursionIntoPreExistingTable verifies that the
// engine recurses into tables that existed before the replacement.
func TestReplaceTagWithElements_RecursionIntoPreExistingTable(t *testing.T) {
	// Build:
	// <w:body>
	//   <w:tbl>
	//     <w:tr><w:tc>
	//       <w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr>
	//       <w:p><w:r><w:t>Keep this</w:t></w:r></w:p>
	//       <w:p><w:r><w:t>[<TAG>]</w:t></w:r></w:p>
	//     </w:tc></w:tr>
	//   </w:tbl>
	// </w:body>
	body := oxml.OxmlElement("w:body")
	tblE := body.CreateElement("tbl")
	tblE.Space = "w"
	trE := tblE.CreateElement("tr")
	trE.Space = "w"
	tcE := trE.CreateElement("tc")
	tcE.Space = "w"
	tcPrE := tcE.CreateElement("tcPr")
	tcPrE.Space = "w"
	tcW := tcPrE.CreateElement("tcW")
	tcW.Space = "w"
	tcW.CreateAttr("w:type", "dxa")
	tcW.CreateAttr("w:w", "4000")

	p1 := tcE.CreateElement("p")
	p1.Space = "w"
	r1 := p1.CreateElement("r")
	r1.Space = "w"
	t1 := r1.CreateElement("t")
	t1.Space = "w"
	t1.SetText("Keep this")

	p2 := tcE.CreateElement("p")
	p2.Space = "w"
	r2 := p2.CreateElement("r")
	r2.Space = "w"
	t2 := r2.CreateElement("t")
	t2.Space = "w"
	t2.SetText("[<TAG>]")

	bic := newBlockItemContainer(body, nil)
	buildFn := func(widthTwips int) ([]*etree.Element, error) {
		return []*etree.Element{oxml.OxmlElement("w:tbl")}, nil
	}

	count, err := bic.replaceTagWithElements("[<TAG>]", buildFn, 9360)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Cell should now contain: <w:tcPr>, <w:p>"Keep this"</w:p>, <w:tbl/>
	// and at least 1 <w:p> (either "Keep this" or trailing empty).
	tcChildren := tcE.ChildElements()
	hasTbl := false
	pCount := 0
	for _, child := range tcChildren {
		if child.Space == "w" && child.Tag == "tbl" {
			hasTbl = true
		}
		if child.Space == "w" && child.Tag == "p" {
			pCount++
		}
	}
	if !hasTbl {
		t.Error("expected nested <w:tbl> in cell")
	}
	if pCount < 1 {
		t.Error("expected at least 1 <w:p> in cell")
	}
}

// TestReplaceWithTable_CellWidthFallback verifies that when a cell has no
// tcW, the parent width is used as fallback.
func TestReplaceWithTable_CellWidthFallback(t *testing.T) {
	body := oxml.OxmlElement("w:body")
	tblE := body.CreateElement("tbl")
	tblE.Space = "w"
	trE := tblE.CreateElement("tr")
	trE.Space = "w"
	tcE := trE.CreateElement("tc")
	tcE.Space = "w"
	// No tcPr → no width → should fallback to parent width.
	pE := tcE.CreateElement("p")
	pE.Space = "w"
	rE := pE.CreateElement("r")
	rE.Space = "w"
	wtE := rE.CreateElement("t")
	wtE.Space = "w"
	wtE.SetText("[<TAG>]")

	bic := newBlockItemContainer(body, nil)

	var receivedWidth int
	buildFn := func(widthTwips int) ([]*etree.Element, error) {
		receivedWidth = widthTwips
		return []*etree.Element{oxml.OxmlElement("w:tbl")}, nil
	}

	_, err := bic.replaceTagWithElements("[<TAG>]", buildFn, 7777)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// Should receive the fallback parent width.
	if receivedWidth != 7777 {
		t.Errorf("buildFn received widthTwips=%d, want 7777 (fallback)", receivedWidth)
	}
}

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

// rtChildrenByTag returns all direct child elements with the given space:tag.
// Prefixed with "rt" (replace-table) to avoid collision with batch3_test.go helpers.
func rtChildrenByTag(el *etree.Element, space, tag string) []*etree.Element {
	var result []*etree.Element
	for _, c := range el.ChildElements() {
		if c.Space == space && c.Tag == tag {
			result = append(result, c)
		}
	}
	return result
}

// rtAssertCellText checks the text content of a raw <w:tc> element at (row, col)
// within a raw <w:tbl> element.
func rtAssertCellText(t *testing.T, tblEl *etree.Element, row, col int, want string) {
	t.Helper()
	trs := rtChildrenByTag(tblEl, "w", "tr")
	if row >= len(trs) {
		t.Fatalf("row %d out of range (have %d rows)", row, len(trs))
	}
	tcs := rtChildrenByTag(trs[row], "w", "tc")
	if col >= len(tcs) {
		t.Fatalf("col %d out of range (have %d cols)", col, len(tcs))
	}
	got := rtCellText(tcs[col])
	if got != want {
		t.Errorf("cell[%d][%d] text = %q, want %q", row, col, got, want)
	}
}

// rtCellText concatenates all <w:t> text inside a raw <w:tc> element.
func rtCellText(tc *etree.Element) string {
	var sb strings.Builder
	for _, p := range tc.ChildElements() {
		if p.Space == "w" && p.Tag == "p" {
			for _, r := range p.ChildElements() {
				if r.Space == "w" && r.Tag == "r" {
					for _, t := range r.ChildElements() {
						if t.Space == "w" && t.Tag == "t" {
							sb.WriteString(t.Text())
						}
					}
				}
			}
		}
	}
	return sb.String()
}

// ---------------------------------------------------------------------------
// Phase 5: Document.ReplaceWithTable integration tests
// ---------------------------------------------------------------------------

func TestDocument_ReplaceWithTable_Simple(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("[<T>]")

	td := TableData{Rows: [][]string{{"A1", "B1"}, {"A2", "B2"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	tables := mustTables(t, doc)
	if len(tables) != 1 {
		t.Fatalf("expected 1 table, got %d", len(tables))
	}
}

func TestDocument_ReplaceWithTable_WithSurroundingText(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("Before [<T>] After")

	td := TableData{Rows: [][]string{{"X"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Should have: p("Before "), tbl, p(" After")
	items := mustIterInnerContent(t, doc)
	if len(items) < 3 {
		t.Fatalf("expected at least 3 items, got %d", len(items))
	}

	// Find paragraphs text.
	var texts []string
	for _, it := range items {
		if it.IsParagraph() {
			texts = append(texts, it.Paragraph().Text())
		}
	}
	foundBefore := false
	foundAfter := false
	for _, txt := range texts {
		if strings.Contains(txt, "Before") {
			foundBefore = true
		}
		if strings.Contains(txt, "After") {
			foundAfter = true
		}
	}
	if !foundBefore {
		t.Error("surrounding text 'Before' not found")
	}
	if !foundAfter {
		t.Error("surrounding text 'After' not found")
	}

	// Should have exactly 1 table in the content.
	tblCount := 0
	for _, it := range items {
		if it.IsTable() {
			tblCount++
		}
	}
	if tblCount != 1 {
		t.Errorf("expected 1 table in content, got %d", tblCount)
	}
}

func TestDocument_ReplaceWithTable_InTableCell(t *testing.T) {
	doc := mustNewDoc(t)
	tbl, err := doc.AddTable(1, 1)
	if err != nil {
		t.Fatalf("AddTable error: %v", err)
	}
	cell, err := tbl.CellAt(0, 0)
	if err != nil {
		t.Fatalf("CellAt error: %v", err)
	}
	cell.SetText("[<T>]")

	td := TableData{Rows: [][]string{{"nested"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// The cell should now have a nested table.
	cellTables := cell.Tables()
	if len(cellTables) != 1 {
		t.Fatalf("expected 1 nested table in cell, got %d", len(cellTables))
	}

	// The cell must still have at least one <w:p> (OOXML invariant).
	cellParas := cell.Paragraphs()
	if len(cellParas) < 1 {
		t.Error("cell must have at least one paragraph (OOXML invariant)")
	}
}

func TestDocument_ReplaceWithTable_InTableCellEntireText(t *testing.T) {
	doc := mustNewDoc(t)
	tbl, err := doc.AddTable(1, 1)
	if err != nil {
		t.Fatalf("AddTable error: %v", err)
	}
	cell, err := tbl.CellAt(0, 0)
	if err != nil {
		t.Fatalf("CellAt error: %v", err)
	}
	cell.SetText("[<T>]")

	td := TableData{Rows: [][]string{{"only"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	// Cell should have: nested <w:tbl> + trailing empty <w:p>.
	cellTables := cell.Tables()
	if len(cellTables) != 1 {
		t.Errorf("expected 1 nested table, got %d", len(cellTables))
	}
	cellParas := cell.Paragraphs()
	if len(cellParas) < 1 {
		t.Error("cell must have trailing empty paragraph")
	}
}

func TestDocument_ReplaceWithTable_NestedTableWidth(t *testing.T) {
	doc := mustNewDoc(t)
	// Create a table with known cell width.
	tbl, err := doc.AddTable(1, 2)
	if err != nil {
		t.Fatalf("AddTable error: %v", err)
	}
	cell, err := tbl.CellAt(0, 0)
	if err != nil {
		t.Fatalf("CellAt error: %v", err)
	}
	cell.SetText("[<T>]")

	// Get the cell width for verification.
	cellWidth, err := cell.Width()
	if err != nil || cellWidth == nil {
		t.Skip("cell width not available — cannot verify nested table width")
	}

	td := TableData{Rows: [][]string{{"a", "b"}}}
	_, err = doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}

	// The nested table should have width equal to the cell width, not body width.
	cellTables := cell.Tables()
	if len(cellTables) != 1 {
		t.Fatalf("expected 1 nested table, got %d", len(cellTables))
	}
	nestedColWidths, err := cellTables[0].CT_Tbl().ColWidths()
	if err != nil {
		t.Fatalf("ColWidths error: %v", err)
	}
	totalNestedWidth := 0
	for _, w := range nestedColWidths {
		totalNestedWidth += w
	}
	// Nested table total width should equal cell width (each col = cellWidth/2).
	if totalNestedWidth != *cellWidth {
		t.Errorf("nested table total width = %d, want %d (cell width)", totalNestedWidth, *cellWidth)
	}

	// Also verify it's NOT the body width.
	bw, err := doc.blockWidth()
	if err != nil {
		t.Fatalf("blockWidth error: %v", err)
	}
	if totalNestedWidth == bw {
		t.Errorf("nested table width equals body width (%d) — should equal cell width", bw)
	}
}

func TestDocument_ReplaceWithTable_InHeader(t *testing.T) {
	doc := mustNewDoc(t)
	sec := doc.Sections().Iter()[0]
	hdr := sec.Header()

	// Create header definition with tag text.
	_, err := hdr.AddParagraph("[<T>]")
	if err != nil {
		t.Fatalf("AddParagraph to header error: %v", err)
	}

	td := TableData{Rows: [][]string{{"header-cell"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count < 1 {
		t.Errorf("count = %d, want >= 1", count)
	}

	// Header should now contain a table.
	hdrTables, err := hdr.Tables()
	if err != nil {
		t.Fatalf("header Tables error: %v", err)
	}
	if len(hdrTables) != 1 {
		t.Errorf("expected 1 table in header, got %d", len(hdrTables))
	}
}

func TestDocument_ReplaceWithTable_HeaderDedup(t *testing.T) {
	doc := mustNewDoc(t)

	// sec0: create header with tag.
	sec0 := doc.Sections().Iter()[0]
	_, err := sec0.Header().AddParagraph("[<T>]")
	if err != nil {
		t.Fatalf("AddParagraph to sec0 header: %v", err)
	}

	// sec1: linked to previous (shares sec0's header part).
	_, err = doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		t.Fatalf("AddSection: %v", err)
	}

	// Re-read sections: AddSectionBreak mutates the sentinel sectPr.
	sec0 = doc.Sections().Iter()[0]
	hdr0 := sec0.Header()

	td := TableData{Rows: [][]string{{"dedup"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}

	// The tag should be replaced only once (deduplication).
	hdrTables, err := hdr0.Tables()
	if err != nil {
		t.Fatalf("header Tables error: %v", err)
	}
	if len(hdrTables) != 1 {
		t.Errorf("expected exactly 1 table in shared header, got %d", len(hdrTables))
	}
	// The total count should be 1 (one header replacement, nothing in body/footer).
	if count != 1 {
		t.Errorf("total count = %d, want 1 (dedup should prevent double replacement)", count)
	}
}

func TestDocument_ReplaceWithTable_RoundTrip(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("Before [<T>] After")

	td := TableData{Rows: [][]string{{"R1C1", "R1C2"}, {"R2C1", "R2C2"}}}
	_, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}

	// Save → reopen.
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save error: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes error: %v", err)
	}

	// Verify structure.
	tables2 := mustTables(t, doc2)
	if len(tables2) != 1 {
		t.Fatalf("round-trip: expected 1 table, got %d", len(tables2))
	}
	paras2 := mustParagraphs(t, doc2)
	foundBefore := false
	foundAfter := false
	for _, p := range paras2 {
		txt := p.Text()
		if strings.Contains(txt, "Before") {
			foundBefore = true
		}
		if strings.Contains(txt, "After") {
			foundAfter = true
		}
	}
	if !foundBefore {
		t.Error("round-trip: 'Before' text lost")
	}
	if !foundAfter {
		t.Error("round-trip: 'After' text lost")
	}
}

func TestDocument_ReplaceWithTable_MultipleOccurrences(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("[<T>]")
	doc.AddParagraph("[<T>]")

	td := TableData{Rows: [][]string{{"cell"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}

	tables := mustTables(t, doc)
	if len(tables) != 2 {
		t.Errorf("expected 2 tables, got %d", len(tables))
	}
}

func TestDocument_ReplaceWithTable_MultipleInOneParagraph(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("A[<T>]B[<T>]C")

	td := TableData{Rows: [][]string{{"x"}}}
	count, err := doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}

	// Should be: p("A"), tbl, p("B"), tbl, p("C").
	items := mustIterInnerContent(t, doc)
	paraCount := 0
	tblCount := 0
	for _, it := range items {
		if it.IsParagraph() {
			paraCount++
		}
		if it.IsTable() {
			tblCount++
		}
	}
	if tblCount != 2 {
		t.Errorf("expected 2 tables, got %d", tblCount)
	}
	if paraCount < 3 {
		t.Errorf("expected at least 3 paragraphs (A, B, C), got %d", paraCount)
	}
}

func TestDocument_ReplaceWithTable_SectionBreak(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("Before [<T>] After")
	if err != nil {
		t.Fatalf("AddParagraph error: %v", err)
	}

	// Manually add a sectPr to this paragraph's pPr (simulates a section break).
	sectPr := &oxml.CT_SectPr{Element: oxml.WrapElement(oxml.OxmlElement("w:sectPr"))}
	p.CT_P().SetSectPr(sectPr)

	td := TableData{Rows: [][]string{{"x"}}}
	_, err = doc.ReplaceWithTable("[<T>]", td)
	if err != nil {
		t.Fatalf("ReplaceWithTable error: %v", err)
	}

	// After replacement: p("Before "), tbl, p(" After" + sectPr).
	// The sectPr should be on the LAST paragraph fragment.
	body, err := doc.getBody()
	if err != nil {
		t.Fatalf("getBody error: %v", err)
	}
	// Find the last <w:p> element that is a direct child of body
	// (excluding body-level sectPr).
	var lastP *etree.Element
	for _, child := range body.Element().ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			lastP = child
		}
	}
	if lastP == nil {
		t.Fatal("no paragraph found in body after replacement")
	}

	// Check that the last paragraph has sectPr in its pPr.
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

func TestDocument_ReplaceWithTable_EmptyTableData(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("[<T>]")

	td := TableData{Rows: nil}
	_, err := doc.ReplaceWithTable("[<T>]", td)
	if err == nil {
		t.Fatal("expected error for empty TableData")
	}
}

func TestDocument_ReplaceWithTable_JaggedRows(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("[<T>]")

	td := TableData{Rows: [][]string{{"a", "b"}, {"c"}}}
	_, err := doc.ReplaceWithTable("[<T>]", td)
	if err == nil {
		t.Fatal("expected error for jagged rows")
	}
	if !strings.Contains(err.Error(), "row 1 has 1 cells, expected 2") {
		t.Errorf("unexpected error message: %v", err)
	}
}

func TestDocument_ReplaceWithTable_NoMatch(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("no tags here")

	td := TableData{Rows: [][]string{{"x"}}}
	count, err := doc.ReplaceWithTable("[<MISSING>]", td)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("count = %d, want 0", count)
	}

	// Document should be unchanged.
	tables := mustTables(t, doc)
	if len(tables) != 0 {
		t.Errorf("expected 0 tables, got %d", len(tables))
	}
}

func TestDocument_ReplaceWithTable_EmptyOld(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("[<T>]")

	td := TableData{Rows: [][]string{{"x"}}}
	count, err := doc.ReplaceWithTable("", td)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("count = %d, want 0 for empty old", count)
	}
}

// ---------------------------------------------------------------------------
// Phase 5: applyToContainerDedup unit tests
// ---------------------------------------------------------------------------

func TestApplyToContainerDedup_SkipsDuplicate(t *testing.T) {
	doc := mustNewDoc(t)

	// sec0 header: create definition with tag.
	sec0 := doc.Sections().Iter()[0]
	_, err := sec0.Header().AddParagraph("[<T>]")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}

	// sec1: linked to previous → shares sec0's header StoryPart.
	_, err = doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		t.Fatalf("AddSection: %v", err)
	}

	// Re-read sections: AddSectionBreak clones the sentinel sectPr into a
	// paragraph and strips headerReferences from the sentinel. The old sec0
	// variable still points at the sentinel (now sec1), so we must refresh.
	sections := doc.Sections().Iter()
	sec0 = sections[0]
	sec1 := sections[1]

	// Call replaceWithTableDedup on both headers with shared seen map.
	seen := make(map[*parts.StoryPart]bool)
	td := TableData{Rows: [][]string{{"d"}}}
	sectWidth := sectionBlockWidth(sec0)

	n1, err := sec0.Header().baseHeaderFooter.replaceWithTableDedup("[<T>]", td, sectWidth, seen)
	if err != nil {
		t.Fatalf("replaceWithTableDedup sec0: %v", err)
	}
	// sec1 header → same StoryPart → should be skipped.
	n2, err := sec1.Header().baseHeaderFooter.replaceWithTableDedup("[<T>]", td, sectWidth, seen)
	if err != nil {
		t.Fatalf("replaceWithTableDedup sec1: %v", err)
	}

	if n1 != 1 {
		t.Errorf("sec0 header: count = %d, want 1", n1)
	}
	if n2 != 0 {
		t.Errorf("sec1 header (duplicate): count = %d, want 0", n2)
	}
}

func TestApplyToContainerDedup_SkipsLinkedToPrevious(t *testing.T) {
	doc := mustNewDoc(t)

	// Default header has no definition → IsLinkedToPrevious == true.
	sec := doc.Sections().Iter()[0]
	hdr := sec.Header()
	if !hdr.IsLinkedToPrevious() {
		t.Fatal("expected default header to be linked to previous")
	}

	// replaceWithTableDedup should return 0 (no definition to process).
	seen := make(map[*parts.StoryPart]bool)
	td := TableData{Rows: [][]string{{"x"}}}
	n, err := hdr.baseHeaderFooter.replaceWithTableDedup("[<T>]", td, 9360, seen)
	if err != nil {
		t.Fatalf("replaceWithTableDedup: %v", err)
	}
	if n != 0 {
		t.Errorf("linked-to-previous header: count = %d, want 0", n)
	}
}
