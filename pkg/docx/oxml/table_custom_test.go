package oxml

import (
	"errors"
	"strings"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func TestNewTbl_Structure(t *testing.T) {
	tbl := NewTbl(3, 4, 9360)
	// Check tblPr present
	tblPr, err := tbl.TblPr()
	if err != nil {
		t.Fatalf("TblPr error: %v", err)
	}
	if tblPr == nil {
		t.Fatal("expected tblPr, got nil")
	}
	// Check tblGrid
	grid, err := tbl.TblGrid()
	if err != nil {
		t.Fatalf("TblGrid error: %v", err)
	}
	cols := grid.GridColList()
	if len(cols) != 4 {
		t.Errorf("expected 4 gridCol, got %d", len(cols))
	}
	// Check column widths
	for _, col := range cols {
		w, err := col.W()
		if err != nil {
			t.Fatalf("W: %v", err)
		}
		if w == nil || *w != 2340 { // 9360/4
			t.Errorf("expected col width 2340, got %v", w)
		}
	}
	// Check rows
	trs := tbl.TrList()
	if len(trs) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(trs))
	}
	// Check cells per row
	for i, tr := range trs {
		tcs := tr.TcList()
		if len(tcs) != 4 {
			t.Errorf("row %d: expected 4 cells, got %d", i, len(tcs))
		}
	}
}

func TestCT_Tbl_ColCount(t *testing.T) {
	tbl := NewTbl(2, 5, 10000)
	got, err := tbl.ColCount()
	if err != nil {
		t.Fatal(err)
	}
	if got != 5 {
		t.Errorf("expected ColCount=5, got %d", got)
	}
}

func TestCT_Tbl_ColWidths(t *testing.T) {
	tbl := NewTbl(1, 3, 9000)
	widths, err := tbl.ColWidths()
	if err != nil {
		t.Fatal(err)
	}
	if len(widths) != 3 {
		t.Fatalf("expected 3 widths, got %d", len(widths))
	}
	for _, w := range widths {
		if w != 3000 {
			t.Errorf("expected 3000, got %d", w)
		}
	}
}

func TestCT_Tbl_IterTcs(t *testing.T) {
	tbl := NewTbl(2, 3, 6000)
	tcs := tbl.IterTcs()
	if len(tcs) != 6 {
		t.Errorf("expected 6 cells, got %d", len(tcs))
	}
}

func TestCT_Tbl_TblStyleVal_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	v, err := tbl.TblStyleVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "" {
		t.Errorf("expected empty, got %q", v)
	}
	if err := tbl.SetTblStyleVal("TableGrid"); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.TblStyleVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "TableGrid" {
		t.Errorf("expected TableGrid, got %q", v)
	}
	if err := tbl.SetTblStyleVal(""); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.TblStyleVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "" {
		t.Errorf("expected empty after clear, got %q", v)
	}
}

func TestCT_Tbl_Alignment_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	v, err := tbl.AlignmentVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	center := enum.WdTableAlignmentCenter
	if err := tbl.SetAlignmentVal(&center); err != nil {
		t.Fatal(err)
	}
	got, err := tbl.AlignmentVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != enum.WdTableAlignmentCenter {
		t.Errorf("expected Center, got %v", got)
	}
	if err := tbl.SetAlignmentVal(nil); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.AlignmentVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_Tbl_BidiVisualVal_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	v, err := tbl.BidiVisualVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	tr := true
	if err := tbl.SetBidiVisualVal(&tr); err != nil {
		t.Fatal(err)
	}
	got, err := tbl.BidiVisualVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != true {
		t.Errorf("expected true, got %v", got)
	}
	if err := tbl.SetBidiVisualVal(nil); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.BidiVisualVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
}

func TestCT_Tbl_Autofit_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	// Default should be true (no tblLayout means autofit)
	v, err := tbl.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if !v {
		t.Error("expected autofit=true by default")
	}
	if err := tbl.SetAutofit(false); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if v {
		t.Error("expected autofit=false after set")
	}
	if err := tbl.SetAutofit(true); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if !v {
		t.Error("expected autofit=true after reset")
	}
}

func TestCT_Row_TrIdx(t *testing.T) {
	tbl := NewTbl(3, 1, 1000)
	trs := tbl.TrList()
	for i, tr := range trs {
		if got := tr.TrIdx(); got != i {
			t.Errorf("row %d: expected TrIdx=%d, got %d", i, i, got)
		}
	}
}

func TestCT_Row_TcAtGridOffset(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tr := tbl.TrList()[0]
	tc, err := tr.TcAtGridOffset(0)
	if err != nil {
		t.Fatal(err)
	}
	if tc == nil {
		t.Fatal("expected non-nil tc at offset 0")
	}
	tc2, err := tr.TcAtGridOffset(2)
	if err != nil {
		t.Fatal(err)
	}
	if tc2 == nil {
		t.Fatal("expected non-nil tc at offset 2")
	}
	_, err = tr.TcAtGridOffset(5)
	if err == nil {
		t.Error("expected error for offset 5")
	}
}

func TestCT_Row_TrHeight_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tr := tbl.TrList()[0]
	if v, err := tr.TrHeightVal(); err != nil {
		t.Fatalf("TrHeightVal: %v", err)
	} else if v != nil {
		t.Errorf("expected nil, got %d", *v)
	}
	h := 720
	if err := tr.SetTrHeightVal(&h); err != nil {
		t.Fatalf("SetTrHeightVal: %v", err)
	}
	got, err := tr.TrHeightVal()
	if err != nil {
		t.Fatalf("TrHeightVal: %v", err)
	}
	if got == nil || *got != 720 {
		t.Errorf("expected 720, got %v", got)
	}
	rule := enum.WdRowHeightRuleExactly
	if err := tr.SetTrHeightHRule(&rule); err != nil {
		t.Fatalf("SetTrHeightHRule: %v", err)
	}
	gotRule, err := tr.TrHeightHRule()
	if err != nil {
		t.Fatalf("TrHeightHRule: %v", err)
	}
	if gotRule == nil || *gotRule != enum.WdRowHeightRuleExactly {
		t.Errorf("expected Exactly, got %v", gotRule)
	}
}

func TestNewTc(t *testing.T) {
	tc := NewTc()
	ps := tc.PList()
	if len(ps) != 1 {
		t.Errorf("expected 1 paragraph, got %d", len(ps))
	}
}

func TestCT_Tc_GridSpan_RoundTrip(t *testing.T) {
	tc := NewTc()
	v, err := tc.GridSpanVal()
	if err != nil {
		t.Fatalf("GridSpanVal: %v", err)
	}
	if v != 1 {
		t.Errorf("expected 1, got %d", v)
	}
	if err := tc.SetGridSpanVal(3); err != nil {
		t.Fatalf("SetGridSpanVal: %v", err)
	}
	v, err = tc.GridSpanVal()
	if err != nil {
		t.Fatalf("GridSpanVal: %v", err)
	}
	if v != 3 {
		t.Errorf("expected 3, got %d", v)
	}
	tc.SetGridSpanVal(1) // should remove
	v, err = tc.GridSpanVal()
	if err != nil {
		t.Fatalf("GridSpanVal: %v", err)
	}
	if v != 1 {
		t.Errorf("expected 1 after reset, got %d", v)
	}
}

func TestCT_Tc_VMerge_RoundTrip(t *testing.T) {
	tc := NewTc()
	if v := tc.VMergeVal(); v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	restart := "restart"
	if err := tc.SetVMergeVal(&restart); err != nil {
		t.Fatalf("SetVMergeVal: %v", err)
	}
	got := tc.VMergeVal()
	if got == nil || *got != "restart" {
		t.Errorf("expected restart, got %v", got)
	}
	if err := tc.SetVMergeVal(nil); err != nil {
		t.Fatalf("SetVMergeVal: %v", err)
	}
	if v := tc.VMergeVal(); v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_Tc_Width_RoundTrip(t *testing.T) {
	tc := NewTc()
	v2, err := tc.WidthTwips()
	if err != nil {
		t.Fatalf("WidthTwips: %v", err)
	}
	if v2 != nil {
		t.Errorf("expected nil, got %d", *v2)
	}
	if err := tc.SetWidthTwips(2880); err != nil {
		t.Fatalf("SetWidthTwips: %v", err)
	}
	got, err := tc.WidthTwips()
	if err != nil {
		t.Fatalf("WidthTwips: %v", err)
	}
	if got == nil || *got != 2880 {
		t.Errorf("expected 2880, got %v", got)
	}
}

func TestCT_Tc_VAlign_RoundTrip(t *testing.T) {
	tc := NewTc()
	va, err := tc.VAlignVal()
	if err != nil {
		t.Fatalf("VAlignVal: %v", err)
	}
	if va != nil {
		t.Errorf("expected nil, got %v", *va)
	}
	center := enum.WdCellVerticalAlignmentCenter
	if err := tc.SetVAlignVal(&center); err != nil {
		t.Fatalf("SetVAlignVal: %v", err)
	}
	got, err := tc.VAlignVal()
	if err != nil {
		t.Fatalf("VAlignVal: %v", err)
	}
	if got == nil || *got != enum.WdCellVerticalAlignmentCenter {
		t.Errorf("expected center, got %v", got)
	}
	if err := tc.SetVAlignVal(nil); err != nil {
		t.Fatalf("SetVAlignVal(nil): %v", err)
	}
	if v, err := tc.VAlignVal(); err != nil {
		t.Fatalf("VAlignVal: %v", err)
	} else if v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_Tc_InnerContentElements(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tc := tbl.TrList()[0].TcList()[0]
	elems := tc.InnerContentElements()
	if len(elems) != 1 {
		t.Errorf("expected 1 inner element, got %d", len(elems))
	}
	if _, ok := elems[0].(*CT_P); !ok {
		t.Error("expected first element to be *CT_P")
	}
}

func TestCT_Tc_ClearContent(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tc := tbl.TrList()[0].TcList()[0]
	tc.ClearContent()
	// Should have no p or tbl children, only tcPr
	if elems := tc.InnerContentElements(); len(elems) != 0 {
		t.Errorf("expected 0 inner elements after clear, got %d", len(elems))
	}
	if tcPr := tc.TcPr(); tcPr == nil {
		t.Error("expected tcPr to be preserved")
	}
}

func TestCT_Tc_IsEmpty(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tc := tbl.TrList()[0].TcList()[0]
	if !tc.IsEmpty() {
		t.Error("expected new cell to be empty")
	}
}

func TestCT_Tc_GridOffset(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	offsets := []int{0, 1, 2}
	for i, tc := range tcs {
		if got, err := tc.GridOffset(); err != nil {
			t.Fatalf("GridOffset: %v", err)
		} else if got != offsets[i] {
			t.Errorf("cell %d: expected offset %d, got %d", i, offsets[i], got)
		}
	}
}

func TestCT_Tc_LeftRight(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	if got, err := tcs[0].Left(); err != nil {
		t.Fatalf("Left: %v", err)
	} else if got != 0 {
		t.Errorf("expected left=0, got %d", got)
	}
	if got, err := tcs[0].Right(); err != nil {
		t.Fatalf("Right: %v", err)
	} else if got != 1 {
		t.Errorf("expected right=1, got %d", got)
	}
	if got, err := tcs[2].Right(); err != nil {
		t.Fatalf("Right: %v", err)
	} else if got != 3 {
		t.Errorf("expected right=3, got %d", got)
	}
}

func TestCT_Tc_TopBottom(t *testing.T) {
	tbl := NewTbl(2, 1, 1000)
	tcs := tbl.IterTcs()
	if got, err := tcs[0].Top(); err != nil {
		t.Fatalf("Top: %v", err)
	} else if got != 0 {
		t.Errorf("expected top=0, got %d", got)
	}
	if got, err := tcs[0].Bottom(); err != nil {
		t.Fatalf("Bottom: %v", err)
	} else if got != 1 {
		t.Errorf("expected bottom=1, got %d", got)
	}
	if got, err := tcs[1].Top(); err != nil {
		t.Fatalf("Top: %v", err)
	} else if got != 1 {
		t.Errorf("expected top=1, got %d", got)
	}
}

func TestCT_Tc_NextTc(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	next := tcs[0].NextTc()
	if next == nil {
		t.Fatal("expected next tc")
	}
	if next.e != tcs[1].e {
		t.Error("next tc should be second cell")
	}
	if last := tcs[2].NextTc(); last != nil {
		t.Error("expected nil for last cell")
	}
}

func TestCT_TblGridCol_GridColIdx(t *testing.T) {
	tbl := NewTbl(1, 4, 4000)
	grid, err := tbl.TblGrid()
	if err != nil {
		t.Fatal(err)
	}
	cols := grid.GridColList()
	for i, col := range cols {
		if got := col.GridColIdx(); got != i {
			t.Errorf("col %d: expected idx %d, got %d", i, i, got)
		}
	}
}

func TestCT_TblWidth_WidthTwips(t *testing.T) {
	tbl := NewTbl(1, 1, 2000)
	tc := tbl.TrList()[0].TcList()[0]
	tcPr := tc.TcPr()
	tcW := tcPr.TcW()
	if tcW == nil {
		t.Fatal("expected tcW")
	}
	w, err := tcW.WidthTwips()
	if err != nil {
		t.Fatalf("WidthTwips: %v", err)
	}
	if w == nil || *w != 2000 {
		t.Errorf("expected 2000, got %v", w)
	}
}

func TestCT_Tc_Merge_Horizontal(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	topTc, err := tcs[0].Merge(tcs[2])
	if err != nil {
		t.Fatal(err)
	}
	gsv, err := topTc.GridSpanVal()
	if err != nil {
		t.Fatalf("GridSpanVal: %v", err)
	}
	if gsv != 3 {
		t.Errorf("expected gridSpan=3, got %d", gsv)
	}
	// After merge, row should have only 1 tc
	remaining := tbl.TrList()[0].TcList()
	if len(remaining) != 1 {
		t.Errorf("expected 1 remaining tc, got %d", len(remaining))
	}
}

// ===========================================================================
// Error propagation tests for tcAbove / tcBelow / Top / Bottom / growTo / Merge
//
// These tests verify that malformed XML attributes in neighboring cells
// are surfaced as errors rather than silently swallowed. Before the fix,
// tcAbove/tcBelow returned nil on error, causing Top/Bottom to fall back
// to incorrect row indices without any error indication.
// ===========================================================================

// corruptGridSpan adds a w:gridSpan child with a non-numeric w:val to the
// cell's tcPr, causing GridSpanVal() to return a ParseAttrError.
func corruptGridSpan(tc *CT_Tc) {
	tcPr := tc.GetOrAddTcPr()
	gs := tcPr.GetOrAddGridSpan()
	gs.e.CreateAttr("w:val", "CORRUPT")
}

// requireParseAttrError asserts that err wraps a *ParseAttrError.
func requireParseAttrError(t *testing.T, err error) {
	t.Helper()
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	var pe *ParseAttrError
	if !errors.As(err, &pe) {
		t.Fatalf("expected *ParseAttrError in chain, got: %v", err)
	}
}

// requireErrorContains asserts err is non-nil and its message contains substr.
func requireErrorContains(t *testing.T, err error, substr string) {
	t.Helper()
	if err == nil {
		t.Fatalf("expected error containing %q, got nil", substr)
	}
	if !strings.Contains(err.Error(), substr) {
		t.Errorf("error %q does not contain %q", err.Error(), substr)
	}
}

// ---------------------------------------------------------------------------
// Helper: build a 2-row, 2-col table with vertical merge on column 0
//
//	Row 0: [A (vMerge=restart)] [B]
//	Row 1: [A (vMerge=continue)] [B]
// ---------------------------------------------------------------------------

func newVMergedTable() *CT_Tbl {
	tbl := NewTbl(2, 2, 4000)
	trs := tbl.TrList()
	r0c0 := trs[0].TcList()[0]
	r1c0 := trs[1].TcList()[0]
	restart := "restart"
	r0c0.SetVMergeVal(&restart)
	cont := "continue"
	r1c0.SetVMergeVal(&cont)
	return tbl
}

// ---------------------------------------------------------------------------
// Happy-path: verify Top/Bottom work correctly on a valid vMerge table
// ---------------------------------------------------------------------------

func TestTopBottom_VMerge_Valid(t *testing.T) {
	tbl := newVMergedTable()
	trs := tbl.TrList()
	r0c0 := trs[0].TcList()[0]
	r1c0 := trs[1].TcList()[0]

	// r0c0: vMerge=restart → Top=0
	top, err := r0c0.Top()
	if err != nil {
		t.Fatalf("r0c0.Top: %v", err)
	}
	if top != 0 {
		t.Errorf("r0c0.Top = %d, want 0", top)
	}

	// r0c0: vMerge=restart, below is continue → Bottom=2
	bot, err := r0c0.Bottom()
	if err != nil {
		t.Fatalf("r0c0.Bottom: %v", err)
	}
	if bot != 2 {
		t.Errorf("r0c0.Bottom = %d, want 2", bot)
	}

	// r1c0: vMerge=continue → Top follows tcAbove → 0
	top, err = r1c0.Top()
	if err != nil {
		t.Fatalf("r1c0.Top: %v", err)
	}
	if top != 0 {
		t.Errorf("r1c0.Top = %d, want 0", top)
	}

	// r1c0: vMerge=continue, no row below → Bottom=2
	bot, err = r1c0.Bottom()
	if err != nil {
		t.Fatalf("r1c0.Bottom: %v", err)
	}
	if bot != 2 {
		t.Errorf("r1c0.Bottom = %d, want 2", bot)
	}
}

// ---------------------------------------------------------------------------
// Top() propagates error when tcAbove fails due to corrupt GridOffset
// (preceding sibling in the same row has corrupt gridSpan)
// ---------------------------------------------------------------------------

func TestTop_PropagatesError_CorruptGridOffset(t *testing.T) {
	// 2 rows, 3 cols so cell[1] has a preceding sibling cell[0]
	tbl := NewTbl(2, 3, 6000)
	trs := tbl.TrList()

	// Set vMerge=continue on row1, cell[1] so Top() calls tcAbove
	cont := "continue"
	r1c1 := trs[1].TcList()[1]
	r1c1.SetVMergeVal(&cont)

	// Corrupt gridSpan of row1, cell[0] — the preceding sibling.
	// tcAbove → GridOffset → iterates preceding siblings → hits corrupt val.
	r1c0 := trs[1].TcList()[0]
	corruptGridSpan(r1c0)

	_, err := r1c1.Top()
	requireParseAttrError(t, err)
	requireErrorContains(t, err, "tcAbove")
	requireErrorContains(t, err, "Top")
}

// ---------------------------------------------------------------------------
// Top() propagates error when tcAbove fails due to corrupt TcAtGridOffset
// (cell in the row above has corrupt gridSpan)
// ---------------------------------------------------------------------------

func TestTop_PropagatesError_CorruptTcAtGridOffset(t *testing.T) {
	// 2 rows, 3 cols
	// Row 0: [cell0_corrupt] [cell1] [cell2]
	// Row 1: [cell0] [cell1 vMerge=continue] [cell2]
	//
	// r1c1.Top() → tcAbove → GridOffset on r1c1 = 1 (normal) →
	// TcAtGridOffset(1) on row0 → iterates cells → calls GridSpanVal
	// on r0c0 → CORRUPT → error.
	tbl := NewTbl(2, 3, 6000)
	trs := tbl.TrList()

	// Set vMerge=continue on row1, cell[1] so Top() calls tcAbove
	cont := "continue"
	r1c1 := trs[1].TcList()[1]
	r1c1.SetVMergeVal(&cont)

	// Corrupt gridSpan of row0, cell[0] — TcAtGridOffset will iterate
	// past it when looking for offset 1 and hit the corrupt value.
	r0c0 := trs[0].TcList()[0]
	corruptGridSpan(r0c0)

	_, err := r1c1.Top()
	requireParseAttrError(t, err)
	requireErrorContains(t, err, "tcAbove")
}

// ---------------------------------------------------------------------------
// Bottom() propagates error when tcBelow fails
// ---------------------------------------------------------------------------

func TestBottom_PropagatesError_CorruptTcBelow(t *testing.T) {
	// 3 rows, 2 cols — vMerge spans all 3 rows in column 1
	tbl := NewTbl(3, 2, 4000)
	trs := tbl.TrList()

	restart := "restart"
	trs[0].TcList()[1].SetVMergeVal(&restart)
	cont := "continue"
	trs[1].TcList()[1].SetVMergeVal(&cont)
	trs[2].TcList()[1].SetVMergeVal(&cont)

	// Corrupt gridSpan in row1, cell[0] (column 0).
	// r0c1.Bottom() → tcBelow → GridOffset on r0c1 = 1 →
	// TcAtGridOffset(1) on row1 → iterates cells → GridSpanVal on
	// row1.cell[0] → CORRUPT → error.
	corruptGridSpan(trs[1].TcList()[0])

	_, err := trs[0].TcList()[1].Bottom()
	requireParseAttrError(t, err)
	requireErrorContains(t, err, "tcBelow")
	requireErrorContains(t, err, "Bottom")
}

// ---------------------------------------------------------------------------
// growTo() propagates error when tcBelow fails (instead of "not enough rows")
// ---------------------------------------------------------------------------

func TestGrowTo_PropagatesError_CorruptTcBelow(t *testing.T) {
	// 2 rows, 2 cols
	tbl := NewTbl(2, 2, 4000)
	trs := tbl.TrList()

	// Corrupt gridSpan in row1, cell[0] (column 0).
	// growTo on r0c1 with height=2 → tcBelow → GridOffset on r0c1 = 1 →
	// TcAtGridOffset(1) on row1 → GridSpanVal on row1.cell[0] → CORRUPT.
	corruptGridSpan(trs[1].TcList()[0])

	r0c1 := trs[0].TcList()[1]
	err := r0c1.growTo(1, 2, r0c1)
	requireParseAttrError(t, err)
	requireErrorContains(t, err, "growTo")
	requireErrorContains(t, err, "tcBelow")

	// Verify the error is NOT the misleading "not enough rows" message
	if strings.Contains(err.Error(), "not enough rows") {
		t.Error("error should describe the real cause, not 'not enough rows'")
	}
}

// ---------------------------------------------------------------------------
// Merge() propagates error from spanDimensions → Top/Bottom
// ---------------------------------------------------------------------------

func TestMerge_PropagatesError_CorruptNeighbor(t *testing.T) {
	// 2 rows, 3 cols
	// Row 0: [A] [B vMerge=restart] [C]
	// Row 1: [D] [E vMerge=continue] [F]
	//
	// Corrupt A's gridSpan. Merging E with F triggers spanDimensions →
	// E.Top() → vMerge=continue → tcAbove → TcAtGridOffset(1) on row0 →
	// iterates past A → GridSpanVal on A → CORRUPT → error.
	tbl := NewTbl(2, 3, 6000)
	trs := tbl.TrList()

	restart := "restart"
	trs[0].TcList()[1].SetVMergeVal(&restart)
	cont := "continue"
	trs[1].TcList()[1].SetVMergeVal(&cont)

	// Corrupt row0, cell[0] — TcAtGridOffset will iterate past it
	corruptGridSpan(trs[0].TcList()[0])

	r1c1 := trs[1].TcList()[1]
	r1c2 := trs[1].TcList()[2]
	_, err := r1c1.Merge(r1c2)
	if err == nil {
		t.Fatal("expected error from Merge, got nil")
	}
	requireParseAttrError(t, err)
}

// ---------------------------------------------------------------------------
// tcAbove / tcBelow return (nil, nil) for boundary rows (no error)
// ---------------------------------------------------------------------------

func TestTcAbove_TopRow_ReturnsNilNil(t *testing.T) {
	tbl := NewTbl(2, 1, 1000)
	r0c0 := tbl.TrList()[0].TcList()[0]

	above, err := r0c0.tcAbove()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if above != nil {
		t.Error("expected nil for top row cell")
	}
}

func TestTcBelow_BottomRow_ReturnsNilNil(t *testing.T) {
	tbl := NewTbl(2, 1, 1000)
	r1c0 := tbl.TrList()[1].TcList()[0]

	below, err := r1c0.tcBelow()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if below != nil {
		t.Error("expected nil for bottom row cell")
	}
}

// ---------------------------------------------------------------------------
// tcAbove / tcBelow return valid cell for interior rows
// ---------------------------------------------------------------------------

func TestTcAbove_ReturnsCell(t *testing.T) {
	tbl := NewTbl(3, 1, 1000)
	trs := tbl.TrList()
	r1c0 := trs[1].TcList()[0]

	above, err := r1c0.tcAbove()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if above == nil {
		t.Fatal("expected non-nil cell above")
	}
	if above.e != trs[0].TcList()[0].e {
		t.Error("tcAbove returned wrong cell")
	}
}

func TestTcBelow_ReturnsCell(t *testing.T) {
	tbl := NewTbl(3, 1, 1000)
	trs := tbl.TrList()
	r1c0 := trs[1].TcList()[0]

	below, err := r1c0.tcBelow()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if below == nil {
		t.Fatal("expected non-nil cell below")
	}
	if below.e != trs[2].TcList()[0].e {
		t.Error("tcBelow returned wrong cell")
	}
}

// ---------------------------------------------------------------------------
// Vertical merge happy path: Merge across 2 rows works correctly
// ---------------------------------------------------------------------------

func TestMerge_Vertical_Valid(t *testing.T) {
	tbl := NewTbl(2, 2, 4000)
	trs := tbl.TrList()
	r0c0 := trs[0].TcList()[0]
	r1c0 := trs[1].TcList()[0]

	topTc, err := r0c0.Merge(r1c0)
	if err != nil {
		t.Fatalf("Merge: %v", err)
	}

	// Top cell should have vMerge=restart
	vm := topTc.VMergeVal()
	if vm == nil || *vm != "restart" {
		t.Errorf("expected vMerge=restart on top cell, got %v", vm)
	}

	// Bottom cell should have vMerge=continue
	r1c0After := trs[1].TcList()[0]
	vm = r1c0After.VMergeVal()
	if vm == nil || *vm != "continue" {
		t.Errorf("expected vMerge=continue on bottom cell, got %v", vm)
	}
}
