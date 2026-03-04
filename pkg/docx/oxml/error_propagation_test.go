package oxml

import (
	"errors"
	"strings"
	"testing"
)

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
