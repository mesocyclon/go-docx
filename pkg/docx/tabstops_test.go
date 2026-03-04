package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// tabstops_test.go — TabStops / TabStop (Batch 1)
// Mirrors Python: tests/text/test_tabstops.py
// -----------------------------------------------------------------------

func makeTestTabStops(t *testing.T, innerXml string) *TabStops {
	t.Helper()
	pPr := makePPr(t, innerXml)
	return newTabStops(pPr)
}

// Mirrors Python: TabStops.it_can_iterate
func TestTabStops_Iter(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4680"/></w:tabs>`)
	items := ts.Iter()
	if len(items) != 2 {
		t.Fatalf("len(Iter()) = %d, want 2", len(items))
	}
}

// Mirrors Python: TabStops.it_can_get_by_index
func TestTabStops_Get(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4680"/></w:tabs>`)

	tab, err := ts.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	if tab == nil {
		t.Fatal("Get(0) returned nil")
	}

	tab2, err := ts.Get(1)
	if err != nil {
		t.Fatal(err)
	}
	if tab2 == nil {
		t.Fatal("Get(1) returned nil")
	}
}

// Mirrors Python: TabStops.it_raises_on_indexed_access_when_empty
func TestTabStops_Get_OutOfRange(t *testing.T) {
	ts := makeTestTabStops(t, ``)
	_, err := ts.Get(0)
	if err == nil {
		t.Error("expected error for Get(0) on empty TabStops")
	}
}

// Mirrors Python: TabStops.it_can_add_a_tab_stop
func TestTabStops_AddTabStop(t *testing.T) {
	ts := makeTestTabStops(t, ``)
	tab, err := ts.AddTabStop(720, enum.WdTabAlignmentLeft, enum.WdTabLeaderSpaces)
	if err != nil {
		t.Fatal(err)
	}
	if tab == nil {
		t.Fatal("AddTabStop returned nil")
	}
	if ts.Len() != 1 {
		t.Errorf("Len() after add = %d, want 1", ts.Len())
	}
}

// Mirrors Python: TabStops.it_can_delete
func TestTabStops_Delete(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4680"/></w:tabs>`)
	if ts.Len() != 2 {
		t.Fatalf("initial Len() = %d, want 2", ts.Len())
	}
	if err := ts.Delete(0); err != nil {
		t.Fatal(err)
	}
	if ts.Len() != 1 {
		t.Errorf("Len() after delete = %d, want 1", ts.Len())
	}
}

// Mirrors Python: TabStops.it_can_clear_all
func TestTabStops_ClearAll(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs>`)
	if ts.Len() != 3 {
		t.Fatalf("initial Len() = %d, want 3", ts.Len())
	}
	ts.ClearAll()
	if ts.Len() != 0 {
		t.Errorf("Len() after ClearAll = %d, want 0", ts.Len())
	}
}

// Mirrors Python: TabStop.position get/set
func TestTabStop_Position(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>`)
	tab, err := ts.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	pos, err := tab.Position()
	if err != nil {
		t.Fatal(err)
	}
	if pos != 720 {
		t.Errorf("Position() = %d, want 720", pos)
	}
}

// Mirrors Python: TabStop.alignment get/set
func TestTabStop_Alignment(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="center" w:pos="4680"/></w:tabs>`)
	tab, err := ts.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	align, err := tab.Alignment()
	if err != nil {
		t.Fatal(err)
	}
	if align != enum.WdTabAlignmentCenter {
		t.Errorf("Alignment() = %v, want CENTER", align)
	}

	if err := tab.SetAlignment(enum.WdTabAlignmentRight); err != nil {
		t.Fatal(err)
	}
	align2, _ := tab.Alignment()
	if align2 != enum.WdTabAlignmentRight {
		t.Errorf("Alignment() after set = %v, want RIGHT", align2)
	}
}

// Mirrors Python: TabStop.leader get/set
func TestTabStop_Leader(t *testing.T) {
	ts := makeTestTabStops(t, `<w:tabs><w:tab w:val="left" w:pos="720" w:leader="dot"/></w:tabs>`)
	tab, err := ts.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	leader, err := tab.Leader()
	if err != nil {
		t.Fatal(err)
	}
	if leader != enum.WdTabLeaderDots {
		t.Errorf("Leader() = %v, want DOTS", leader)
	}
}
