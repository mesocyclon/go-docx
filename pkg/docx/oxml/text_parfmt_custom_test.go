package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func TestCT_PPr_SpacingBefore_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	if sb, err := pPr.SpacingBefore(); err != nil {
		t.Fatalf("SpacingBefore: %v", err)
	} else if sb != nil {
		t.Error("expected nil spacing before for new pPr")
	}

	v := 240 // 240 twips
	if err := pPr.SetSpacingBefore(&v); err != nil {
		t.Fatalf("SetSpacingBefore: %v", err)
	}
	got, err := pPr.SpacingBefore()
	if err != nil {
		t.Fatalf("SpacingBefore: %v", err)
	}
	if got == nil || *got != 240 {
		t.Errorf("expected 240, got %v", got)
	}
}

func TestCT_PPr_SpacingAfter_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	v := 120
	if err := pPr.SetSpacingAfter(&v); err != nil {
		t.Fatalf("SetSpacingAfter: %v", err)
	}
	got, err := pPr.SpacingAfter()
	if err != nil {
		t.Fatalf("SpacingAfter: %v", err)
	}
	if got == nil || *got != 120 {
		t.Errorf("expected 120, got %v", got)
	}
}

func TestCT_PPr_SpacingLineRule(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	// Set line without lineRule → default to MULTIPLE
	line := 480
	if err := pPr.SetSpacingLine(&line); err != nil {
		t.Fatalf("SetSpacingLine: %v", err)
	}
	got, err := pPr.SpacingLineRule()
	if err != nil {
		t.Fatalf("SpacingLineRule: %v", err)
	}
	if got == nil || *got != enum.WdLineSpacingMultiple {
		t.Errorf("expected MULTIPLE default, got %v", got)
	}
}

func TestCT_PPr_IndLeft_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	if il, err := pPr.IndLeft(); err != nil {
		t.Fatalf("IndLeft: %v", err)
	} else if il != nil {
		t.Error("expected nil indent for new pPr")
	}

	v := 720 // 720 twips = 0.5 inch
	if err := pPr.SetIndLeft(&v); err != nil {
		t.Fatalf("SetIndLeft: %v", err)
	}
	got, err := pPr.IndLeft()
	if err != nil {
		t.Fatalf("IndLeft: %v", err)
	}
	if got == nil || *got != 720 {
		t.Errorf("expected 720, got %v", got)
	}
}

func TestCT_PPr_FirstLineIndent(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	// Positive first-line indent
	v := 360
	if err := pPr.SetFirstLineIndent(&v); err != nil {
		t.Fatalf("SetFirstLineIndent: %v", err)
	}
	got, err := pPr.FirstLineIndent()
	if err != nil {
		t.Fatalf("FirstLineIndent: %v", err)
	}
	if got == nil || *got != 360 {
		t.Errorf("expected 360, got %v", got)
	}

	// Negative (hanging) indent
	neg := -720
	if err := pPr.SetFirstLineIndent(&neg); err != nil {
		t.Fatalf("SetFirstLineIndent: %v", err)
	}
	got, err = pPr.FirstLineIndent()
	if err != nil {
		t.Fatalf("FirstLineIndent: %v", err)
	}
	if got == nil || *got != -720 {
		t.Errorf("expected -720 (hanging), got %v", got)
	}

	// Nil clears both
	if err := pPr.SetFirstLineIndent(nil); err != nil {
		t.Fatalf("SetFirstLineIndent: %v", err)
	}
	got, err = pPr.FirstLineIndent()
	if err != nil {
		t.Fatalf("FirstLineIndent: %v", err)
	}
	if got != nil {
		t.Errorf("expected nil after clearing, got %v", got)
	}
}

func TestCT_PPr_KeepLines_TriState(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	if pPr.KeepLinesVal() != nil {
		t.Error("expected nil keepLines for new pPr")
	}

	v := true
	if err := pPr.SetKeepLinesVal(&v); err != nil {
		t.Fatalf("SetKeepLinesVal: %v", err)
	}
	got := pPr.KeepLinesVal()
	if got == nil || !*got {
		t.Error("expected *true for keepLines")
	}

	if err := pPr.SetKeepLinesVal(nil); err != nil {
		t.Fatalf("SetKeepLinesVal: %v", err)
	}
	if pPr.KeepLinesVal() != nil {
		t.Error("expected nil after removing keepLines")
	}
}

func TestCT_PPr_PageBreakBefore(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{e: pPrEl}}

	v := true
	if err := pPr.SetPageBreakBeforeVal(&v); err != nil {
		t.Fatalf("SetPageBreakBeforeVal: %v", err)
	}
	got := pPr.PageBreakBeforeVal()
	if got == nil || !*got {
		t.Error("expected *true for pageBreakBefore")
	}
}

// --- CT_TabStops tests ---

func TestCT_TabStops_InsertTabInOrder(t *testing.T) {
	tabsEl := OxmlElement("w:tabs")
	tabs := &CT_TabStops{Element{e: tabsEl}}

	if _, err := tabs.InsertTabInOrder(2880, enum.WdTabAlignmentCenter, enum.WdTabLeaderDots); err != nil {
		t.Fatalf("InsertTabInOrder: %v", err)
	}
	if _, err := tabs.InsertTabInOrder(720, enum.WdTabAlignmentLeft, enum.WdTabLeaderSpaces); err != nil {
		t.Fatalf("InsertTabInOrder: %v", err)
	}
	if _, err := tabs.InsertTabInOrder(5760, enum.WdTabAlignmentRight, enum.WdTabLeaderDashes); err != nil {
		t.Fatalf("InsertTabInOrder: %v", err)
	}

	list := tabs.TabList()
	if len(list) != 3 {
		t.Fatalf("expected 3 tabs, got %d", len(list))
	}

	// Verify order
	pos0, _ := list[0].Pos()
	pos1, _ := list[1].Pos()
	pos2, _ := list[2].Pos()
	if pos0 != 720 || pos1 != 2880 || pos2 != 5760 {
		t.Errorf("tabs not in order: %d, %d, %d", pos0, pos1, pos2)
	}
}
