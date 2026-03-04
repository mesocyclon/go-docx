package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func TestCT_SectPr_PageWidth_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}
	if v, err := sp.PageWidth(); err != nil {
		t.Fatalf("PageWidth: %v", err)
	} else if v != nil {
		t.Errorf("expected nil, got %d", *v)
	}
	w := 12240
	if err := sp.SetPageWidth(&w); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	got, err := sp.PageWidth()
	if err != nil {
		t.Fatalf("PageWidth: %v", err)
	}
	if got == nil || *got != 12240 {
		t.Errorf("expected 12240, got %v", got)
	}
	if err := sp.SetPageWidth(nil); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	if v, err := sp.PageWidth(); err != nil {
		t.Fatalf("PageWidth: %v", err)
	} else if v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_SectPr_PageHeight_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}
	h := 15840
	if err := sp.SetPageHeight(&h); err != nil {
		t.Fatalf("SetPageHeight: %v", err)
	}
	got, err := sp.PageHeight()
	if err != nil {
		t.Fatalf("PageHeight: %v", err)
	}
	if got == nil || *got != 15840 {
		t.Errorf("expected 15840, got %v", got)
	}
}

func TestCT_SectPr_Orientation_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}
	// Default portrait
	if o, err := sp.Orientation(); err != nil {
		t.Fatalf("Orientation: %v", err)
	} else if o != enum.WdOrientationPortrait {
		t.Error("expected portrait by default")
	}
	if err := sp.SetOrientation(enum.WdOrientationLandscape); err != nil {
		t.Fatalf("SetOrientation(landscape): %v", err)
	}
	if o, err := sp.Orientation(); err != nil {
		t.Fatalf("Orientation: %v", err)
	} else if o != enum.WdOrientationLandscape {
		t.Error("expected landscape")
	}
	if err := sp.SetOrientation(enum.WdOrientationPortrait); err != nil {
		t.Fatalf("SetOrientation(portrait): %v", err)
	}
	// After setting portrait, orient attr should be removed (default)
	pgSz := sp.PgSz()
	if pgSz != nil {
		_, ok := pgSz.GetAttr("w:orient")
		if ok {
			t.Error("expected orient attr to be removed for portrait")
		}
	}
}

func TestCT_SectPr_StartType_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}
	if st, err := sp.StartType(); err != nil {
		t.Fatalf("StartType: %v", err)
	} else if st != enum.WdSectionStartNewPage {
		t.Error("expected NEW_PAGE by default")
	}
	if err := sp.SetStartType(enum.WdSectionStartContinuous); err != nil {
		t.Fatal(err)
	}
	if st, err := sp.StartType(); err != nil {
		t.Fatalf("StartType: %v", err)
	} else if st != enum.WdSectionStartContinuous {
		t.Error("expected Continuous")
	}
	if err := sp.SetStartType(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}
	if sp.Type() != nil {
		t.Error("expected type element removed for NEW_PAGE")
	}
}

func TestCT_SectPr_TitlePg_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}
	if sp.TitlePgVal() {
		t.Error("expected false by default")
	}
	if err := sp.SetTitlePgVal(true); err != nil {
		t.Fatalf("SetTitlePgVal: %v", err)
	}
	if !sp.TitlePgVal() {
		t.Error("expected true after set")
	}
	if err := sp.SetTitlePgVal(false); err != nil {
		t.Fatalf("SetTitlePgVal: %v", err)
	}
	if sp.TitlePg() != nil {
		t.Error("expected titlePg element removed")
	}
}

func TestCT_SectPr_Margins_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}

	top := 1440
	if err := sp.SetTopMargin(&top); err != nil {
		t.Fatalf("SetTopMargin: %v", err)
	}
	got, err := sp.TopMargin()
	if err != nil {
		t.Fatalf("TopMargin: %v", err)
	}
	if got == nil || *got != 1440 {
		t.Errorf("top: expected 1440, got %v", got)
	}

	bottom := 1440
	if err := sp.SetBottomMargin(&bottom); err != nil {
		t.Fatalf("SetBottomMargin: %v", err)
	}
	if got, err := sp.BottomMargin(); err != nil {
		t.Fatalf("BottomMargin: %v", err)
	} else if got == nil || *got != 1440 {
		t.Errorf("bottom: expected 1440, got %v", got)
	}

	left := 1800
	if err := sp.SetLeftMargin(&left); err != nil {
		t.Fatalf("SetLeftMargin: %v", err)
	}
	if got, err := sp.LeftMargin(); err != nil {
		t.Fatalf("LeftMargin: %v", err)
	} else if got == nil || *got != 1800 {
		t.Errorf("left: expected 1800, got %v", got)
	}

	right := 1800
	if err := sp.SetRightMargin(&right); err != nil {
		t.Fatalf("SetRightMargin: %v", err)
	}
	if got, err := sp.RightMargin(); err != nil {
		t.Fatalf("RightMargin: %v", err)
	} else if got == nil || *got != 1800 {
		t.Errorf("right: expected 1800, got %v", got)
	}

	hdr := 720
	if err := sp.SetHeaderMargin(&hdr); err != nil {
		t.Fatalf("SetHeaderMargin: %v", err)
	}
	if got, err := sp.HeaderMargin(); err != nil {
		t.Fatalf("HeaderMargin: %v", err)
	} else if got == nil || *got != 720 {
		t.Errorf("header: expected 720, got %v", got)
	}

	ftr := 720
	if err := sp.SetFooterMargin(&ftr); err != nil {
		t.Fatalf("SetFooterMargin: %v", err)
	}
	if got, err := sp.FooterMargin(); err != nil {
		t.Fatalf("FooterMargin: %v", err)
	} else if got == nil || *got != 720 {
		t.Errorf("footer: expected 720, got %v", got)
	}

	gut := 0
	if err := sp.SetGutterMargin(&gut); err != nil {
		t.Fatalf("SetGutterMargin: %v", err)
	}
}

func TestCT_SectPr_Clone(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}
	w := 12240
	if err := sp.SetPageWidth(&w); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	sp.e.CreateAttr("w:rsidR", "00A12345")

	cloned := sp.Clone()
	// Width should be preserved
	if cw, err := cloned.PageWidth(); err != nil {
		t.Fatalf("PageWidth: %v", err)
	} else if cw == nil || *cw != 12240 {
		t.Errorf("expected cloned width 12240, got %v", cw)
	}
	// rsid should be removed
	if _, ok := cloned.GetAttr("w:rsidR"); ok {
		t.Error("expected rsid attribute to be removed in clone")
	}
	// Modifying clone shouldn't affect original
	w2 := 9999
	if err := cloned.SetPageWidth(&w2); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	if orig, err := sp.PageWidth(); err != nil {
		t.Fatalf("PageWidth: %v", err)
	} else if orig == nil || *orig != 12240 {
		t.Error("original should be unchanged")
	}
}

func TestCT_SectPr_HeaderFooterRef(t *testing.T) {
	sp := &CT_SectPr{Element{e: OxmlElement("w:sectPr")}}

	// Add header ref
	_, err := sp.AddHeaderRef(enum.WdHeaderFooterIndexPrimary, "rId1")
	if err != nil {
		t.Fatalf("AddHeaderRef: %v", err)
	}
	ref, err := sp.GetHeaderRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if ref == nil {
		t.Fatal("expected header ref")
	}
	rId, _ := ref.RId()
	if rId != "rId1" {
		t.Errorf("expected rId1, got %s", rId)
	}

	// Add footer ref
	if _, err = sp.AddFooterRef(enum.WdHeaderFooterIndexPrimary, "rId2"); err != nil {
		t.Fatalf("AddFooterRef: %v", err)
	}
	fRef, err := sp.GetFooterRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if fRef == nil {
		t.Fatal("expected footer ref")
	}

	// Remove header ref
	removed, err := sp.RemoveHeaderRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if removed != "rId1" {
		t.Errorf("expected removed rId1, got %s", removed)
	}
	refAfter, err := sp.GetHeaderRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if refAfter != nil {
		t.Error("expected header ref to be removed")
	}

	// Remove footer ref
	removedF, err := sp.RemoveFooterRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if removedF != "rId2" {
		t.Errorf("expected removed rId2, got %s", removedF)
	}
}

func TestCT_HdrFtr_InnerContentElements(t *testing.T) {
	hf := &CT_HdrFtr{Element{e: OxmlElement("w:hdr")}}
	hf.AddP()
	hf.AddTbl()
	elems := hf.InnerContentElements()
	if len(elems) != 2 {
		t.Errorf("expected 2, got %d", len(elems))
	}
}
