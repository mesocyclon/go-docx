package oxml

import (
	"testing"
)

func TestCT_Hyperlink_Text(t *testing.T) {
	hEl := OxmlElement("w:hyperlink")
	h := &CT_Hyperlink{Element{e: hEl}}

	r := h.AddR()
	r.AddTWithText("Click here")

	got := h.HyperlinkText()
	if got != "Click here" {
		t.Errorf("HyperlinkText() = %q, want %q", got, "Click here")
	}
}

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
