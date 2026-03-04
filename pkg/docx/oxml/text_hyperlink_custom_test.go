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
