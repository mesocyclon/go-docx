package oxml

import (
	"testing"
)

func TestCT_R_AddTWithText_PreservesSpace(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	t1 := r.AddTWithText(" hello ")
	// Check xml:space="preserve" is set
	val := t1.e.SelectAttrValue("xml:space", "")
	if val != "preserve" {
		t.Errorf("expected xml:space=preserve for text with spaces, got %q", val)
	}

	r2El := OxmlElement("w:r")
	r2 := &CT_R{Element{e: r2El}}
	t2 := r2.AddTWithText("hello")
	val2 := t2.e.SelectAttrValue("xml:space", "")
	if val2 != "" {
		t.Errorf("expected no xml:space for trimmed text, got %q", val2)
	}
}

func TestCT_R_RunText(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.AddTWithText("Hello")
	r.AddTab()
	r.AddTWithText("World")

	got := r.RunText()
	if got != "Hello\tWorld" {
		t.Errorf("RunText() = %q, want %q", got, "Hello\tWorld")
	}
}

func TestCT_R_RunTextWithBr(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.AddTWithText("Line1")
	r.AddBr() // default type = textWrapping → "\n"
	r.AddTWithText("Line2")

	got := r.RunText()
	if got != "Line1\nLine2" {
		t.Errorf("RunText() = %q, want %q", got, "Line1\nLine2")
	}
}

func TestCT_R_ClearContent(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.GetOrAddRPr()
	r.AddTWithText("text")
	r.AddBr()

	r.ClearContent()

	if r.RPr() == nil {
		t.Error("rPr should be preserved after ClearContent")
	}
	if len(r.TList()) != 0 || len(r.BrList()) != 0 {
		t.Error("content should be removed after ClearContent")
	}
}

func TestCT_R_Style_RoundTrip(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	if s, err := r.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style for new run")
	}

	s := "Emphasis"
	if err := r.SetStyle(&s); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	got, err := r.Style()
	if err != nil {
		t.Fatalf("Style: %v", err)
	}
	if got == nil || *got != "Emphasis" {
		t.Errorf("expected Emphasis style, got %v", got)
	}

	if err := r.SetStyle(nil); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	if s, err := r.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style after removing")
	}
}

func TestCT_R_SetRunText(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{e: rEl}}

	r.GetOrAddRPr() // should be preserved
	r.SetRunText("Hello\tWorld\nNew")

	// Check rPr still exists
	if r.RPr() == nil {
		t.Error("rPr should be preserved after SetRunText")
	}

	got := r.RunText()
	if got != "Hello\tWorld\nNew" {
		t.Errorf("after SetRunText, RunText() = %q, want %q", got, "Hello\tWorld\nNew")
	}
}

// --- CT_Br tests ---

func TestCT_Br_TextEquivalent(t *testing.T) {
	// Default (textWrapping)
	br1 := &CT_Br{Element{e: OxmlElement("w:br")}}
	if br1.TextEquivalent() != "\n" {
		t.Error("expected newline for default break type")
	}

	// Page break
	br2 := &CT_Br{Element{e: OxmlElement("w:br")}}
	if err := br2.SetType("page"); err != nil {
		t.Fatalf("SetType: %v", err)
	}
	if br2.TextEquivalent() != "" {
		t.Error("expected empty string for page break")
	}
}
