package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func TestCT_P_ParagraphText(t *testing.T) {
	// Build <w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	r1 := p.AddR()
	r1.AddTWithText("Hello ")

	r2 := p.AddR()
	r2.AddTWithText("World")

	got := p.ParagraphText()
	if got != "Hello World" {
		t.Errorf("CT_P.ParagraphText() = %q, want %q", got, "Hello World")
	}
}

func TestCT_P_ParagraphTextWithHyperlink(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	r1 := p.AddR()
	r1.AddTWithText("Click ")

	h := p.AddHyperlink()
	hr := h.AddR()
	hr.AddTWithText("here")

	r2 := p.AddR()
	r2.AddTWithText(" now")

	got := p.ParagraphText()
	if got != "Click here now" {
		t.Errorf("CT_P.ParagraphText() = %q, want %q", got, "Click here now")
	}
}

func TestCT_P_Alignment_RoundTrip(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	// Initially nil
	if a, err := p.Alignment(); err != nil {
		t.Fatalf("Alignment: %v", err)
	} else if a != nil {
		t.Error("expected nil alignment for new paragraph")
	}

	// Set center
	center := enum.WdParagraphAlignmentCenter
	if err := p.SetAlignment(&center); err != nil {
		t.Fatalf("SetAlignment: %v", err)
	}
	got, err := p.Alignment()
	if err != nil {
		t.Fatalf("Alignment: %v", err)
	}
	if got == nil || *got != enum.WdParagraphAlignmentCenter {
		t.Errorf("expected center alignment, got %v", got)
	}

	// Set nil removes
	if err := p.SetAlignment(nil); err != nil {
		t.Fatalf("SetAlignment(nil): %v", err)
	}
	if a, err := p.Alignment(); err != nil {
		t.Fatalf("Alignment: %v", err)
	} else if a != nil {
		t.Error("expected nil alignment after setting nil")
	}
}

func TestCT_P_Style_RoundTrip(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	if s, err := p.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style for new paragraph")
	}

	s := "Heading1"
	if err := p.SetStyle(&s); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	got, err := p.Style()
	if err != nil {
		t.Fatalf("Style: %v", err)
	}
	if got == nil || *got != "Heading1" {
		t.Errorf("expected Heading1 style, got %v", got)
	}

	if err := p.SetStyle(nil); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	if s, err := p.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style after removing")
	}
}

func TestCT_P_ClearContent(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	p.GetOrAddPPr() // adds pPr
	p.AddR()
	p.AddR()
	p.AddHyperlink()

	p.ClearContent()

	// pPr should remain
	if p.PPr() == nil {
		t.Error("pPr should be preserved after ClearContent")
	}
	// r and hyperlink should be gone
	if len(p.RList()) != 0 {
		t.Error("runs should be removed after ClearContent")
	}
	if len(p.HyperlinkList()) != 0 {
		t.Error("hyperlinks should be removed after ClearContent")
	}
}

func TestCT_P_AddPBefore(t *testing.T) {
	// Create a parent body with one paragraph
	body := OxmlElement("w:body")
	pEl := OxmlElement("w:p")
	body.AddChild(pEl)
	p := &CT_P{Element{e: pEl}}

	newP := p.AddPBefore()
	if newP == nil {
		t.Fatal("AddPBefore returned nil")
	}

	// The new paragraph should be before the original
	children := body.ChildElements()
	if len(children) != 2 {
		t.Fatalf("expected 2 children, got %d", len(children))
	}
	if children[0] != newP.e {
		t.Error("new paragraph should be first child")
	}
	if children[1] != p.e {
		t.Error("original paragraph should be second child")
	}
}

func TestCT_P_InnerContentElements(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{e: pEl}}

	p.GetOrAddPPr()
	p.AddR()
	p.AddHyperlink()
	p.AddR()

	elems := p.InnerContentElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 inner content elements, got %d", len(elems))
	}
	// First should be CT_R, second CT_Hyperlink, third CT_R
	if _, ok := elems[0].(*CT_R); !ok {
		t.Error("first element should be *CT_R")
	}
	if _, ok := elems[1].(*CT_Hyperlink); !ok {
		t.Error("second element should be *CT_Hyperlink")
	}
	if _, ok := elems[2].(*CT_R); !ok {
		t.Error("third element should be *CT_R")
	}
}
