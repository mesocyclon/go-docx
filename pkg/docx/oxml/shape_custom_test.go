package oxml

import (
	"testing"
)

func TestNewPicInline_Structure(t *testing.T) {
	inline, err := NewPicInline(1, "rId5", "image1.png", 914400, 457200)
	if err != nil {
		t.Fatal(err)
	}

	// Check extent dimensions
	cx, err := inline.ExtentCx()
	if err != nil {
		t.Fatalf("ExtentCx error: %v", err)
	}
	if cx != 914400 {
		t.Errorf("expected cx=914400, got %d", cx)
	}
	cy, err := inline.ExtentCy()
	if err != nil {
		t.Fatalf("ExtentCy error: %v", err)
	}
	if cy != 457200 {
		t.Errorf("expected cy=457200, got %d", cy)
	}

	// Check docPr
	docPr, err := inline.DocPr()
	if err != nil {
		t.Fatalf("DocPr error: %v", err)
	}
	id, err := docPr.Id()
	if err != nil {
		t.Fatalf("docPr.Id() error: %v", err)
	}
	if id != 1 {
		t.Errorf("expected docPr id=1, got %d", id)
	}
	name, err := docPr.Name()
	if err != nil {
		t.Fatalf("docPr.Name() error: %v", err)
	}
	if name != "Picture 1" {
		t.Errorf("expected docPr name='Picture 1', got %q", name)
	}

	// Check graphic data URI
	graphic, err := inline.Graphic()
	if err != nil {
		t.Fatalf("Graphic error: %v", err)
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		t.Fatalf("GraphicData error: %v", err)
	}
	uri, err := gd.Uri()
	if err != nil {
		t.Fatalf("graphicData.Uri() error: %v", err)
	}
	if uri != "http://schemas.openxmlformats.org/drawingml/2006/picture" {
		t.Errorf("unexpected graphicData uri: %q", uri)
	}

	// Check pic element exists in graphicData
	pic := gd.Pic()
	if pic == nil {
		t.Fatal("expected pic:pic inside graphicData, got nil")
	}

	// Check blipFill has the right rId
	blipFill, err := pic.BlipFill()
	if err != nil {
		t.Fatalf("BlipFill error: %v", err)
	}
	embed := blipFill.Blip().Embed()
	if embed != "rId5" {
		t.Errorf("expected blip embed='rId5', got %q", embed)
	}
}

func TestCT_Inline_SetExtent(t *testing.T) {
	inline, err := NewPicInline(1, "rId1", "test.png", 100, 200)
	if err != nil {
		t.Fatal(err)
	}
	if err := inline.SetExtentCx(300); err != nil {
		t.Fatal(err)
	}
	if err := inline.SetExtentCy(400); err != nil {
		t.Fatal(err)
	}
	cx, err := inline.ExtentCx()
	if err != nil {
		t.Fatal(err)
	}
	if cx != 300 {
		t.Errorf("expected cx=300, got %d", cx)
	}
	cy, err := inline.ExtentCy()
	if err != nil {
		t.Fatal(err)
	}
	if cy != 400 {
		t.Errorf("expected cy=400, got %d", cy)
	}
}

func TestCT_ShapeProperties_CxCy(t *testing.T) {
	// Parse a pic:spPr with xfrm/ext
	xml := `<pic:spPr xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>`
	el, err := ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	spPr := &CT_ShapeProperties{Element{e: el}}

	cx, err := spPr.Cx()
	if err != nil {
		t.Fatalf("Cx: %v", err)
	}
	if cx == nil || *cx != 914400 {
		t.Errorf("expected cx=914400, got %v", cx)
	}
	cy, err := spPr.Cy()
	if err != nil {
		t.Fatalf("Cy: %v", err)
	}
	if cy == nil || *cy != 457200 {
		t.Errorf("expected cy=457200, got %v", cy)
	}

	// Set new values
	if err := spPr.SetCx(1234); err != nil {
		t.Fatalf("SetCx: %v", err)
	}
	if err := spPr.SetCy(5678); err != nil {
		t.Fatalf("SetCy: %v", err)
	}
	cx, err = spPr.Cx()
	if err != nil {
		t.Fatalf("Cx: %v", err)
	}
	if cx == nil || *cx != 1234 {
		t.Errorf("after set, expected cx=1234, got %v", cx)
	}
	cy, err = spPr.Cy()
	if err != nil {
		t.Fatalf("Cy: %v", err)
	}
	if cy == nil || *cy != 5678 {
		t.Errorf("after set, expected cy=5678, got %v", cy)
	}
}

func TestCT_ShapeProperties_CxCy_NoXfrm(t *testing.T) {
	xml := `<pic:spPr xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></pic:spPr>`
	el, _ := ParseXml([]byte(xml))
	spPr := &CT_ShapeProperties{Element{e: el}}

	if cx, err := spPr.Cx(); err != nil {
		t.Fatalf("Cx: %v", err)
	} else if cx != nil {
		t.Errorf("expected nil cx on empty spPr, got %v", cx)
	}
	if cy, err := spPr.Cy(); err != nil {
		t.Fatalf("Cy: %v", err)
	} else if cy != nil {
		t.Errorf("expected nil cy on empty spPr, got %v", cy)
	}
}
