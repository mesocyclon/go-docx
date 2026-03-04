package oxml

import (
	"testing"
)

func TestCT_Body_InnerContentElements(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/><w:tbl/><w:p/><w:sectPr/>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{e: el}}

	elems := body.InnerContentElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 elements (sectPr excluded), got %d", len(elems))
	}
	if _, ok := elems[0].(*CT_P); !ok {
		t.Error("expected first element to be CT_P")
	}
	if _, ok := elems[1].(*CT_Tbl); !ok {
		t.Error("expected second element to be CT_Tbl")
	}
	if _, ok := elems[2].(*CT_P); !ok {
		t.Error("expected third element to be CT_P")
	}
}

func TestCT_Body_ClearContent(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/><w:tbl/><w:p/><w:sectPr/>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{e: el}}

	body.ClearContent()

	// Only sectPr should remain
	children := body.e.ChildElements()
	if len(children) != 1 {
		t.Fatalf("expected 1 child (sectPr), got %d", len(children))
	}
	if children[0].Tag != "sectPr" {
		t.Errorf("expected remaining child to be sectPr, got %s", children[0].Tag)
	}
}

func TestCT_Body_ClearContent_NoSectPr(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/><w:tbl/>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{e: el}}

	body.ClearContent()
	if len(body.e.ChildElements()) != 0 {
		t.Errorf("expected 0 children, got %d", len(body.e.ChildElements()))
	}
}

func TestCT_Document_SectPrList(t *testing.T) {
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` +
		`<w:p/>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` +
		`<w:sectPr/>` +
		`</w:body>` +
		`</w:document>`
	el, _ := ParseXml([]byte(xml))
	doc := &CT_Document{Element{e: el}}

	sectPrs := doc.SectPrList()
	if len(sectPrs) != 3 {
		t.Errorf("expected 3 sectPr elements, got %d", len(sectPrs))
	}
}

func TestCT_Body_AddSectionBreak(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/>` +
		`<w:sectPr><w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{e: el}}

	sentinelSectPr := body.AddSectionBreak()
	if sentinelSectPr == nil {
		t.Fatal("expected sentinel sectPr, got nil")
	}

	// The sentinel should have no headerReference now (removed)
	refs := sentinelSectPr.HeaderReferenceList()
	if len(refs) != 0 {
		t.Errorf("expected 0 headerReferences on sentinel, got %d", len(refs))
	}

	// There should be a new paragraph with sectPr (the clone)
	pList := body.PList()
	// Should be at least 2 paragraphs now (original + new one with sectPr)
	if len(pList) < 2 {
		t.Fatalf("expected at least 2 paragraphs, got %d", len(pList))
	}
}

// --- Factory method tests ---

func TestNewDecimalNumber(t *testing.T) {
	dn, err := NewDecimalNumber("w:abstractNumId", 42)
	if err != nil {
		t.Fatalf("NewDecimalNumber: %v", err)
	}
	v, err := dn.Val()
	if err != nil {
		t.Fatal(err)
	}
	if v != 42 {
		t.Errorf("expected val=42, got %d", v)
	}
}

func TestNewCtString(t *testing.T) {
	s, err := NewCtString("w:pStyle", "Heading1")
	if err != nil {
		t.Fatalf("NewCtString: %v", err)
	}
	v, err := s.Val()
	if err != nil {
		t.Fatal(err)
	}
	if v != "Heading1" {
		t.Errorf("expected val='Heading1', got %q", v)
	}
}
