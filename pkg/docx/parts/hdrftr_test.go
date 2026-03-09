package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// Mirrors Python: it_can_create_a_new_header_part
func TestHeaderPart_New(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		t.Fatalf("NewHeaderPart: %v", err)
	}
	if hp == nil {
		t.Fatal("NewHeaderPart returned nil")
	}
	if hp.Element() == nil {
		t.Error("new HeaderPart has nil element")
	}
	// Verify the root element is w:hdr
	root := hp.Element()
	if root.Tag != "hdr" {
		t.Errorf("root tag = %q, want %q", root.Tag, "hdr")
	}
}

// Mirrors Python: it_loads_default_header_XML_from_a_template_to_help
func TestHeaderPart_TemplateHasContent(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	// Default header template should contain at least one paragraph
	children := hp.Element().ChildElements()
	hasParagraph := false
	for _, c := range children {
		if c.Tag == "p" {
			hasParagraph = true
			break
		}
	}
	if !hasParagraph {
		t.Error("default header template should contain at least one w:p element")
	}
}

// Mirrors Python: it_can_create_a_new_footer_part
func TestFooterPart_New(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	fp, err := NewFooterPart(pkg)
	if err != nil {
		t.Fatalf("NewFooterPart: %v", err)
	}
	if fp == nil {
		t.Fatal("NewFooterPart returned nil")
	}
	if fp.Element() == nil {
		t.Error("new FooterPart has nil element")
	}
	root := fp.Element()
	if root.Tag != "ftr" {
		t.Errorf("root tag = %q, want %q", root.Tag, "ftr")
	}
}

// Mirrors Python: it_loads_default_footer_XML_from_a_template_to_help
func TestFooterPart_TemplateHasContent(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	fp, err := NewFooterPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	children := fp.Element().ChildElements()
	hasParagraph := false
	for _, c := range children {
		if c.Tag == "p" {
			hasParagraph = true
			break
		}
	}
	if !hasParagraph {
		t.Error("default footer template should contain at least one w:p element")
	}
}

// Test LoadHeaderPart constructor
func TestHeaderPart_Load(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
</w:hdr>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadHeaderPart("/word/header1.xml", opc.CTWmlHeader, "", blob, pkg)
	if err != nil {
		t.Fatalf("LoadHeaderPart: %v", err)
	}
	hp, ok := part.(*HeaderPart)
	if !ok {
		t.Fatalf("LoadHeaderPart returned %T, want *HeaderPart", part)
	}
	if hp.Element() == nil {
		t.Error("loaded HeaderPart has nil element")
	}
}

// Test LoadFooterPart constructor
func TestFooterPart_Load(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
</w:ftr>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadFooterPart("/word/footer1.xml", opc.CTWmlFooter, "", blob, pkg)
	if err != nil {
		t.Fatalf("LoadFooterPart: %v", err)
	}
	fp, ok := part.(*FooterPart)
	if !ok {
		t.Fatalf("LoadFooterPart returned %T, want *FooterPart", part)
	}
	if fp.Element() == nil {
		t.Error("loaded FooterPart has nil element")
	}
}
