package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// Mirrors Python: it_provides_access_to_the_numbering_definitions
func TestNumberingPart_Element(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
  </w:abstractNum>
</w:numbering>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/numbering.xml", opc.CTWmlNumbering, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	np := NewNumberingPart(xp)

	if np.Element() == nil {
		t.Fatal("NumberingPart Element() is nil")
	}
}

// Test NumberingPart via LoadNumberingPart constructor
func TestNumberingPart_Load(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadNumberingPart("/word/numbering.xml", opc.CTWmlNumbering, "", blob, pkg)
	if err != nil {
		t.Fatalf("LoadNumberingPart: %v", err)
	}
	np, ok := part.(*NumberingPart)
	if !ok {
		t.Fatalf("LoadNumberingPart returned %T, want *NumberingPart", part)
	}
	if np.Element() == nil {
		t.Error("loaded NumberingPart has nil element")
	}
}

func TestNumberingPart_Numbering_Valid(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	num, err := np.Numbering()
	if err != nil {
		t.Fatal(err)
	}
	if num == nil {
		t.Error("Numbering() should not return nil on valid part")
	}
}

func TestNumberingPart_Numbering_PartName(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if np.PartName() != "/word/numbering.xml" {
		t.Errorf("partname = %q, want /word/numbering.xml", np.PartName())
	}
}

func TestDefaultNumberingPart_Creates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if np.Element() == nil {
		t.Error("element is nil")
	}
	if np.PartName() != "/word/numbering.xml" {
		t.Errorf("partname = %q, want /word/numbering.xml", np.PartName())
	}
}
