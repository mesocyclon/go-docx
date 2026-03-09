package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// Mirrors Python: it_provides_access_to_its_styles
func TestStylesPart_Styles(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/styles.xml", opc.CTWmlStyles, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	sp := NewStylesPart(xp)

	styles, err := sp.Styles()
	if err != nil {
		t.Fatalf("Styles(): %v", err)
	}
	if styles == nil {
		t.Fatal("Styles() returned nil")
	}
	// Verify we can access the underlying CT_Styles
	list := styles.StyleList()
	if len(list) == 0 {
		t.Error("expected at least one style in CT_Styles")
	}
}

// Mirrors Python: it_can_construct_a_default_styles_part_to_help
func TestStylesPart_Default(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	sp, err := DefaultStylesPart(pkg)
	if err != nil {
		t.Fatalf("DefaultStylesPart: %v", err)
	}
	if sp == nil {
		t.Fatal("DefaultStylesPart returned nil")
	}
	// Verify element is valid
	if sp.Element() == nil {
		t.Error("default StylesPart has nil element")
	}
	// Verify Styles() works on default
	styles, err := sp.Styles()
	if err != nil {
		t.Fatalf("default Styles(): %v", err)
	}
	if styles == nil {
		t.Fatal("default Styles() returned nil")
	}
}

// Mirrors Python: it_is_used_by_loader_to_construct_*_part (via PartFactory)
func TestLoadStylesPart_ReturnsStylesPart(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadStylesPart("/word/styles.xml", opc.CTWmlStyles, "", blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*StylesPart); !ok {
		t.Errorf("LoadStylesPart returned %T, want *StylesPart", part)
	}
}
