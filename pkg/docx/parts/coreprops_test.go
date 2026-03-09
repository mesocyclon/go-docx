package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

func TestCorePropertiesPart_CT_ValidElement(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	ct, err := cp.CT()
	if err != nil {
		t.Fatal(err)
	}
	if ct == nil {
		t.Fatal("CT() should return non-nil")
	}
}

func TestCorePropertiesPart_CT_PartName(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if cp.PartName() != "/docProps/core.xml" {
		t.Errorf("partname = %q, want /docProps/core.xml", cp.PartName())
	}
}

func TestDefaultCorePropertiesPart_HasMetadata(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if cp.Element() == nil {
		t.Fatal("element is nil")
	}
	// Verify partname
	if cp.PartName() != "/docProps/core.xml" {
		t.Errorf("partname = %q, want /docProps/core.xml", cp.PartName())
	}
}

func TestCoreProperties_CreatesDefault(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	cp, err := dp.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	if cp == nil {
		t.Fatal("CoreProperties returned nil")
	}
	ct, err := cp.CT()
	if err != nil {
		t.Fatal(err)
	}
	if ct == nil {
		t.Error("CT() returned nil on default core properties")
	}
}

func TestCoreProperties_ReturnsExisting(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	cp1, err := dp.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	cp2, err := dp.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	if cp1 != cp2 {
		t.Error("CoreProperties should return same instance via package rel cache")
	}
}
