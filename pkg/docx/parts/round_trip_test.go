package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

func TestRoundTrip_DefaultDocx(t *testing.T) {
	// Open default.docx
	docxBytes, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}
	factory := NewDocxPartFactory()
	pkg1, err := opc.OpenBytes(docxBytes, factory)
	if err != nil {
		t.Fatalf("opening default.docx: %v", err)
	}

	// Verify we got a DocumentPart
	mainPart1, err := pkg1.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart (round 1): %v", err)
	}
	dp1, ok := mainPart1.(*DocumentPart)
	if !ok {
		t.Fatalf("MainDocumentPart is %T, want *DocumentPart", mainPart1)
	}
	body1, err := dp1.Body()
	if err != nil {
		t.Fatalf("Body (round 1): %v", err)
	}
	if body1 == nil {
		t.Fatal("Body is nil (round 1)")
	}

	// Save to bytes
	savedBytes, err := pkg1.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes: %v", err)
	}
	if len(savedBytes) == 0 {
		t.Fatal("saved bytes are empty")
	}

	// Re-open
	pkg2, err := opc.OpenBytes(savedBytes, factory)
	if err != nil {
		t.Fatalf("opening saved bytes: %v", err)
	}

	mainPart2, err := pkg2.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart (round 2): %v", err)
	}
	dp2, ok := mainPart2.(*DocumentPart)
	if !ok {
		t.Fatalf("MainDocumentPart (round 2) is %T, want *DocumentPart", mainPart2)
	}
	body2, err := dp2.Body()
	if err != nil {
		t.Fatalf("Body (round 2): %v", err)
	}
	if body2 == nil {
		t.Fatal("Body is nil (round 2)")
	}

	// Verify body XML is preserved — compare child element count as a basic check
	children1 := body1.RawElement().ChildElements()
	children2 := body2.RawElement().ChildElements()
	if len(children1) != len(children2) {
		t.Errorf("body child count: round1=%d, round2=%d", len(children1), len(children2))
	}
}

func TestRoundTrip_WithHeaderFooter(t *testing.T) {
	docxBytes, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}
	factory := NewDocxPartFactory()
	pkg1, err := opc.OpenBytes(docxBytes, factory)
	if err != nil {
		t.Fatalf("opening default.docx: %v", err)
	}

	dp := getDocumentPart(t, pkg1)

	// Add header and footer
	_, hdrRId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}
	_, ftrRId, err := dp.AddFooterPart()
	if err != nil {
		t.Fatal(err)
	}

	// Save and reopen
	savedBytes, err := pkg1.SaveToBytes()
	if err != nil {
		t.Fatal(err)
	}

	pkg2, err := opc.OpenBytes(savedBytes, factory)
	if err != nil {
		t.Fatalf("opening saved bytes: %v", err)
	}

	dp2 := getDocumentPart(t, pkg2)

	// Verify header and footer relationships survived
	hdrRel := dp2.Rels().GetByRID(hdrRId)
	if hdrRel == nil {
		t.Errorf("header relationship %q not found after round-trip", hdrRId)
	}
	ftrRel := dp2.Rels().GetByRID(ftrRId)
	if ftrRel == nil {
		t.Errorf("footer relationship %q not found after round-trip", ftrRId)
	}
}
