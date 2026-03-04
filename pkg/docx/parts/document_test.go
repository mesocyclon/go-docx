package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// openDefaultDocx opens the embedded default.docx using the WML part factory.
func openDefaultDocx(t *testing.T) *opc.OpcPackage {
	t.Helper()
	docxBytes, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}
	factory := NewDocxPartFactory()
	pkg, err := opc.OpenBytes(docxBytes, factory)
	if err != nil {
		t.Fatalf("opening default.docx: %v", err)
	}
	return pkg
}

func getDocumentPart(t *testing.T, pkg *opc.OpcPackage) *DocumentPart {
	t.Helper()
	mainPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}
	dp, ok := mainPart.(*DocumentPart)
	if !ok {
		t.Fatalf("MainDocumentPart is %T, want *DocumentPart", mainPart)
	}
	return dp
}

func TestOpenDefaultDocx_DocumentPartType(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	if dp == nil {
		t.Fatal("DocumentPart is nil")
	}
}

func TestDocumentPart_Body_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	body, err := dp.Body()
	if err != nil {
		t.Fatal(err)
	}
	if body == nil {
		t.Fatal("Body is nil")
	}
}

func TestDocumentPart_StylesPart_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	sp, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp == nil {
		t.Fatal("StylesPart is nil")
	}
}

func TestDocumentPart_StylesPart_Cached(t *testing.T) {
	// In Python, _styles_part is @property (not lazyproperty), but the
	// relationship graph acts as the cache — same object returned each time.
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	sp1, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	sp2, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp1 != sp2 {
		t.Error("StylesPart should return cached instance")
	}
}

func TestDocumentPart_StylesPart_CreatesDefault(t *testing.T) {
	// Create a minimal document part with no styles relationship
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	factory := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(factory)
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	sp, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp == nil {
		t.Fatal("StylesPart should create default when absent")
	}
	// Verify it's discoverable via relationship graph (acts as cache)
	sp2, _ := dp.StylesPart()
	if sp != sp2 {
		t.Error("default StylesPart should be found via relationship graph")
	}
}

func TestDocumentPart_SettingsPart_CreatesDefault(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	sp, err := dp.SettingsPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp == nil {
		t.Fatal("SettingsPart should create default when absent")
	}
}

func TestDocumentPart_CommentsPart_CreatesDefault(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	cp, err := dp.CommentsPart()
	if err != nil {
		t.Fatal(err)
	}
	if cp == nil {
		t.Fatal("CommentsPart should create default when absent")
	}
}

func TestDocumentPart_AddHeaderPart(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	hp, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}
	if rId == "" {
		t.Error("rId is empty")
	}
	if hp == nil {
		t.Fatal("HeaderPart is nil")
	}
	if hp.Element() == nil {
		t.Error("HeaderPart element is nil")
	}
}

func TestDocumentPart_AddFooterPart(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	fp, rId, err := dp.AddFooterPart()
	if err != nil {
		t.Fatal(err)
	}
	if rId == "" {
		t.Error("rId is empty")
	}
	if fp == nil {
		t.Fatal("FooterPart is nil")
	}
	if fp.Element() == nil {
		t.Error("FooterPart element is nil")
	}
}

func TestDocumentPart_DropHeaderPart(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	_, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}

	// Verify relationship exists
	rel := dp.Rels().GetByRID(rId)
	if rel == nil {
		t.Fatal("relationship should exist before drop")
	}

	dp.DropHeaderPart(rId)

	rel = dp.Rels().GetByRID(rId)
	if rel != nil {
		t.Error("relationship should be deleted after drop (no XML refs)")
	}
}

func TestDocumentPart_HeaderPartByRID(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	hp, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}

	got, err := dp.HeaderPartByRID(rId)
	if err != nil {
		t.Fatal(err)
	}
	if got != hp {
		t.Error("HeaderPartByRID should return the same part")
	}
}

func TestDocumentPart_FooterPartByRID(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	fp, rId, err := dp.AddFooterPart()
	if err != nil {
		t.Fatal(err)
	}

	got, err := dp.FooterPartByRID(rId)
	if err != nil {
		t.Fatal(err)
	}
	if got != fp {
		t.Error("FooterPartByRID should return the same part")
	}
}

func TestDocumentPart_HeaderPartByRID_NotFound(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	_, err := dp.HeaderPartByRID("rId999")
	if err == nil {
		t.Error("expected error for non-existent rId")
	}
}

func TestDocumentPart_Styles_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	styles, err := dp.Styles()
	if err != nil {
		t.Fatal(err)
	}
	if styles == nil {
		t.Fatal("Styles should not be nil")
	}
}

func TestDocumentPart_Settings_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	settings, err := dp.Settings()
	if err != nil {
		t.Fatal(err)
	}
	if settings == nil {
		t.Fatal("Settings should not be nil")
	}
}
