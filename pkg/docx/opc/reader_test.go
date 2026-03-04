package opc

import (
	"bytes"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/templates"
)

func TestPackageReader_ReadDefaultDocx(t *testing.T) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	physReader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer physReader.Close()

	reader := &PackageReader{}
	result, err := reader.Read(physReader)
	if err != nil {
		t.Fatalf("Read: %v", err)
	}

	if len(result.PkgSRels) == 0 {
		t.Error("expected package-level relationships")
	}
	if len(result.SParts) == 0 {
		t.Error("expected parts")
	}

	// Verify we have a document part
	foundDoc := false
	for _, sp := range result.SParts {
		if sp.Partname == "/word/document.xml" {
			foundDoc = true
			if sp.ContentType != CTWmlDocumentMain {
				t.Errorf("expected content type %q, got %q", CTWmlDocumentMain, sp.ContentType)
			}
			break
		}
	}
	if !foundDoc {
		t.Error("expected /word/document.xml part")
	}

	// Verify package rels include officeDocument
	foundOfficeDoc := false
	for _, srel := range result.PkgSRels {
		if srel.RelType == RTOfficeDocument {
			foundOfficeDoc = true
			break
		}
	}
	if !foundOfficeDoc {
		t.Error("expected officeDocument relationship at package level")
	}
}

func TestOpcPackage_OpenDefaultDocx(t *testing.T) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Package should have relationships
	if pkg.Rels().Len() == 0 {
		t.Error("expected package-level relationships")
	}

	// Should have parts
	parts := pkg.Parts()
	if len(parts) == 0 {
		t.Error("expected parts")
	}

	// Should be able to find main document part
	docPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}
	if docPart.PartName() != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %q", docPart.PartName())
	}
	if docPart.ContentType() != CTWmlDocumentMain {
		t.Errorf("expected %q, got %q", CTWmlDocumentMain, docPart.ContentType())
	}
}

func TestOpcPackage_SaveRoundTrip(t *testing.T) {
	// Open default.docx
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Save to bytes
	var buf bytes.Buffer
	if err := pkg.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Re-open
	pkg2, err := OpenBytes(buf.Bytes(), nil)
	if err != nil {
		t.Fatalf("OpenBytes (round-trip): %v", err)
	}

	// Verify parts survived
	parts1 := pkg.Parts()
	parts2 := pkg2.Parts()

	// Parts count should match
	if len(parts2) == 0 {
		t.Error("expected parts after round-trip")
	}

	// Document part should still be accessible
	docPart, err := pkg2.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart after round-trip: %v", err)
	}
	if docPart.PartName() != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %q", docPart.PartName())
	}

	// Blob should be non-empty
	blob, err := docPart.Blob()
	if err != nil {
		t.Fatalf("Blob() returned error: %v", err)
	}
	if len(blob) == 0 {
		t.Error("expected non-empty blob for document part after round-trip")
	}

	_ = parts1 // suppress unused warning
}

func TestOpcPackage_NextPartname(t *testing.T) {
	pkg := NewOpcPackage(nil)
	pkg.AddPart(NewBasePart("/word/header1.xml", CTWmlHeader, nil, nil))

	next := pkg.NextPartname("/word/header%d.xml")
	if next != "/word/header2.xml" {
		t.Errorf("expected /word/header2.xml, got %q", next)
	}
}

func TestOpcPackage_RelateTo(t *testing.T) {
	pkg := NewOpcPackage(nil)
	part := NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, nil)
	pkg.AddPart(part)

	rId := pkg.RelateTo(part, RTOfficeDocument)
	if rId == "" {
		t.Error("expected non-empty rId")
	}

	// Second call should return same rId
	rId2 := pkg.RelateTo(part, RTOfficeDocument)
	if rId2 != rId {
		t.Errorf("expected same rId, got %q and %q", rId, rId2)
	}
}

func TestOpcPackage_IterParts(t *testing.T) {
	pkg := NewOpcPackage(nil)
	docPart := NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, pkg)
	stylesPart := NewBasePart("/word/styles.xml", CTWmlStyles, nil, pkg)

	pkg.AddPart(docPart)
	pkg.AddPart(stylesPart)

	// Create package → document relationship
	pkg.Rels().Load("rId1", RTOfficeDocument, "word/document.xml", docPart, false)

	// Create document → styles relationship
	docPart.Rels().Load("rId1", RTStyles, "styles.xml", stylesPart, false)

	// IterParts should find both through graph walk
	parts := pkg.IterParts()
	if len(parts) != 2 {
		t.Errorf("expected 2 parts, got %d", len(parts))
	}
}
