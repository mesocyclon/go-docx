package opc

import (
	"bytes"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/templates"
)

// ---------------------------------------------------------------------------
// Integration tests — exercises the full Open → modify → Save → re-Open
// pipeline to verify that every layer (zip I/O, rels, content types,
// part factory) works together correctly.
// ---------------------------------------------------------------------------

// TestIntegration_ModifyAndRoundTrip opens the default template, adds a
// brand-new custom XML part, saves, re-opens, and verifies the new part
// survived the round-trip.
func TestIntegration_ModifyAndRoundTrip(t *testing.T) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// --- mutate: add a custom part ---
	customXml := []byte(`<?xml version="1.0" encoding="UTF-8"?><root><value>hello</value></root>`)
	customPart := NewBasePart("/customXml/item1.xml", "application/xml", customXml, pkg)
	pkg.AddPart(customPart)
	pkg.RelateTo(customPart, "http://example.com/custom")

	// --- save ---
	saved, err := pkg.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes: %v", err)
	}
	if len(saved) == 0 {
		t.Fatal("saved bytes are empty")
	}

	// --- re-open ---
	pkg2, err := OpenBytes(saved, nil)
	if err != nil {
		t.Fatalf("OpenBytes (round-trip): %v", err)
	}

	// The original document part must survive
	docPart, err := pkg2.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart after round-trip: %v", err)
	}
	if docPart.PartName() != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %q", docPart.PartName())
	}

	// The custom part must also survive
	cp, ok := pkg2.PartByName("/customXml/item1.xml")
	if !ok {
		t.Fatal("custom part /customXml/item1.xml not found after round-trip")
	}
	blob, err := cp.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}
	if !bytes.Contains(blob, []byte("hello")) {
		t.Error("expected custom part blob to contain 'hello'")
	}
}

// TestIntegration_PartCountPreserved verifies that opening and saving
// a template produces the same number of parts.
func TestIntegration_PartCountPreserved(t *testing.T) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	pkg1, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	originalCount := len(pkg1.Parts())

	saved, err := pkg1.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes: %v", err)
	}

	pkg2, err := OpenBytes(saved, nil)
	if err != nil {
		t.Fatalf("OpenBytes (round-trip): %v", err)
	}

	if got := len(pkg2.Parts()); got != originalCount {
		t.Errorf("part count changed: %d → %d", originalCount, got)
	}
}

// TestIntegration_RelationshipPreserved checks that package-level
// relationship count is preserved through a round-trip.
func TestIntegration_RelationshipPreserved(t *testing.T) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	pkg1, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	originalRelCount := pkg1.Rels().Len()

	saved, err := pkg1.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes: %v", err)
	}

	pkg2, err := OpenBytes(saved, nil)
	if err != nil {
		t.Fatalf("OpenBytes (round-trip): %v", err)
	}

	if got := pkg2.Rels().Len(); got != originalRelCount {
		t.Errorf("package rel count changed: %d → %d", originalRelCount, got)
	}
}

// TestIntegration_BlobContentPreserved verifies that the main document part
// blob content is identical across a save/reload cycle.
func TestIntegration_BlobContentPreserved(t *testing.T) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}

	pkg1, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	docPart1, err := pkg1.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}
	blob1, err := docPart1.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}

	saved, err := pkg1.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes: %v", err)
	}

	pkg2, err := OpenBytes(saved, nil)
	if err != nil {
		t.Fatalf("OpenBytes (round-trip): %v", err)
	}
	docPart2, err := pkg2.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart (round-trip): %v", err)
	}
	blob2, err := docPart2.Blob()
	if err != nil {
		t.Fatalf("Blob (round-trip): %v", err)
	}

	if !bytes.Equal(blob1, blob2) {
		t.Error("document part blob content changed after round-trip")
	}
}

// ---------------------------------------------------------------------------
// Benchmarks
// ---------------------------------------------------------------------------

func BenchmarkOpenBytes(b *testing.B) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		b.Fatal(err)
	}
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		pkg, err := OpenBytes(data, nil)
		if err != nil {
			b.Fatal(err)
		}
		_ = pkg
	}
}

func BenchmarkSaveToBytes(b *testing.B) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		b.Fatal(err)
	}
	pkg, err := OpenBytes(data, nil)
	if err != nil {
		b.Fatal(err)
	}
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		out, err := pkg.SaveToBytes()
		if err != nil {
			b.Fatal(err)
		}
		_ = out
	}
}

func BenchmarkRoundTrip(b *testing.B) {
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		b.Fatal(err)
	}
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		pkg, err := OpenBytes(data, nil)
		if err != nil {
			b.Fatal(err)
		}
		out, err := pkg.SaveToBytes()
		if err != nil {
			b.Fatal(err)
		}
		_ = out
	}
}
