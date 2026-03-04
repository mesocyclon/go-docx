package opc

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

// buildTestZip creates a minimal .docx ZIP in memory from a map of member
// names to contents.  This allows tests to craft packages with intentional
// structural defects (missing parts, incomplete content types, etc.).
func buildTestZip(t *testing.T, members map[string]string) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, body := range members {
		fw, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %q: %v", name, err)
		}
		if _, err := fw.Write([]byte(body)); err != nil {
			t.Fatalf("zip write %q: %v", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return buf.Bytes()
}

// minimalContentTypes has only the document.xml override and xml/rels defaults.
const minimalContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

// minimalDocumentXml is a bare-minimum w:document element.
const minimalDocumentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>`

// ---------------------------------------------------------------------------
// Reader-level tests (walkParts)
// ---------------------------------------------------------------------------

// TestWalkParts_SkipsMissingMember verifies that a relationship pointing to a
// ZIP member that does not exist is silently skipped at the reader level
// (the part is not added to SParts), rather than causing a hard failure.
func TestWalkParts_SkipsMissingMember(t *testing.T) {
	t.Parallel()

	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/footer1.xml"/>
</Relationships>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml": minimalContentTypes,
		"_rels/.rels":         pkgRels,
		"word/document.xml":   minimalDocumentXml,
		// word/footer1.xml intentionally absent
	})

	physReader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer physReader.Close()

	reader := &PackageReader{}
	result, err := reader.Read(physReader)
	if err != nil {
		t.Fatalf("Read should succeed with dangling rel, got: %v", err)
	}

	// Only document.xml should have been loaded as a part.
	if len(result.SParts) != 1 {
		t.Fatalf("expected 1 serialized part, got %d", len(result.SParts))
	}
	if result.SParts[0].Partname != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %q", result.SParts[0].Partname)
	}
}

// TestWalkParts_SkipsMissingContentType verifies that a part physically present
// in the ZIP but with a truly unknown extension (not in [Content_Types].xml and
// not in the well-known fallback table) is skipped at the reader level.
func TestWalkParts_SkipsMissingContentType(t *testing.T) {
	t.Parallel()

	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="word/media/data.xyz99"/>
</Relationships>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml":   minimalContentTypes, // no Default for "xyz99"
		"_rels/.rels":           pkgRels,
		"word/document.xml":     minimalDocumentXml,
		"word/media/data.xyz99": "unknown-format-bytes",
	})

	physReader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer physReader.Close()

	reader := &PackageReader{}
	result, err := reader.Read(physReader)
	if err != nil {
		t.Fatalf("Read should succeed with missing content type, got: %v", err)
	}

	if len(result.SParts) != 1 {
		t.Fatalf("expected 1 serialized part (only document.xml), got %d", len(result.SParts))
	}
}

// ---------------------------------------------------------------------------
// Full Open path tests — dangling rels must be PRESERVED (not dropped)
// ---------------------------------------------------------------------------

// TestOpenBytes_DanglingPartLevelRel_Preserved verifies that when a part-level
// .rels references a missing ZIP member, the relationship is preserved with
// TargetPart=nil so that the .rels file survives a round-trip.
// This is the fix for the "file doesn't open without repair dialog" bug.
func TestOpenBytes_DanglingPartLevelRel_Preserved(t *testing.T) {
	t.Parallel()

	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`

	// document.xml.rels references fontTable.xml which is MISSING from ZIP.
	docRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    Target="fontTable.xml"/>
</Relationships>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml":          minimalContentTypes,
		"_rels/.rels":                  pkgRels,
		"word/document.xml":            minimalDocumentXml,
		"word/_rels/document.xml.rels": docRels,
		// word/fontTable.xml intentionally absent
	})

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes should succeed, got: %v", err)
	}

	docPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}

	// The dangling rId1 must be preserved on the document part.
	rel := docPart.Rels().GetByRID("rId1")
	if rel == nil {
		t.Fatal("dangling rId1 relationship was dropped — must be preserved")
	}
	if rel.TargetPart != nil {
		t.Error("dangling rel should have nil TargetPart")
	}
	if rel.TargetRef != "fontTable.xml" {
		t.Errorf("expected TargetRef %q, got %q", "fontTable.xml", rel.TargetRef)
	}
	if rel.RelType != RTFontTable {
		t.Errorf("expected RelType %q, got %q", RTFontTable, rel.RelType)
	}
}

// TestOpenBytes_DanglingPackageLevelRel_Preserved verifies that a dangling
// package-level relationship is preserved with TargetPart=nil.
func TestOpenBytes_DanglingPackageLevelRel_Preserved(t *testing.T) {
	t.Parallel()

	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
    Target="docProps/thumbnail.jpeg"/>
</Relationships>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml": minimalContentTypes,
		"_rels/.rels":         pkgRels,
		"word/document.xml":   minimalDocumentXml,
		// docProps/thumbnail.jpeg intentionally absent
	})

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes should succeed, got: %v", err)
	}

	// Document part must be accessible.
	docPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}
	if docPart.PartName() != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %q", docPart.PartName())
	}

	// The dangling rId2 (thumbnail) must be preserved at package level.
	rel := pkg.Rels().GetByRID("rId2")
	if rel == nil {
		t.Fatal("dangling rId2 relationship was dropped — must be preserved")
	}
	if rel.TargetPart != nil {
		t.Error("dangling rel should have nil TargetPart")
	}
	if rel.TargetRef != "docProps/thumbnail.jpeg" {
		t.Errorf("expected TargetRef %q, got %q", "docProps/thumbnail.jpeg", rel.TargetRef)
	}
}

// ---------------------------------------------------------------------------
// Round-trip test — the critical scenario from the bug report
// ---------------------------------------------------------------------------

// TestRoundTrip_DanglingRel_RelsFilePreserved verifies the full Open → Save
// round-trip: a part-level .rels with a dangling reference must be written
// back so that Word doesn't show a repair dialog.
//
// This reproduces the exact bug: header2.xml has r:id="rId1" referencing an
// image via its .rels, but the image is missing from the ZIP.  After round-trip,
// the .rels MUST still be present — without it Word can't resolve rId1.
func TestRoundTrip_DanglingRel_RelsFilePreserved(t *testing.T) {
	t.Parallel()

	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`

	headerContentType := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`

	docRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    Target="header1.xml"/>
</Relationships>`

	// header1.xml.rels: rId1 → image that does NOT exist in ZIP.
	headerRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
</Relationships>`

	headerXml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p><w:r><w:t>Header with image ref</w:t></w:r></w:p>
</w:hdr>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml":          headerContentType,
		"_rels/.rels":                  pkgRels,
		"word/document.xml":            minimalDocumentXml,
		"word/_rels/document.xml.rels": docRels,
		"word/header1.xml":             headerXml,
		"word/_rels/header1.xml.rels":  headerRels,
		// word/media/image1.png intentionally absent
	})

	// Open
	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Save
	var buf bytes.Buffer
	if err := pkg.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Inspect the output ZIP
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("reading output zip: %v", err)
	}

	memberMap := make(map[string]*zip.File)
	for _, f := range zr.File {
		memberMap[f.Name] = f
	}

	// The header's .rels file MUST be present.
	relsFile, ok := memberMap["word/_rels/header1.xml.rels"]
	if !ok {
		var names []string
		for n := range memberMap {
			names = append(names, n)
		}
		t.Fatalf("word/_rels/header1.xml.rels missing from output ZIP.\nMembers: %v", names)
	}

	// Read the .rels content and verify the dangling relationship survived.
	rc, err := relsFile.Open()
	if err != nil {
		t.Fatalf("opening rels file: %v", err)
	}
	var relsBuf bytes.Buffer
	relsBuf.ReadFrom(rc)
	rc.Close()

	relsContent := relsBuf.String()
	if !strings.Contains(relsContent, "rId1") {
		t.Errorf(".rels should contain rId1, got:\n%s", relsContent)
	}
	if !strings.Contains(relsContent, "media/image1.png") {
		t.Errorf(".rels should reference media/image1.png, got:\n%s", relsContent)
	}
	if !strings.Contains(relsContent, RTImage) {
		t.Errorf(".rels should contain image relationship type, got:\n%s", relsContent)
	}
}
