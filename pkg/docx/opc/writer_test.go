package opc

import (
	"bytes"
	"fmt"
	"testing"
)

func TestPackageWriter_WriteEmptyPackage(t *testing.T) {
	t.Parallel()

	var buf bytes.Buffer
	pw := &PackageWriter{}
	rels := NewRelationships("/")

	err := pw.Write(&buf, rels, nil)
	if err != nil {
		t.Fatalf("Write empty package: %v", err)
	}

	// The output should be a valid ZIP containing [Content_Types].xml and /_rels/.rels
	reader, err := NewPhysPkgReaderFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("reading back empty package: %v", err)
	}
	defer reader.Close()

	ctBlob, err := reader.ContentTypesXml()
	if err != nil {
		t.Fatalf("ContentTypesXml: %v", err)
	}
	if len(ctBlob) == 0 {
		t.Error("expected non-empty [Content_Types].xml")
	}

	relsBlob, err := reader.RelsXmlFor(PackageURI)
	if err != nil {
		t.Fatalf("RelsXmlFor: %v", err)
	}
	if relsBlob == nil {
		t.Error("expected package-level .rels to exist")
	}
}

func TestPackageWriter_WriteParts(t *testing.T) {
	t.Parallel()

	var buf bytes.Buffer
	pw := &PackageWriter{}
	rels := NewRelationships("/")

	part1 := NewBasePart("/word/document.xml", CTWmlDocumentMain, []byte("<w:document/>"), nil)
	part2 := NewBasePart("/word/styles.xml", CTWmlStyles, []byte("<w:styles/>"), nil)

	// Add a relationship from part1 to part2
	part1.Rels().Add(RTStyles, "styles.xml", part2, false)

	// Package-level rel to part1
	rels.Add(RTOfficeDocument, "word/document.xml", part1, false)

	err := pw.Write(&buf, rels, []Part{part1, part2})
	if err != nil {
		t.Fatalf("Write: %v", err)
	}

	// Read back and verify
	reader, err := NewPhysPkgReaderFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("reading back: %v", err)
	}
	defer reader.Close()

	blob, err := reader.BlobFor("/word/document.xml")
	if err != nil {
		t.Fatalf("BlobFor document: %v", err)
	}
	if string(blob) != "<w:document/>" {
		t.Errorf("document blob: got %q, want %q", string(blob), "<w:document/>")
	}

	blob2, err := reader.BlobFor("/word/styles.xml")
	if err != nil {
		t.Fatalf("BlobFor styles: %v", err)
	}
	if string(blob2) != "<w:styles/>" {
		t.Errorf("styles blob: got %q, want %q", string(blob2), "<w:styles/>")
	}

	// Part1 should have a .rels file
	partRels, err := reader.RelsXmlFor("/word/document.xml")
	if err != nil {
		t.Fatalf("RelsXmlFor document: %v", err)
	}
	if partRels == nil {
		t.Error("expected document part to have .rels")
	}
}

func TestPackageWriter_PartWithNoRels(t *testing.T) {
	t.Parallel()

	var buf bytes.Buffer
	pw := &PackageWriter{}
	rels := NewRelationships("/")

	// Part with empty rels (Len == 0)
	part := NewBasePart("/word/settings.xml", CTWmlSettings, []byte("<w:settings/>"), nil)
	rels.Add(RTSettings, "word/settings.xml", part, false)

	err := pw.Write(&buf, rels, []Part{part})
	if err != nil {
		t.Fatalf("Write: %v", err)
	}

	reader, err := NewPhysPkgReaderFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("reading back: %v", err)
	}
	defer reader.Close()

	// Part's .rels should not exist (empty relationships)
	partRels, err := reader.RelsXmlFor("/word/settings.xml")
	if err != nil {
		t.Fatalf("RelsXmlFor: %v", err)
	}
	if partRels != nil {
		t.Error("expected no .rels for part with empty relationships")
	}
}

// errorPart is a test Part whose Blob() always returns an error.
type errorPart struct {
	BasePart
}

func (p *errorPart) Blob() ([]byte, error) {
	return nil, fmt.Errorf("simulated blob error")
}

func TestPackageWriter_BlobError(t *testing.T) {
	t.Parallel()

	var buf bytes.Buffer
	pw := &PackageWriter{}
	rels := NewRelationships("/")
	part := &errorPart{BasePart: *NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, nil)}

	err := pw.Write(&buf, rels, []Part{part})
	if err == nil {
		t.Fatal("expected error from Blob(), got nil")
	}
}
