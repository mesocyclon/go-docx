package opc

import (
	"bytes"
	"errors"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/templates"
)

func loadDefaultDocx(t *testing.T) []byte {
	t.Helper()
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}
	return data
}

func TestPhysPkgReader_Open(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	uris := reader.URIs()
	if len(uris) == 0 {
		t.Fatal("expected non-empty URIs list")
	}

	// Check that /word/document.xml is among the URIs
	found := false
	for _, uri := range uris {
		if uri == "/word/document.xml" {
			found = true
			break
		}
	}
	if !found {
		t.Error("expected /word/document.xml in URIs list")
	}
}

func TestPhysPkgReader_BlobFor(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.BlobFor("/word/document.xml")
	if err != nil {
		t.Fatalf("BlobFor: %v", err)
	}
	if len(blob) == 0 {
		t.Error("expected non-empty blob for /word/document.xml")
	}
}

func TestPhysPkgReader_ContentTypesXml(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.ContentTypesXml()
	if err != nil {
		t.Fatalf("ContentTypesXml: %v", err)
	}
	if len(blob) == 0 {
		t.Error("expected non-empty [Content_Types].xml")
	}
	if !bytes.Contains(blob, []byte("ContentType")) {
		t.Error("expected [Content_Types].xml to contain ContentType")
	}
}

func TestPhysPkgReader_RelsXmlFor(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.RelsXmlFor(PackageURI)
	if err != nil {
		t.Fatalf("RelsXmlFor: %v", err)
	}
	if blob == nil {
		t.Error("expected package-level .rels to exist")
	}
	if !bytes.Contains(blob, []byte("Relationship")) {
		t.Error("expected .rels to contain Relationship elements")
	}
}

func TestPhysPkgWriter_RoundTrip(t *testing.T) {
	// Write a simple package and read it back
	var buf bytes.Buffer
	writer := NewPhysPkgWriter(&buf)
	err := writer.Write("/test/data.xml", []byte("<root/>"))
	if err != nil {
		t.Fatalf("Write: %v", err)
	}
	err = writer.Close()
	if err != nil {
		t.Fatalf("Close: %v", err)
	}

	// Read back
	reader, err := NewPhysPkgReaderFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.BlobFor("/test/data.xml")
	if err != nil {
		t.Fatalf("BlobFor: %v", err)
	}
	if string(blob) != "<root/>" {
		t.Errorf("got %q, want %q", string(blob), "<root/>")
	}
}

// TestNewPhysPkgReader_OLE2_ReturnsEncryptedError verifies that opening an
// OLE2 Compound Document (encrypted .docx) returns ErrEncryptedPackage with
// a clear message, not a confusing "not a valid zip file" error.
func TestNewPhysPkgReader_OLE2_ReturnsEncryptedError(t *testing.T) {
	t.Parallel()

	// Minimal OLE2 header: 8-byte magic + padding to a plausible size.
	// Real encrypted .docx files start with these exact bytes.
	ole2Header := make([]byte, 512)
	copy(ole2Header, []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})

	_, err := NewPhysPkgReaderFromBytes(ole2Header)
	if err == nil {
		t.Fatal("expected error for OLE2 input, got nil")
	}

	if !errors.Is(err, ErrEncryptedPackage) {
		t.Errorf("expected ErrEncryptedPackage, got: %v", err)
	}

	// Must NOT match ErrNotZipPackage — it's more specific.
	if errors.Is(err, ErrNotZipPackage) {
		t.Error("OLE2 error should not also match ErrNotZipPackage")
	}

	// Error message should mention encryption.
	msg := err.Error()
	if !containsSubstring(msg, "encrypted") && !containsSubstring(msg, "OLE2") {
		t.Errorf("error message should mention encryption, got: %s", msg)
	}
}

// TestNewPhysPkgReader_RandomBytes_ReturnsNotZipError verifies that garbage
// input returns ErrNotZipPackage with a clear message.
func TestNewPhysPkgReader_RandomBytes_ReturnsNotZipError(t *testing.T) {
	t.Parallel()

	garbage := []byte("this is not a zip file at all, just random text content")

	_, err := NewPhysPkgReaderFromBytes(garbage)
	if err == nil {
		t.Fatal("expected error for garbage input, got nil")
	}

	if !errors.Is(err, ErrNotZipPackage) {
		t.Errorf("expected ErrNotZipPackage, got: %v", err)
	}

	// Must NOT match ErrEncryptedPackage.
	if errors.Is(err, ErrEncryptedPackage) {
		t.Error("garbage input should not match ErrEncryptedPackage")
	}
}

// TestNewPhysPkgReader_EmptyInput_ReturnsNotZipError verifies that empty
// input returns ErrNotZipPackage.
func TestNewPhysPkgReader_EmptyInput_ReturnsNotZipError(t *testing.T) {
	t.Parallel()

	_, err := NewPhysPkgReaderFromBytes([]byte{})
	if err == nil {
		t.Fatal("expected error for empty input, got nil")
	}

	if !errors.Is(err, ErrNotZipPackage) {
		t.Errorf("expected ErrNotZipPackage, got: %v", err)
	}
}

// TestNewPhysPkgReader_ErrorsUnwrap verifies that the original zip error
// is preserved in the error chain for debugging.
func TestNewPhysPkgReader_ErrorsUnwrap(t *testing.T) {
	t.Parallel()

	garbage := []byte("not a zip")
	_, err := NewPhysPkgReaderFromBytes(garbage)
	if err == nil {
		t.Fatal("expected error")
	}

	// The original zip error should be reachable via Unwrap.
	msg := err.Error()
	if !containsSubstring(msg, "zip") {
		t.Errorf("error chain should include original zip error, got: %s", msg)
	}
}

func containsSubstring(s, substr string) bool {
	return len(s) >= len(substr) && searchSubstring(s, substr)
}

func searchSubstring(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}
