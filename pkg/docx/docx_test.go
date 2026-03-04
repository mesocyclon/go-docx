package docx

import (
	"bytes"
	"testing"
)

func TestNew(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error: %v", err)
	}
	if doc == nil {
		t.Fatal("New() returned nil")
	}
	// The built-in default.docx body contains only w:sectPr, no w:p.
	// Matches Python: len(Document().paragraphs) == 0.
	paras, err2 := doc.Paragraphs()
	if err2 != nil {
		t.Fatalf("Paragraphs() error: %v", err2)
	}
	if len(paras) != 0 {
		t.Errorf("expected 0 paragraphs in new doc, got %d", len(paras))
	}
}

func TestNew_Sections(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error: %v", err)
	}
	sections := doc.Sections()
	if sections.Len() < 1 {
		t.Errorf("expected at least 1 section, got %d", sections.Len())
	}
}

func TestOpenBytes_InvalidData(t *testing.T) {
	_, err := OpenBytes([]byte("not a zip file"))
	if err == nil {
		t.Error("OpenBytes(invalid) expected error, got nil")
	}
}

func TestOpenBytes_RoundTrip(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error: %v", err)
	}
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save() error: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes() error: %v", err)
	}
	if doc2 == nil {
		t.Fatal("OpenBytes() returned nil")
	}
}
