package opc

import (
	"testing"
)

func TestParseContentTypes(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`

	ct, err := ParseContentTypes([]byte(xml))
	if err != nil {
		t.Fatalf("ParseContentTypes: %v", err)
	}

	// Test override
	got, err := ct.ContentType("/word/document.xml")
	if err != nil {
		t.Fatalf("ContentType: %v", err)
	}
	if got != CTWmlDocumentMain {
		t.Errorf("got %q, want %q", got, CTWmlDocumentMain)
	}

	// Test default
	got, err = ct.ContentType("/word/someother.xml")
	if err != nil {
		t.Fatalf("ContentType for xml: %v", err)
	}
	if got != CTXml {
		t.Errorf("got %q, want %q", got, CTXml)
	}

	// Test missing
	_, err = ct.ContentType("/nonexistent.xyz")
	if err == nil {
		t.Error("expected error for unknown partname/extension")
	}
}

func TestParseRelationships(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com" TargetMode="External"/>
</Relationships>`

	srels, err := ParseRelationships([]byte(xml), "/")
	if err != nil {
		t.Fatalf("ParseRelationships: %v", err)
	}
	if len(srels) != 2 {
		t.Fatalf("expected 2 rels, got %d", len(srels))
	}

	// Internal rel
	if srels[0].RID != "rId1" {
		t.Errorf("expected rId1, got %q", srels[0].RID)
	}
	if srels[0].RelType != RTOfficeDocument {
		t.Errorf("wrong reltype: %q", srels[0].RelType)
	}
	if srels[0].TargetMode != TargetModeInternal {
		t.Errorf("expected Internal, got %q", srels[0].TargetMode)
	}
	if srels[0].IsExternal() {
		t.Error("expected internal relationship")
	}
	pn := srels[0].TargetPartname()
	if pn != "/word/document.xml" {
		t.Errorf("TargetPartname = %q, want /word/document.xml", pn)
	}

	// External rel
	if srels[1].RID != "rId2" {
		t.Errorf("expected rId2, got %q", srels[1].RID)
	}
	if !srels[1].IsExternal() {
		t.Error("expected external relationship")
	}
	if srels[1].TargetRef != "http://example.com" {
		t.Errorf("expected http://example.com, got %q", srels[1].TargetRef)
	}
}

func TestParseRelationships_Nil(t *testing.T) {
	srels, err := ParseRelationships(nil, "/")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(srels) != 0 {
		t.Errorf("expected empty, got %d", len(srels))
	}
}

func TestSerializeContentTypes_RoundTrip(t *testing.T) {
	parts := []PartInfo{
		{PartName: "/word/document.xml", ContentType: CTWmlDocumentMain},
		{PartName: "/word/styles.xml", ContentType: CTWmlStyles},
		{PartName: "/word/media/image1.png", ContentType: CTPng},
	}

	blob, err := SerializeContentTypes(parts)
	if err != nil {
		t.Fatalf("SerializeContentTypes: %v", err)
	}

	ct, err := ParseContentTypes(blob)
	if err != nil {
		t.Fatalf("ParseContentTypes: %v", err)
	}

	// image1.png should be resolved via default extension
	got, err := ct.ContentType("/word/media/image1.png")
	if err != nil {
		t.Fatalf("ContentType png: %v", err)
	}
	if got != CTPng {
		t.Errorf("got %q, want %q", got, CTPng)
	}

	// document.xml should be override
	got, err = ct.ContentType("/word/document.xml")
	if err != nil {
		t.Fatalf("ContentType docxml: %v", err)
	}
	if got != CTWmlDocumentMain {
		t.Errorf("got %q, want %q", got, CTWmlDocumentMain)
	}
}

func TestSerializeRelationships_RoundTrip(t *testing.T) {
	rels := NewRelationships("/word")

	// Use Load to add with known rId
	part := NewBasePart("/word/styles.xml", CTWmlStyles, nil, nil)
	rels.Load("rId1", RTStyles, "styles.xml", part, false)
	rels.Load("rId2", RTHyperlink, "http://example.com", nil, true)

	blob, err := SerializeRelationships(rels)
	if err != nil {
		t.Fatalf("SerializeRelationships: %v", err)
	}

	srels, err := ParseRelationships(blob, "/word")
	if err != nil {
		t.Fatalf("ParseRelationships: %v", err)
	}

	if len(srels) != 2 {
		t.Fatalf("expected 2 rels, got %d", len(srels))
	}

	// Check first rel
	if srels[0].RID != "rId1" {
		t.Errorf("expected rId1, got %q", srels[0].RID)
	}
	if srels[0].RelType != RTStyles {
		t.Errorf("wrong reltype: %q", srels[0].RelType)
	}

	// Check second rel (external)
	if srels[1].RID != "rId2" {
		t.Errorf("expected rId2, got %q", srels[1].RID)
	}
	if !srels[1].IsExternal() {
		t.Error("expected external")
	}
}

func TestContentTypeMap_CaseInsensitive(t *testing.T) {
	ct := NewContentTypeMap()
	ct.AddDefault("XML", CTXml)

	got, err := ct.ContentType("/test.xml")
	if err != nil {
		t.Fatalf("ContentType: %v", err)
	}
	if got != CTXml {
		t.Errorf("got %q, want %q", got, CTXml)
	}
}

// TestContentType_InfersPngFromExtension verifies that a ContentTypeMap
// without an explicit Default for "png" still resolves it via the well-known
// extension fallback.  This fixes the "no content type for partname" errors
// on bnc780044_spacing.docx, bnc891663.docx, etc.
func TestContentType_InfersPngFromExtension(t *testing.T) {
	t.Parallel()

	// Only xml and rels defaults — no png.
	ct := NewContentTypeMap()
	ct.AddDefault("xml", CTXml)
	ct.AddDefault("rels", CTOpcRelationships)

	got, err := ct.ContentType("/word/media/image1.png")
	if err != nil {
		t.Fatalf("expected png to be inferred, got error: %v", err)
	}
	if got != CTPng {
		t.Errorf("got %q, want %q", got, CTPng)
	}
}

// TestContentType_InfersJpegFromExtension verifies jpg/jpeg inference.
func TestContentType_InfersJpegFromExtension(t *testing.T) {
	t.Parallel()

	ct := NewContentTypeMap()

	for _, ext := range []string{"jpg", "jpeg", "jpe"} {
		uri := PackURI("/word/media/photo." + ext)
		got, err := ct.ContentType(uri)
		if err != nil {
			t.Errorf("extension %q: expected inference, got error: %v", ext, err)
			continue
		}
		if got != CTJpeg {
			t.Errorf("extension %q: got %q, want %q", ext, got, CTJpeg)
		}
	}
}

// TestContentType_InfersCommonFormats verifies inference for all image and
// embedded formats that appeared in real-world failing documents.
func TestContentType_InfersCommonFormats(t *testing.T) {
	t.Parallel()

	ct := NewContentTypeMap() // completely empty — no defaults at all

	cases := []struct {
		ext  string
		want string
	}{
		{"png", CTPng},
		{"jpg", CTJpeg},
		{"jpeg", CTJpeg},
		{"gif", CTGif},
		{"bmp", CTBmp},
		{"tiff", CTTiff},
		{"tif", CTTiff},
		{"emf", CTXEmf},
		{"wmf", CTXWmf},
		{"xlsx", CTSmlSheet},
	}

	for _, tc := range cases {
		uri := PackURI("/word/media/file." + tc.ext)
		got, err := ct.ContentType(uri)
		if err != nil {
			t.Errorf("extension %q: expected %q, got error: %v", tc.ext, tc.want, err)
			continue
		}
		if got != tc.want {
			t.Errorf("extension %q: got %q, want %q", tc.ext, got, tc.want)
		}
	}
}

// TestContentType_ExplicitDefaultTakesPrecedence verifies that an explicit
// Default in [Content_Types].xml wins over the well-known fallback.
func TestContentType_ExplicitDefaultTakesPrecedence(t *testing.T) {
	t.Parallel()

	ct := NewContentTypeMap()
	ct.AddDefault("png", "image/custom-png") // non-standard override

	got, err := ct.ContentType("/word/media/image1.png")
	if err != nil {
		t.Fatalf("ContentType: %v", err)
	}
	if got != "image/custom-png" {
		t.Errorf("explicit default should win, got %q", got)
	}
}

// TestContentType_OverrideTakesPrecedenceOverInference verifies that an
// explicit Override wins over inference.
func TestContentType_OverrideTakesPrecedenceOverInference(t *testing.T) {
	t.Parallel()

	ct := NewContentTypeMap()
	ct.AddOverride("/word/media/image1.png", "image/custom-override")

	got, err := ct.ContentType("/word/media/image1.png")
	if err != nil {
		t.Fatalf("ContentType: %v", err)
	}
	if got != "image/custom-override" {
		t.Errorf("override should win, got %q", got)
	}
}

// TestContentType_UnknownExtensionStillFails verifies that truly unknown
// extensions still return an error.
func TestContentType_UnknownExtensionStillFails(t *testing.T) {
	t.Parallel()

	ct := NewContentTypeMap()

	_, err := ct.ContentType("/word/data/something.xyz123")
	if err == nil {
		t.Error("expected error for completely unknown extension")
	}
}

// TestContentType_InferenceCaseInsensitive verifies that inference works
// regardless of extension case (PNG, Png, pNg, etc).
func TestContentType_InferenceCaseInsensitive(t *testing.T) {
	t.Parallel()

	ct := NewContentTypeMap()

	for _, ext := range []string{"PNG", "Png", "pNg"} {
		uri := PackURI("/word/media/image." + ext)
		got, err := ct.ContentType(uri)
		if err != nil {
			t.Errorf("extension %q: expected inference, got error: %v", ext, err)
			continue
		}
		if got != CTPng {
			t.Errorf("extension %q: got %q, want %q", ext, got, CTPng)
		}
	}
}
