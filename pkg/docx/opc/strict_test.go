package opc

import (
	"testing"
)

func TestNormalizeRelType_StrictToTransitional(t *testing.T) {
	t.Parallel()

	cases := []struct {
		input string
		want  string
	}{
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
			RTOfficeDocument,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/styles",
			RTStyles,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/header",
			RTHeader,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/footer",
			RTFooter,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/image",
			RTImage,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable",
			RTFontTable,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/theme",
			RTTheme,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/settings",
			RTSettings,
		},
		{
			"http://purl.oclc.org/ooxml/officeDocument/relationships/numbering",
			RTNumbering,
		},
	}

	for _, tc := range cases {
		got := NormalizeRelType(tc.input)
		if got != tc.want {
			t.Errorf("NormalizeRelType(%q)\n  got  %q\n  want %q", tc.input, got, tc.want)
		}
	}
}

func TestNormalizeRelType_TransitionalUnchanged(t *testing.T) {
	t.Parallel()

	transitional := []string{
		RTOfficeDocument,
		RTStyles,
		RTHeader,
		RTImage,
		RTCoreProperties,
		RTThumbnail,
		RTHyperlink,
	}

	for _, rt := range transitional {
		got := NormalizeRelType(rt)
		if got != rt {
			t.Errorf("NormalizeRelType(%q) should pass through, got %q", rt, got)
		}
	}
}

func TestNormalizeRelType_UnrelatedUnchanged(t *testing.T) {
	t.Parallel()

	// Microsoft-extension relationship types should not be modified.
	ms := "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
	got := NormalizeRelType(ms)
	if got != ms {
		t.Errorf("NormalizeRelType should not touch MS URIs, got %q", got)
	}
}

func TestParseRelationships_NormalizesStrictRelType(t *testing.T) {
	t.Parallel()

	// .rels XML with a strict relationship type URI.
	relsXml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles"
    Target="word/styles.xml"/>
</Relationships>`

	srels, err := ParseRelationships([]byte(relsXml), "/")
	if err != nil {
		t.Fatalf("ParseRelationships: %v", err)
	}
	if len(srels) != 2 {
		t.Fatalf("expected 2 rels, got %d", len(srels))
	}

	if srels[0].RelType != RTOfficeDocument {
		t.Errorf("rId1: expected %q, got %q", RTOfficeDocument, srels[0].RelType)
	}
	if srels[1].RelType != RTStyles {
		t.Errorf("rId2: expected %q, got %q", RTStyles, srels[1].RelType)
	}
}

func TestOpenBytes_StrictPackageRels(t *testing.T) {
	t.Parallel()

	// Package .rels with strict relationship type for officeDocument.
	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml": minimalContentTypes,
		"_rels/.rels":         pkgRels,
		"word/document.xml":   minimalDocumentXml,
	})

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes with strict rels should succeed, got: %v", err)
	}

	docPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart should find document via normalized rel type, got: %v", err)
	}
	if docPart.PartName() != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %q", docPart.PartName())
	}
}

func TestOpenBytes_StrictPartLevelRels(t *testing.T) {
	t.Parallel()

	pkgRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`

	// document.xml.rels uses strict URIs for styles and settings.
	docRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/settings"
    Target="settings.xml"/>
</Relationships>`

	stylesContentTypes := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`

	stylesXml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`

	settingsXml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`

	data := buildTestZip(t, map[string]string{
		"[Content_Types].xml":          stylesContentTypes,
		"_rels/.rels":                  pkgRels,
		"word/document.xml":            minimalDocumentXml,
		"word/_rels/document.xml.rels": docRels,
		"word/styles.xml":              stylesXml,
		"word/settings.xml":            settingsXml,
	})

	pkg, err := OpenBytes(data, nil)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	docPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}

	// Part-level rels should be normalized — lookup by transitional RTStyles should work.
	stylesRel := docPart.Rels().GetByRID("rId1")
	if stylesRel == nil {
		t.Fatal("rId1 (styles) not found on document part")
	}
	if stylesRel.RelType != RTStyles {
		t.Errorf("expected normalized %q, got %q", RTStyles, stylesRel.RelType)
	}

	settingsRel := docPart.Rels().GetByRID("rId2")
	if settingsRel == nil {
		t.Fatal("rId2 (settings) not found on document part")
	}
	if settingsRel.RelType != RTSettings {
		t.Errorf("expected normalized %q, got %q", RTSettings, settingsRel.RelType)
	}
}
