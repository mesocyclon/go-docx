package opc

import (
	"testing"
)

// ---------------------------------------------------------------------------
// BasePart
// ---------------------------------------------------------------------------

func TestBasePart_Accessors(t *testing.T) {
	t.Parallel()

	pkg := NewOpcPackage(nil)
	blob := []byte("binary data")
	part := NewBasePart("/word/document.xml", CTWmlDocumentMain, blob, pkg)

	if part.PartName() != "/word/document.xml" {
		t.Errorf("PartName: got %q", part.PartName())
	}
	if part.ContentType() != CTWmlDocumentMain {
		t.Errorf("ContentType: got %q", part.ContentType())
	}
	gotBlob, err := part.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}
	if string(gotBlob) != "binary data" {
		t.Errorf("Blob: got %q", string(gotBlob))
	}
	if part.Rels() == nil {
		t.Error("Rels should not be nil")
	}
	if part.Package() != pkg {
		t.Error("Package mismatch")
	}

	// SetPartName
	part.SetPartName("/word/newname.xml")
	if part.PartName() != "/word/newname.xml" {
		t.Errorf("after SetPartName: got %q", part.PartName())
	}

	// SetBlob
	part.SetBlob([]byte("new data"))
	gotBlob, _ = part.Blob()
	if string(gotBlob) != "new data" {
		t.Errorf("after SetBlob: got %q", string(gotBlob))
	}

	// SetRels
	newRels := NewRelationships("/word")
	part.SetRels(newRels)
	if part.Rels() != newRels {
		t.Error("SetRels did not update")
	}

	// BeforeMarshal and AfterUnmarshal should be no-ops (no panic)
	part.BeforeMarshal()
	part.AfterUnmarshal()
}

// ---------------------------------------------------------------------------
// XmlPart
// ---------------------------------------------------------------------------

func TestXmlPart_FromValidXml(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0" encoding="UTF-8"?><root><child/></root>`)
	part, err := NewXmlPart("/word/document.xml", CTWmlDocumentMain, xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}
	el := part.Element()
	if el == nil {
		t.Fatal("Element should not be nil")
	}
	if el.Tag != "root" {
		t.Errorf("expected root tag, got %q", el.Tag)
	}
}

func TestXmlPart_FromInvalidXml(t *testing.T) {
	t.Parallel()

	garbage := []byte("this is not XML at all <<<>>>")
	_, err := NewXmlPart("/word/document.xml", CTWmlDocumentMain, garbage, nil)
	if err == nil {
		t.Fatal("expected error for invalid XML, got nil")
	}
}

func TestXmlPart_Blob_RoundTrip(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0" encoding="UTF-8"?><root><child attr="val"></child></root>`)
	part, err := NewXmlPart("/word/document.xml", CTWmlDocumentMain, xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}

	blob, err := part.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}
	if len(blob) == 0 {
		t.Fatal("expected non-empty blob")
	}
	// Should contain XML declaration
	if !containsSubstring(string(blob), "<?xml") {
		t.Error("blob should contain <?xml declaration")
	}
	// Should contain our content
	if !containsSubstring(string(blob), "root") {
		t.Error("blob should contain root element")
	}
}

func TestXmlPart_Blob_NilDoc(t *testing.T) {
	t.Parallel()

	part := &XmlPart{
		BasePart: *NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, nil),
		doc:      nil,
	}

	blob, err := part.Blob()
	if err != nil {
		t.Fatalf("Blob with nil doc: %v", err)
	}
	if blob != nil {
		t.Errorf("expected nil blob for nil doc, got %d bytes", len(blob))
	}
}

func TestXmlPart_SetElement(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0"?><old/>`)
	part, err := NewXmlPart("/test.xml", "application/xml", xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}
	if part.Element().Tag != "old" {
		t.Fatalf("expected 'old' tag, got %q", part.Element().Tag)
	}

	newXml := []byte(`<?xml version="1.0"?><new/>`)
	part2, _ := NewXmlPart("/test2.xml", "application/xml", newXml, nil)
	part.SetElement(part2.Element())

	if part.Element().Tag != "new" {
		t.Errorf("after SetElement: expected 'new' tag, got %q", part.Element().Tag)
	}
}

func TestXmlPartFromElement(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0"?><root/>`)
	original, err := NewXmlPart("/temp.xml", "application/xml", xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}

	part := NewXmlPartFromElement("/word/document.xml", CTWmlDocumentMain, original.Element(), nil)
	if part.PartName() != "/word/document.xml" {
		t.Errorf("PartName: got %q", part.PartName())
	}
	if part.Element().Tag != "root" {
		t.Errorf("Element tag: got %q", part.Element().Tag)
	}
}

// ---------------------------------------------------------------------------
// PartFactory
// ---------------------------------------------------------------------------

func TestPartFactory_ContentTypeMap(t *testing.T) {
	t.Parallel()

	factory := NewPartFactory()
	factory.Register(CTWmlDocumentMain, func(pn PackURI, ct, rt string, blob []byte, pkg *OpcPackage) (Part, error) {
		return NewXmlPart(pn, ct, blob, pkg)
	})

	xml := []byte(`<?xml version="1.0"?><w:document/>`)
	part, err := factory.New("/word/document.xml", CTWmlDocumentMain, RTOfficeDocument, xml, nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	if _, ok := part.(*XmlPart); !ok {
		t.Errorf("expected *XmlPart, got %T", part)
	}
}

func TestPartFactory_Selector(t *testing.T) {
	t.Parallel()

	factory := NewPartFactory()
	// Register a content-type constructor (should NOT be used)
	factory.Register(CTWmlDocumentMain, func(pn PackURI, ct, rt string, blob []byte, pkg *OpcPackage) (Part, error) {
		return NewBasePart(pn, ct, blob, pkg), nil
	})
	// Register a selector that overrides
	factory.SetSelector(func(ct, rt string) PartConstructor {
		if rt == RTOfficeDocument {
			return func(pn PackURI, ct, rt string, blob []byte, pkg *OpcPackage) (Part, error) {
				return NewXmlPart(pn, ct, blob, pkg)
			}
		}
		return nil
	})

	xml := []byte(`<?xml version="1.0"?><w:document/>`)
	part, err := factory.New("/word/document.xml", CTWmlDocumentMain, RTOfficeDocument, xml, nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	// Selector should have produced XmlPart, not BasePart
	if _, ok := part.(*XmlPart); !ok {
		t.Errorf("expected selector to produce *XmlPart, got %T", part)
	}
}

func TestPartFactory_DefaultFallback(t *testing.T) {
	t.Parallel()

	factory := NewPartFactory()

	blob := []byte("binary data")
	part, err := factory.New("/word/media/image1.png", "image/png", RTImage, blob, nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	if _, ok := part.(*BasePart); !ok {
		t.Errorf("expected *BasePart fallback, got %T", part)
	}
	gotBlob, _ := part.Blob()
	if string(gotBlob) != "binary data" {
		t.Errorf("blob mismatch: got %q", string(gotBlob))
	}
}

// ---------------------------------------------------------------------------
// escapeAttrWhitespace
// ---------------------------------------------------------------------------

func TestEscapeAttrWhitespace_NewlinesInAttr(t *testing.T) {
	t.Parallel()
	input := []byte(`<v:textpath string="Line1&#10;Line2&#10;"/>`)
	// After etree parse→serialize, &#10; becomes literal \n:
	broken := []byte("<v:textpath string=\"Line1\nLine2\n\"/>")
	got := escapeAttrWhitespace(broken)
	want := `<v:textpath string="Line1&#10;Line2&#10;"/>`
	if string(got) != want {
		t.Errorf("escapeAttrWhitespace:\n got: %q\nwant: %q", string(got), want)
	}
	// Original with &#10; already encoded should pass through (no literal \n).
	got2 := escapeAttrWhitespace(input)
	if string(got2) != string(input) {
		t.Errorf("should not modify already-escaped: %q", string(got2))
	}
}

func TestEscapeAttrWhitespace_TabsAndCR(t *testing.T) {
	t.Parallel()
	input := []byte("<el attr=\"a\tb\rc\"/>")
	got := escapeAttrWhitespace(input)
	want := `<el attr="a&#9;b&#13;c"/>`
	if string(got) != want {
		t.Errorf("got: %q\nwant: %q", string(got), want)
	}
}

func TestEscapeAttrWhitespace_TextContentUntouched(t *testing.T) {
	t.Parallel()
	// Newlines in text content (outside tags) must NOT be escaped.
	input := []byte("<root>\n  <child>text\nhere</child>\n</root>")
	got := escapeAttrWhitespace(input)
	if string(got) != string(input) {
		t.Errorf("text content was modified:\n got: %q\norig: %q", string(got), string(input))
	}
}

func TestEscapeAttrWhitespace_SingleQuoteAttr(t *testing.T) {
	t.Parallel()
	input := []byte("<el attr='a\nb'/>")
	got := escapeAttrWhitespace(input)
	want := "<el attr='a&#10;b'/>"
	if string(got) != want {
		t.Errorf("got: %q\nwant: %q", string(got), want)
	}
}

func TestEscapeAttrWhitespace_NoSpecialChars(t *testing.T) {
	t.Parallel()
	input := []byte(`<root><child attr="value">text</child></root>`)
	got := escapeAttrWhitespace(input)
	// Should return same slice (no allocation).
	if &got[0] != &input[0] {
		t.Error("expected same slice when no escaping needed")
	}
}

func TestEscapeAttrWhitespace_VMLRealistic(t *testing.T) {
	t.Parallel()
	// Simulates what etree produces for fdo74110.docx v:textpath
	input := []byte(`<v:textpath style="font-family:&quot;Noto Sans&quot;;font-size:28pt" string="IBM RoadRunner` + "\n" + `Blade Center` + "\n" + `QS22/LS21 Cluster` + "\n" + `"></v:textpath>`)
	got := escapeAttrWhitespace(input)
	want := `<v:textpath style="font-family:&quot;Noto Sans&quot;;font-size:28pt" string="IBM RoadRunner&#10;Blade Center&#10;QS22/LS21 Cluster&#10;"></v:textpath>`
	if string(got) != want {
		t.Errorf("VML roundtrip:\n got: %s\nwant: %s", string(got), want)
	}
}

func TestXmlPart_Blob_EscapesAttrNewlines(t *testing.T) {
	t.Parallel()
	// XML with &#10; in attribute — should survive parse→serialize roundtrip.
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<root><el attr="line1&#10;line2"></el></root>`
	xp, err := NewXmlPart("/test.xml", CTXml, []byte(xml), nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}
	blob, err := xp.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}
	s := string(blob)
	if !contains(s, "line1&#10;line2") {
		t.Errorf("&#10; not preserved in attribute:\n%s", s)
	}
}

func contains(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}