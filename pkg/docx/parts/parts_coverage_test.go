package parts

import (
	"bytes"
	goimage "image"
	"image/png"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/image"
	"github.com/vortex/go-docx/pkg/docx/opc"
)

// =========================================================================
// Test helpers
// =========================================================================

// newDocPartWithBody creates a minimal DocumentPart with <w:body/> and
// attached OPC package. Optionally pre-wires extra relationship parts.
func newDocPartWithBody(t *testing.T) (*DocumentPart, *opc.OpcPackage) {
	t.Helper()
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)
	pkg := opc.NewOpcPackage(NewDocxPartFactory())
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)
	return dp, pkg
}

// wireNumberingPart creates and wires a NumberingPart to dp.
func wireNumberingPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *NumberingPart {
	t.Helper()
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(np)
	dp.Rels().GetOrAdd(opc.RTNumbering, np)
	return np
}

// wireFootnotesPart creates and wires a FootnotesPart to dp.
func wireFootnotesPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *FootnotesPart {
	t.Helper()
	fp, err := DefaultFootnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(fp)
	dp.Rels().GetOrAdd(opc.RTFootnotes, fp)
	return fp
}

// wireEndnotesPart creates and wires an EndnotesPart to dp.
func wireEndnotesPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *EndnotesPart {
	t.Helper()
	ep, err := DefaultEndnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(ep)
	dp.Rels().GetOrAdd(opc.RTEndnotes, ep)
	return ep
}

// wireCommentsPart creates and wires a CommentsPart to dp.
func wireCommentsPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *CommentsPart {
	t.Helper()
	cp, err := DefaultCommentsPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(cp)
	dp.Rels().GetOrAdd(opc.RTComments, cp)
	return cp
}

// stylesXML returns styles XML with a single paragraph style.
func stylesXML(styleID, styleName string, isDefault bool) []byte {
	def := ""
	if isDefault {
		def = ` w:default="1"`
	}
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="` + styleID + `"` + def + `>
    <w:name w:val="` + styleName + `"/>
  </w:style>
</w:styles>`)
}

// wireStylesPartFromXML wires a custom styles part with given XML.
func wireStylesPartFromXML(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage, xml []byte) *StylesPart {
	t.Helper()
	xp, err := opc.NewXmlPart("/word/styles.xml", opc.CTWmlStyles, xml, pkg)
	if err != nil {
		t.Fatal(err)
	}
	sp := NewStylesPart(xp)
	pkg.AddPart(sp)
	dp.Rels().GetOrAdd(opc.RTStyles, sp)
	return sp
}

// minimumPNG returns a valid 1x1 PNG blob.
func minimumPNG() []byte {
	img := goimage.NewRGBA(goimage.Rect(0, 0, 1, 1))
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		panic(err)
	}
	return buf.Bytes()
}

// =========================================================================
// BookmarkID allocation — document.go:69-154
// =========================================================================

func TestNextBookmarkID_EmptyDocument(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	id1 := dp.NextBookmarkID()
	if id1 != 1 {
		t.Errorf("first bookmark ID = %d, want 1", id1)
	}
	id2 := dp.NextBookmarkID()
	if id2 != 2 {
		t.Errorf("second bookmark ID = %d, want 2", id2)
	}
}

func TestNextBookmarkID_ExistingBookmarks(t *testing.T) {
	// Create document with existing bookmark IDs.
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="5" w:name="b1"/>
      <w:bookmarkEnd w:id="5"/>
      <w:bookmarkStart w:id="10" w:name="b2"/>
      <w:bookmarkEnd w:id="10"/>
    </w:p>
  </w:body>
</w:document>`)
	pkg := opc.NewOpcPackage(NewDocxPartFactory())
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	id := dp.NextBookmarkID()
	if id != 11 {
		t.Errorf("NextBookmarkID after existing max=10: got %d, want 11", id)
	}
}

func TestNextBookmarkID_ScansHeadersAndFooters(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)

	// Add a header with bookmark id=20
	hp, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}
	_ = rId
	// Inject a bookmarkStart into the header element
	bms := etree.NewElement("w:bookmarkStart")
	bms.CreateAttr("w:id", "20")
	bms.CreateAttr("w:name", "hdr_bm")
	hp.Element().AddChild(bms)

	// Add a footer with bookmark id=30
	fp, _, err := dp.AddFooterPart()
	if err != nil {
		t.Fatal(err)
	}
	bms2 := etree.NewElement("w:bookmarkStart")
	bms2.CreateAttr("w:id", "30")
	bms2.CreateAttr("w:name", "ftr_bm")
	fp.Element().AddChild(bms2)
	_ = pkg

	id := dp.NextBookmarkID()
	if id != 31 {
		t.Errorf("NextBookmarkID should scan headers/footers: got %d, want 31", id)
	}
}

func TestNextBookmarkID_ScansComments(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)

	// Wire comments part with embedded bookmark
	commentsBlob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test">
    <w:p>
      <w:bookmarkStart w:id="50" w:name="c_bm"/>
      <w:bookmarkEnd w:id="50"/>
    </w:p>
  </w:comment>
</w:comments>`)
	xp, err := opc.NewXmlPart("/word/comments.xml", opc.CTWmlComments, commentsBlob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	cp := NewCommentsPart(xp)
	pkg.AddPart(cp)
	dp.Rels().GetOrAdd(opc.RTComments, cp)

	id := dp.NextBookmarkID()
	if id != 51 {
		t.Errorf("NextBookmarkID should scan comments: got %d, want 51", id)
	}
}

func TestNextBookmarkNameSuffix_Monotonic(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	s1 := dp.NextBookmarkNameSuffix()
	s2 := dp.NextBookmarkNameSuffix()
	s3 := dp.NextBookmarkNameSuffix()
	if s1 != 1 || s2 != 2 || s3 != 3 {
		t.Errorf("suffixes = %d,%d,%d, want 1,2,3", s1, s2, s3)
	}
}

func TestCollectMaxBookmarkID_NoBookmarks(t *testing.T) {
	el := etree.NewElement("w:body")
	child := etree.NewElement("w:p")
	el.AddChild(child)
	if got := collectMaxBookmarkID(el); got != 0 {
		t.Errorf("collectMaxBookmarkID with no bookmarks = %d, want 0", got)
	}
}

func TestCollectMaxBookmarkID_NestedBookmarks(t *testing.T) {
	root := etree.NewElement("w:body")
	p := etree.NewElement("w:p")
	root.AddChild(p)
	bms := etree.NewElement("w:bookmarkStart")
	bms.CreateAttr("w:id", "7")
	p.AddChild(bms)
	bme := etree.NewElement("w:bookmarkEnd")
	bme.CreateAttr("w:id", "7")
	p.AddChild(bme)

	// Nested deeper
	tc := etree.NewElement("w:tc")
	p2 := etree.NewElement("w:p")
	bms2 := etree.NewElement("w:bookmarkStart")
	bms2.CreateAttr("w:id", "15")
	p2.AddChild(bms2)
	tc.AddChild(p2)
	root.AddChild(tc)

	if got := collectMaxBookmarkID(root); got != 15 {
		t.Errorf("collectMaxBookmarkID nested = %d, want 15", got)
	}
}

func TestCollectMaxBookmarkID_InvalidID(t *testing.T) {
	root := etree.NewElement("w:body")
	bms := etree.NewElement("w:bookmarkStart")
	bms.CreateAttr("w:id", "abc") // non-numeric
	root.AddChild(bms)

	if got := collectMaxBookmarkID(root); got != 0 {
		t.Errorf("collectMaxBookmarkID with invalid id = %d, want 0", got)
	}
}

// =========================================================================
// NumberingPart resolution — document.go:297-335
// =========================================================================

func TestNumberingPart_Existing(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireNumberingPart(t, dp, pkg)

	got, err := dp.NumberingPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("NumberingPart should return the wired part")
	}
}

func TestNumberingPart_Cached(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireNumberingPart(t, dp, pkg)

	np1, _ := dp.NumberingPart()
	np2, _ := dp.NumberingPart()
	if np1 != np2 {
		t.Error("NumberingPart should be cached after first access")
	}
}

func TestNumberingPart_NotFound(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	_, err := dp.NumberingPart()
	if err == nil {
		t.Error("expected error when no numbering part exists")
	}
}

func TestGetOrAddNumberingPart_CreatesWhenMissing(t *testing.T) {
	dp, _ := newDocPartWithBody(t)

	np, err := dp.GetOrAddNumberingPart()
	if err != nil {
		t.Fatal(err)
	}
	if np == nil {
		t.Fatal("GetOrAddNumberingPart returned nil")
	}

	// Should return element
	numEl, err := np.Numbering()
	if err != nil {
		t.Fatal(err)
	}
	if numEl == nil {
		t.Error("Numbering() returned nil on default numbering part")
	}
}

func TestGetOrAddNumberingPart_ReturnsExisting(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireNumberingPart(t, dp, pkg)

	got, err := dp.GetOrAddNumberingPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("GetOrAddNumberingPart should return existing part")
	}
}

// =========================================================================
// FootnotesPart resolution — document.go:343-376
// =========================================================================

func TestFootnotesPart_Existing(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireFootnotesPart(t, dp, pkg)

	got, err := dp.FootnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("FootnotesPart should return the wired part")
	}
}

func TestFootnotesPart_NotFound(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	_, err := dp.FootnotesPart()
	if err == nil {
		t.Error("expected error when no footnotes part exists")
	}
}

func TestGetOrAddFootnotesPart_CreatesWhenMissing(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	fp, err := dp.GetOrAddFootnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if fp == nil {
		t.Fatal("GetOrAddFootnotesPart returned nil")
	}
	if fp.Element() == nil {
		t.Error("default footnotes part element is nil")
	}
}

func TestGetOrAddFootnotesPart_ReturnsExisting(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireFootnotesPart(t, dp, pkg)

	got, err := dp.GetOrAddFootnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("GetOrAddFootnotesPart should return existing part")
	}
}

// =========================================================================
// EndnotesPart resolution — document.go:384-417
// =========================================================================

func TestEndnotesPart_Existing(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireEndnotesPart(t, dp, pkg)

	got, err := dp.EndnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("EndnotesPart should return the wired part")
	}
}

func TestEndnotesPart_NotFound(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	_, err := dp.EndnotesPart()
	if err == nil {
		t.Error("expected error when no endnotes part exists")
	}
}

func TestGetOrAddEndnotesPart_CreatesWhenMissing(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	ep, err := dp.GetOrAddEndnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if ep == nil {
		t.Fatal("GetOrAddEndnotesPart returned nil")
	}
	if ep.Element() == nil {
		t.Error("default endnotes part element is nil")
	}
}

func TestGetOrAddEndnotesPart_ReturnsExisting(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireEndnotesPart(t, dp, pkg)

	got, err := dp.GetOrAddEndnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("GetOrAddEndnotesPart should return existing part")
	}
}

// =========================================================================
// HasCommentsPart / CommentsElement — document.go:464-503
// =========================================================================

func TestHasCommentsPart_False(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	if dp.HasCommentsPart() {
		t.Error("HasCommentsPart should return false when no comments wired")
	}
}

func TestHasCommentsPart_True(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireCommentsPart(t, dp, pkg)
	if !dp.HasCommentsPart() {
		t.Error("HasCommentsPart should return true when comments wired")
	}
}

func TestCommentsElement_ReturnsElement(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireCommentsPart(t, dp, pkg)

	ce, err := dp.CommentsElement()
	if err != nil {
		t.Fatal(err)
	}
	if ce == nil {
		t.Error("CommentsElement should not be nil")
	}
}

func TestCommentsElement_CreatesDefaultWhenMissing(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	ce, err := dp.CommentsElement()
	if err != nil {
		t.Fatal(err)
	}
	if ce == nil {
		t.Error("CommentsElement should auto-create default")
	}
}

// =========================================================================
// CoreProperties — document.go:513-537
// =========================================================================

func TestCoreProperties_CreatesDefault(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	cp, err := dp.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	if cp == nil {
		t.Fatal("CoreProperties returned nil")
	}
	ct, err := cp.CT()
	if err != nil {
		t.Fatal(err)
	}
	if ct == nil {
		t.Error("CT() returned nil on default core properties")
	}
}

func TestCoreProperties_ReturnsExisting(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	cp1, err := dp.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	cp2, err := dp.CoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	if cp1 != cp2 {
		t.Error("CoreProperties should return same instance via package rel cache")
	}
}

// =========================================================================
// CorePropertiesPart.CT() — coreprops.go:28-34
// =========================================================================

func TestCorePropertiesPart_CT_ValidElement(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	ct, err := cp.CT()
	if err != nil {
		t.Fatal(err)
	}
	if ct == nil {
		t.Fatal("CT() should return non-nil")
	}
}

func TestCorePropertiesPart_CT_PartName(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if cp.PartName() != "/docProps/core.xml" {
		t.Errorf("partname = %q, want /docProps/core.xml", cp.PartName())
	}
}

// =========================================================================
// DefaultCorePropertiesPart — coreprops.go:44-61
// =========================================================================

func TestDefaultCorePropertiesPart_HasMetadata(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if cp.Element() == nil {
		t.Fatal("element is nil")
	}
	// Verify partname
	if cp.PartName() != "/docProps/core.xml" {
		t.Errorf("partname = %q, want /docProps/core.xml", cp.PartName())
	}
}

// =========================================================================
// NumberingPart.Numbering() — numbering.go:25-31
// =========================================================================

func TestNumberingPart_Numbering_Valid(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	num, err := np.Numbering()
	if err != nil {
		t.Fatal(err)
	}
	if num == nil {
		t.Error("Numbering() should not return nil on valid part")
	}
}

func TestNumberingPart_Numbering_PartName(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if np.PartName() != "/word/numbering.xml" {
		t.Errorf("partname = %q, want /word/numbering.xml", np.PartName())
	}
}

// =========================================================================
// DefaultNumberingPart — numbering.go:38-49
// =========================================================================

func TestDefaultNumberingPart_Creates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if np.Element() == nil {
		t.Error("element is nil")
	}
	if np.PartName() != "/word/numbering.xml" {
		t.Errorf("partname = %q, want /word/numbering.xml", np.PartName())
	}
}

// =========================================================================
// DefaultFootnotesPart / DefaultEndnotesPart — footnotes.go:47-107
// =========================================================================

func TestDefaultFootnotesPart_Creates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	fp, err := DefaultFootnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if fp.Element() == nil {
		t.Fatal("default footnotes part element is nil")
	}
	if fp.PartName() != "/word/footnotes.xml" {
		t.Errorf("partname = %q, want /word/footnotes.xml", fp.PartName())
	}
}

func TestDefaultEndnotesPart_Creates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	ep, err := DefaultEndnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if ep.Element() == nil {
		t.Fatal("default endnotes part element is nil")
	}
	if ep.PartName() != "/word/endnotes.xml" {
		t.Errorf("partname = %q, want /word/endnotes.xml", ep.PartName())
	}
}

// =========================================================================
// GetStyle — document.go:555-573
// =========================================================================

func TestGetStyle_NilID_ReturnsDefault(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	s, err := dp.GetStyle(nil, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if s == nil {
		t.Fatal("GetStyle with nil ID should return default style")
	}
	if s.StyleId() != "Normal" {
		t.Errorf("default style id = %q, want Normal", s.StyleId())
	}
}

func TestGetStyle_MatchingID(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Heading1", "heading 1", false))

	id := "Heading1"
	s, err := dp.GetStyle(&id, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if s == nil {
		t.Fatal("GetStyle should find style by ID")
	}
	if s.StyleId() != "Heading1" {
		t.Errorf("style id = %q, want Heading1", s.StyleId())
	}
}

func TestGetStyle_NonexistentID_FallsBackToDefault(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	id := "NonExistent"
	s, err := dp.GetStyle(&id, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	// Should fall back to default
	if s != nil && s.StyleId() != "Normal" {
		t.Errorf("fallback style id = %q, want Normal", s.StyleId())
	}
}

func TestGetStyle_WrongType_FallsBackToDefault(t *testing.T) {
	// Style exists as paragraph type but we request character type.
	xml := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="DefaultParagraphFont" w:default="1">
    <w:name w:val="Default Paragraph Font"/>
  </w:style>
</w:styles>`)
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, xml)

	id := "Normal" // paragraph type
	s, err := dp.GetStyle(&id, enum.WdStyleTypeCharacter)
	if err != nil {
		t.Fatal(err)
	}
	// Should return character default, not the paragraph Normal
	if s != nil && s.StyleId() == "Normal" {
		t.Error("GetStyle should not return paragraph style for character type request")
	}
}

// =========================================================================
// GetStyleID — document.go:592-629
// =========================================================================

// mockStyledObject implements styledObject interface for testing.
type mockStyledObject struct {
	styleID   string
	styleType enum.WdStyleType
	typeErr   error
}

func (m *mockStyledObject) StyleID() string                  { return m.styleID }
func (m *mockStyledObject) Type() (enum.WdStyleType, error)  { return m.styleType, m.typeErr }

func TestGetStyleID_NilInput(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	id, err := dp.GetStyleID(nil, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id != nil {
		t.Errorf("GetStyleID(nil) should return nil, got %v", *id)
	}
}

func TestGetStyleID_StringName(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	xml := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
  </w:style>
</w:styles>`)
	wireStylesPartFromXML(t, dp, pkg, xml)

	id, err := dp.GetStyleID("heading 1", enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id == nil {
		t.Fatal("GetStyleID should return non-nil for non-default style")
	}
	if *id != "Heading1" {
		t.Errorf("GetStyleID = %q, want Heading1", *id)
	}
}

func TestGetStyleID_StringName_DefaultReturnsNil(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	// When style resolves to the default, GetStyleID returns nil per Python behavior.
	id, err := dp.GetStyleID("Normal", enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id != nil {
		t.Errorf("GetStyleID for default style should return nil, got %q", *id)
	}
}

func TestGetStyleID_StyledObject(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	xml := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="MyStyle">
    <w:name w:val="My Style"/>
  </w:style>
</w:styles>`)
	wireStylesPartFromXML(t, dp, pkg, xml)

	obj := &mockStyledObject{styleID: "MyStyle", styleType: enum.WdStyleTypeParagraph}
	id, err := dp.GetStyleID(obj, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id == nil || *id != "MyStyle" {
		t.Errorf("GetStyleID with styledObject = %v, want MyStyle", id)
	}
}

func TestGetStyleID_StyledObject_DefaultReturnsNil(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	obj := &mockStyledObject{styleID: "Normal", styleType: enum.WdStyleTypeParagraph}
	id, err := dp.GetStyleID(obj, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id != nil {
		t.Error("GetStyleID for default style object should return nil")
	}
}

func TestGetStyleID_StyledObject_TypeMismatch(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	obj := &mockStyledObject{styleID: "Normal", styleType: enum.WdStyleTypeCharacter}
	_, err := dp.GetStyleID(obj, enum.WdStyleTypeParagraph)
	if err == nil {
		t.Error("GetStyleID should error on type mismatch")
	}
}

func TestGetStyleID_UnsupportedType(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	_, err := dp.GetStyleID(42, enum.WdStyleTypeParagraph)
	if err == nil {
		t.Error("GetStyleID should error on unsupported type (int)")
	}
}

// =========================================================================
// InlineShapeElements — document.go:640-658
// =========================================================================

func TestInlineShapeElements_NoInlines(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	inlines, err := dp.InlineShapeElements()
	if err != nil {
		t.Fatal(err)
	}
	if len(inlines) != 0 {
		t.Errorf("expected 0 inlines, got %d", len(inlines))
	}
}

func TestInlineShapeElements_FindsInlines(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="1000" cy="2000"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="3000" cy="4000"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`)
	pkg := opc.NewOpcPackage(NewDocxPartFactory())
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	inlines, err := dp.InlineShapeElements()
	if err != nil {
		t.Fatal(err)
	}
	if len(inlines) != 2 {
		t.Errorf("expected 2 inlines, got %d", len(inlines))
	}
}

// =========================================================================
// StoryPart — story.go:32-97, 121-159, 168-213
// =========================================================================

func TestNewStoryPart(t *testing.T) {
	el := etree.NewElement("w:body")
	xp := opc.NewXmlPartFromElement("/word/document.xml", opc.CTWmlDocumentMain, el, nil)
	sp := NewStoryPart(xp)
	if sp == nil {
		t.Fatal("NewStoryPart returned nil")
	}
	if sp.Element() != el {
		t.Error("NewStoryPart should wrap the element")
	}
}

func TestStoryPart_DocumentPart_NoPackage(t *testing.T) {
	el := etree.NewElement("w:hdr")
	xp := opc.NewXmlPartFromElement("/word/header1.xml", opc.CTWmlHeader, el, nil)
	sp := NewStoryPart(xp)

	// documentPart() should error if no package is set
	_, err := sp.GetStyle(nil, enum.WdStyleTypeParagraph)
	if err == nil {
		t.Error("GetStyle should error when no package/document part available")
	}
}

func TestStoryPart_GetStyle_DelegatesToDocPart(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	// Create a header part wired to the same package
	hp, _, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}
	// hp.StoryPart should delegate GetStyle to dp
	s, err := hp.GetStyle(nil, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if s == nil {
		t.Error("header StoryPart.GetStyle should resolve through document part")
	}
}

func TestStoryPart_GetStyleID_DelegatesToDocPart(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wireStylesPartFromXML(t, dp, pkg, stylesXML("Normal", "Normal", true))

	hp, _, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}
	id, err := hp.GetStyleID(nil, enum.WdStyleTypeParagraph)
	if err != nil {
		t.Fatal(err)
	}
	if id != nil {
		t.Errorf("GetStyleID(nil) should return nil, got %v", *id)
	}
}

func TestStoryPart_WmlPackage_Nil(t *testing.T) {
	el := etree.NewElement("w:body")
	xp := opc.NewXmlPartFromElement("/word/document.xml", opc.CTWmlDocumentMain, el, nil)
	sp := NewStoryPart(xp)
	wp := sp.wmlPackage()
	if wp != nil {
		t.Error("wmlPackage should return nil when no package set")
	}
}

func TestStoryPart_WmlPackage_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	wp := NewWmlPackage(pkg)
	pkg.SetAppPackage(wp)

	dp := getDocumentPart(t, pkg)
	got := dp.wmlPackage()
	if got != wp {
		t.Error("wmlPackage should return the WmlPackage from OpcPackage.AppPackage")
	}
}

// =========================================================================
// StoryPart.GetOrAddImage / NewPicInline — story.go:45-68
// =========================================================================

func TestStoryPart_GetOrAddImage(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	blob := []byte("fake image data")
	ip := NewImagePartWithMeta("/word/media/image1.png", opc.CTPng, blob, 100, 100, 96, 96, "test.png")
	pkg.AddPart(ip)

	rId, result := dp.GetOrAddImage(ip)
	if rId == "" {
		t.Error("GetOrAddImage returned empty rId")
	}
	if result != ip {
		t.Error("GetOrAddImage should return same image part")
	}

	// Call again — should return same rId (idempotent via GetOrAdd)
	rId2, _ := dp.GetOrAddImage(ip)
	if rId2 != rId {
		t.Errorf("GetOrAddImage not idempotent: rId1=%q, rId2=%q", rId, rId2)
	}
}

func TestStoryPart_NewPicInline(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	ip := NewImagePartWithMeta("/word/media/image1.png", opc.CTPng, nil, 100, 200, 96, 96, "test.png")
	pkg.AddPart(ip)

	inline, err := dp.NewPicInline(ip, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if inline == nil {
		t.Fatal("NewPicInline returned nil")
	}
}

// =========================================================================
// StoryPart.GetOrAddImageFromReader / NewPicInlineFromReader — story.go:168-213
// =========================================================================

func TestStoryPart_GetOrAddImageFromReader_NoWmlPackage(t *testing.T) {
	el := etree.NewElement("w:body")
	xp := opc.NewXmlPartFromElement("/word/document.xml", opc.CTWmlDocumentMain, el, nil)
	sp := NewStoryPart(xp)

	r := bytes.NewReader(minimumPNG())
	_, _, err := sp.GetOrAddImageFromReader(r)
	if err == nil {
		t.Error("GetOrAddImageFromReader should error without WmlPackage")
	}
}

func TestStoryPart_GetOrAddImageFromReader(t *testing.T) {
	pkg := openDefaultDocx(t)
	wp := NewWmlPackage(pkg)
	pkg.SetAppPackage(wp)
	dp := getDocumentPart(t, pkg)

	pngBlob := minimumPNG()
	r := bytes.NewReader(pngBlob)
	rId, ip, err := dp.GetOrAddImageFromReader(r)
	if err != nil {
		t.Fatal(err)
	}
	if rId == "" {
		t.Error("rId is empty")
	}
	if ip == nil {
		t.Fatal("image part is nil")
	}

	// Verify dedup: adding same image again returns same part
	r2 := bytes.NewReader(pngBlob)
	rId2, ip2, err := dp.GetOrAddImageFromReader(r2)
	if err != nil {
		t.Fatal(err)
	}
	if ip2 != ip {
		t.Error("same image blob should be deduped")
	}
	if rId2 != rId {
		t.Errorf("same image should get same rId: %q vs %q", rId, rId2)
	}
}

func TestStoryPart_NewPicInlineFromReader(t *testing.T) {
	pkg := openDefaultDocx(t)
	wp := NewWmlPackage(pkg)
	pkg.SetAppPackage(wp)
	dp := getDocumentPart(t, pkg)

	r := bytes.NewReader(minimumPNG())
	inline, err := dp.NewPicInlineFromReader(r, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if inline == nil {
		t.Fatal("NewPicInlineFromReader returned nil")
	}
}

// =========================================================================
// ImagePart — SetImageMeta, PxWidth/PxHeight/HorzDpi/VertDpi, ensureMeta
// =========================================================================

func TestImagePart_SetImageMeta(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, nil, nil)
	ip.SetImageMeta(800, 600, 150, 150)

	w, err := ip.PxWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w != 800 {
		t.Errorf("PxWidth = %d, want 800", w)
	}

	h, err := ip.PxHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h != 600 {
		t.Errorf("PxHeight = %d, want 600", h)
	}

	hd, err := ip.HorzDpi()
	if err != nil {
		t.Fatal(err)
	}
	if hd != 150 {
		t.Errorf("HorzDpi = %d, want 150", hd)
	}

	vd, err := ip.VertDpi()
	if err != nil {
		t.Fatal(err)
	}
	if vd != 150 {
		t.Errorf("VertDpi = %d, want 150", vd)
	}
}

func TestImagePart_LazyMetaFromBlob(t *testing.T) {
	// Create ImagePart with a real PNG blob (no metadata set upfront).
	// ensureMeta should parse the blob automatically.
	pngBlob := minimumPNG()
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, pngBlob, nil)

	w, err := ip.PxWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w != 1 {
		t.Errorf("PxWidth = %d, want 1 (1x1 PNG)", w)
	}

	h, err := ip.PxHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h != 1 {
		t.Errorf("PxHeight = %d, want 1", h)
	}
}

func TestImagePart_EnsureMeta_EmptyBlob(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, []byte{}, nil)
	_, err := ip.PxWidth()
	if err == nil {
		t.Error("PxWidth should error on empty blob")
	}
}

func TestImagePart_NewImagePartFromImage(t *testing.T) {
	pngBlob := minimumPNG()
	// Verify blob is valid
	_, _, err := goimage.Decode(bytes.NewReader(pngBlob))
	if err != nil {
		t.Fatal(err)
	}

	imgLib, err := image.FromBlob(pngBlob, "photo.png")
	if err != nil {
		t.Fatal(err)
	}

	ip := NewImagePartFromImage(imgLib, pngBlob)
	if ip == nil {
		t.Fatal("NewImagePartFromImage returned nil")
	}
	fn := ip.Filename()
	if fn == "" {
		t.Error("Filename should not be empty")
	}

	// Hash should be carried from Image
	h, err := ip.Hash()
	if err != nil {
		t.Fatal(err)
	}
	if h == "" {
		t.Error("Hash should not be empty")
	}
}

// =========================================================================
// WmlPackage.AfterUnmarshal — wmlpackage.go:65-85
// =========================================================================

func TestWmlPackage_AfterUnmarshal_GathersImages(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	wp := NewWmlPackage(pkg)

	// Create a document part with an image relationship
	el := etree.NewElement("w:document")
	xp := opc.NewXmlPartFromElement("/word/document.xml", opc.CTWmlDocumentMain, el, pkg)
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	ip := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("data"), pkg)
	pkg.AddPart(ip)
	dp.Rels().GetOrAdd(opc.RTImage, ip)

	if wp.ImageParts().Len() != 0 {
		t.Fatal("should start with 0 image parts")
	}

	wp.AfterUnmarshal()

	if wp.ImageParts().Len() != 1 {
		t.Errorf("AfterUnmarshal should gather image parts: got %d, want 1", wp.ImageParts().Len())
	}
}

func TestWmlPackage_AfterUnmarshal_SkipsDuplicates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	wp := NewWmlPackage(pkg)

	el := etree.NewElement("w:document")
	xp := opc.NewXmlPartFromElement("/word/document.xml", opc.CTWmlDocumentMain, el, pkg)
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	ip := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("data"), pkg)
	pkg.AddPart(ip)
	// Wire same image via two different relationships
	dp.Rels().GetOrAdd(opc.RTImage, ip)

	wp.AfterUnmarshal()
	wp.AfterUnmarshal() // second call

	if wp.ImageParts().Len() != 1 {
		t.Errorf("AfterUnmarshal should skip duplicates: got %d, want 1", wp.ImageParts().Len())
	}
}

func TestWmlPackage_AfterUnmarshal_SkipsExternal(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	wp := NewWmlPackage(pkg)

	// No image rels — just an external rel (won't have TargetPart)
	el := etree.NewElement("w:document")
	xp := opc.NewXmlPartFromElement("/word/document.xml", opc.CTWmlDocumentMain, el, pkg)
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	wp.AfterUnmarshal()

	if wp.ImageParts().Len() != 0 {
		t.Errorf("AfterUnmarshal should ignore non-image rels: got %d, want 0", wp.ImageParts().Len())
	}
}

// =========================================================================
// ImageParts.All — wmlpackage.go:125-127
// =========================================================================

func TestImageParts_All(t *testing.T) {
	ips := NewImageParts()
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("a"), nil)
	ip2 := NewImagePart("/word/media/image2.png", opc.CTPng, []byte("b"), nil)
	ips.Append(ip1)
	ips.Append(ip2)

	all := ips.All()
	if len(all) != 2 {
		t.Errorf("All() len = %d, want 2", len(all))
	}
	if all[0] != ip1 || all[1] != ip2 {
		t.Error("All() should return parts in insertion order")
	}
}

func TestImageParts_All_Empty(t *testing.T) {
	ips := NewImageParts()
	all := ips.All()
	if len(all) != 0 {
		t.Errorf("All() on empty = %d, want 0", len(all))
	}
}
