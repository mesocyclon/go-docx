package parts

import (
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// openDefaultDocx opens the embedded default.docx using the WML part factory.
func openDefaultDocx(t *testing.T) *opc.OpcPackage {
	t.Helper()
	docxBytes, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}
	factory := NewDocxPartFactory()
	pkg, err := opc.OpenBytes(docxBytes, factory)
	if err != nil {
		t.Fatalf("opening default.docx: %v", err)
	}
	return pkg
}

func getDocumentPart(t *testing.T, pkg *opc.OpcPackage) *DocumentPart {
	t.Helper()
	mainPart, err := pkg.MainDocumentPart()
	if err != nil {
		t.Fatalf("MainDocumentPart: %v", err)
	}
	dp, ok := mainPart.(*DocumentPart)
	if !ok {
		t.Fatalf("MainDocumentPart is %T, want *DocumentPart", mainPart)
	}
	return dp
}

func TestOpenDefaultDocx_DocumentPartType(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	if dp == nil {
		t.Fatal("DocumentPart is nil")
	}
}

func TestDocumentPart_Body_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	body, err := dp.Body()
	if err != nil {
		t.Fatal(err)
	}
	if body == nil {
		t.Fatal("Body is nil")
	}
}

func TestDocumentPart_StylesPart_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	sp, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp == nil {
		t.Fatal("StylesPart is nil")
	}
}

func TestDocumentPart_StylesPart_Cached(t *testing.T) {
	// In Python, _styles_part is @property (not lazyproperty), but the
	// relationship graph acts as the cache — same object returned each time.
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	sp1, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	sp2, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp1 != sp2 {
		t.Error("StylesPart should return cached instance")
	}
}

func TestDocumentPart_StylesPart_CreatesDefault(t *testing.T) {
	// Create a minimal document part with no styles relationship
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	factory := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(factory)
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	sp, err := dp.StylesPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp == nil {
		t.Fatal("StylesPart should create default when absent")
	}
	// Verify it's discoverable via relationship graph (acts as cache)
	sp2, _ := dp.StylesPart()
	if sp != sp2 {
		t.Error("default StylesPart should be found via relationship graph")
	}
}

func TestDocumentPart_SettingsPart_CreatesDefault(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	sp, err := dp.SettingsPart()
	if err != nil {
		t.Fatal(err)
	}
	if sp == nil {
		t.Fatal("SettingsPart should create default when absent")
	}
}

func TestDocumentPart_CommentsPart_CreatesDefault(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)

	cp, err := dp.CommentsPart()
	if err != nil {
		t.Fatal(err)
	}
	if cp == nil {
		t.Fatal("CommentsPart should create default when absent")
	}
}

func TestDocumentPart_AddHeaderPart(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	hp, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}
	if rId == "" {
		t.Error("rId is empty")
	}
	if hp == nil {
		t.Fatal("HeaderPart is nil")
	}
	if hp.Element() == nil {
		t.Error("HeaderPart element is nil")
	}
}

func TestDocumentPart_AddFooterPart(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	fp, rId, err := dp.AddFooterPart()
	if err != nil {
		t.Fatal(err)
	}
	if rId == "" {
		t.Error("rId is empty")
	}
	if fp == nil {
		t.Fatal("FooterPart is nil")
	}
	if fp.Element() == nil {
		t.Error("FooterPart element is nil")
	}
}

func TestDocumentPart_DropHeaderPart(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	_, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}

	// Verify relationship exists
	rel := dp.Rels().GetByRID(rId)
	if rel == nil {
		t.Fatal("relationship should exist before drop")
	}

	dp.DropHeaderPart(rId)

	rel = dp.Rels().GetByRID(rId)
	if rel != nil {
		t.Error("relationship should be deleted after drop (no XML refs)")
	}
}

func TestDocumentPart_HeaderPartByRID(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	hp, rId, err := dp.AddHeaderPart()
	if err != nil {
		t.Fatal(err)
	}

	got, err := dp.HeaderPartByRID(rId)
	if err != nil {
		t.Fatal(err)
	}
	if got != hp {
		t.Error("HeaderPartByRID should return the same part")
	}
}

func TestDocumentPart_FooterPartByRID(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	fp, rId, err := dp.AddFooterPart()
	if err != nil {
		t.Fatal(err)
	}

	got, err := dp.FooterPartByRID(rId)
	if err != nil {
		t.Fatal(err)
	}
	if got != fp {
		t.Error("FooterPartByRID should return the same part")
	}
}

func TestDocumentPart_HeaderPartByRID_NotFound(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)

	_, err := dp.HeaderPartByRID("rId999")
	if err == nil {
		t.Error("expected error for non-existent rId")
	}
}

func TestDocumentPart_Styles_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	styles, err := dp.Styles()
	if err != nil {
		t.Fatal(err)
	}
	if styles == nil {
		t.Fatal("Styles should not be nil")
	}
}

func TestDocumentPart_Settings_NotNil(t *testing.T) {
	pkg := openDefaultDocx(t)
	dp := getDocumentPart(t, pkg)
	settings, err := dp.Settings()
	if err != nil {
		t.Fatal(err)
	}
	if settings == nil {
		t.Fatal("Settings should not be nil")
	}
}

// =========================================================================
// BookmarkID allocation — document.go
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
// NumberingPart resolution — document.go
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
// HasCommentsPart / CommentsElement — document.go
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
// GetStyle — document.go
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
// GetStyleID — document.go
// =========================================================================

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
// InlineShapeElements — document.go
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
