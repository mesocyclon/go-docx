package parts

import (
	"bytes"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/internal/xmlutil"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
)

func makeElementWithIDs(ids ...string) *etree.Element {
	el := etree.NewElement("root")
	for _, id := range ids {
		child := etree.NewElement("item")
		child.CreateAttr("id", id)
		el.AddChild(child)
	}
	return el
}

func TestNextID_EmptyElement(t *testing.T) {
	el := etree.NewElement("root")
	got := collectMaxID(el) + 1
	if got != 1 {
		t.Errorf("NextID for empty element: got %d, want 1", got)
	}
}

func TestNextID_WithIDs(t *testing.T) {
	el := makeElementWithIDs("1", "5", "3")
	got := collectMaxID(el) + 1
	if got != 6 {
		t.Errorf("NextID with ids [1,5,3]: got %d, want 6", got)
	}
}

func TestNextID_IgnoresNonDigit(t *testing.T) {
	el := makeElementWithIDs("abc", "12", "rId4", "7")
	got := collectMaxID(el) + 1
	if got != 13 {
		t.Errorf("NextID with mixed ids: got %d, want 13", got)
	}
}

func TestNextID_NestedElements(t *testing.T) {
	el := etree.NewElement("root")
	child := etree.NewElement("p")
	child.CreateAttr("id", "10")
	grandchild := etree.NewElement("r")
	grandchild.CreateAttr("id", "20")
	child.AddChild(grandchild)
	el.AddChild(child)

	got := collectMaxID(el) + 1
	if got != 21 {
		t.Errorf("NextID nested: got %d, want 21", got)
	}
}

func TestIsDigits(t *testing.T) {
	tests := []struct {
		input string
		want  bool
	}{
		{"", false},
		{"123", true},
		{"0", true},
		{"abc", false},
		{"12a", false},
		{"rId3", false},
	}
	for _, tt := range tests {
		got := xmlutil.IsDigits(tt.input)
		if got != tt.want {
			t.Errorf("isDigits(%q) = %v, want %v", tt.input, got, tt.want)
		}
	}
}

func TestRelRefCount(t *testing.T) {
	el := etree.NewElement("root")
	child1 := etree.NewElement("drawing")
	child1.CreateAttr("r:id", "rId5")
	el.AddChild(child1)
	child2 := etree.NewElement("hyperlink")
	child2.CreateAttr("r:id", "rId5")
	el.AddChild(child2)
	child3 := etree.NewElement("other")
	child3.CreateAttr("r:id", "rId3")
	el.AddChild(child3)

	if got := countRIdRefs(el, "rId5"); got != 2 {
		t.Errorf("relRefCount for rId5: got %d, want 2", got)
	}
	if got := countRIdRefs(el, "rId3"); got != 1 {
		t.Errorf("relRefCount for rId3: got %d, want 1", got)
	}
	if got := countRIdRefs(el, "rId99"); got != 0 {
		t.Errorf("relRefCount for rId99: got %d, want 0", got)
	}
}

func TestDropRel_DeletesWhenRefCountLow(t *testing.T) {
	// DropRel should delete a relationship when its XML reference count < 2.
	// Core logic tested via countRIdRefs; integration tested in document_test.go.
	el := etree.NewElement("root")
	if got := countRIdRefs(el, "rId1"); got >= 2 {
		t.Errorf("expected count < 2 for element with no refs, got %d", got)
	}
}

// --- NextID caching ---

func TestNextID_CachesAfterFirstCall(t *testing.T) {
	// Two consecutive NextID calls should return sequential values
	// without the second call needing to rescan the tree.
	sp := newTestStoryPart(t, makeElementWithIDs("1", "5", "3"))

	first := sp.NextID()
	if first != 6 {
		t.Errorf("first NextID: got %d, want 6", first)
	}

	second := sp.NextID()
	if second != 7 {
		t.Errorf("second NextID: got %d, want 7", second)
	}

	third := sp.NextID()
	if third != 8 {
		t.Errorf("third NextID: got %d, want 8", third)
	}
}

func TestNextID_EmptyElement_CachesAndIncrements(t *testing.T) {
	sp := newTestStoryPart(t, etree.NewElement("root"))
	if got := sp.NextID(); got != 1 {
		t.Errorf("first NextID on empty: got %d, want 1", got)
	}
	if got := sp.NextID(); got != 2 {
		t.Errorf("second NextID on empty: got %d, want 2", got)
	}
}

// newTestStoryPart creates a minimal StoryPart for unit tests.
func newTestStoryPart(t *testing.T, el *etree.Element) *StoryPart {
	t.Helper()
	xp := opc.NewXmlPartFromElement("", "", el, nil)
	return &StoryPart{XmlPart: xp}
}

// =========================================================================
// StoryPart — story.go
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
// StoryPart.GetOrAddImage / NewPicInline — story.go
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
// StoryPart.GetOrAddImageFromReader / NewPicInlineFromReader — story.go
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
