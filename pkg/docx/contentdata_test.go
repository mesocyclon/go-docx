package docx

import (
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// ---------------------------------------------------------------------------
// Phase 1: isRelAttr tests
// ---------------------------------------------------------------------------

func TestIsRelAttr_Prefix_RId(t *testing.T) {
	attr := etree.Attr{Space: "r", Key: "id", Value: "rId1"}
	if !isRelAttr(attr) {
		t.Error("expected r:id to be a rel attr")
	}
}

func TestIsRelAttr_Prefix_REmbed(t *testing.T) {
	attr := etree.Attr{Space: "r", Key: "embed", Value: "rId2"}
	if !isRelAttr(attr) {
		t.Error("expected r:embed to be a rel attr")
	}
}

func TestIsRelAttr_Prefix_RLink(t *testing.T) {
	attr := etree.Attr{Space: "r", Key: "link", Value: "rId3"}
	if !isRelAttr(attr) {
		t.Error("expected r:link to be a rel attr")
	}
}

func TestIsRelAttr_NonRelAttr(t *testing.T) {
	cases := []etree.Attr{
		{Space: "w", Key: "val", Value: "Normal"},
		{Space: "w", Key: "id", Value: "42"},
		{Space: "r", Key: "val", Value: "rId1"},
		{Space: "", Key: "id", Value: "rId1"},
		{Space: "r", Key: "type", Value: "rId1"},
	}
	for _, attr := range cases {
		if isRelAttr(attr) {
			t.Errorf("expected %s:%s to NOT be a rel attr", attr.Space, attr.Key)
		}
	}
}

func TestIsRelAttr_FullNS(t *testing.T) {
	// Full namespace URI form (rare but valid).
	attr := etree.Attr{
		Space: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Key:   "embed",
		Value: "rId7",
	}
	if !isRelAttr(attr) {
		t.Error("expected full-NS r:embed to be a rel attr")
	}
}

func TestIsRelAttr_FullNS_Id(t *testing.T) {
	attr := etree.Attr{
		Space: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Key:   "id",
		Value: "rId1",
	}
	if !isRelAttr(attr) {
		t.Error("expected full-NS r:id to be a rel attr")
	}
}

func TestIsRelAttr_FullNS_Link(t *testing.T) {
	attr := etree.Attr{
		Space: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Key:   "link",
		Value: "rId1",
	}
	if !isRelAttr(attr) {
		t.Error("expected full-NS r:link to be a rel attr")
	}
}

// ---------------------------------------------------------------------------
// Phase 1: collectReferencedRIds tests
// ---------------------------------------------------------------------------

// makeElement is a test helper that builds an etree.Element from raw XML.
func makeElement(t *testing.T, xml string) *etree.Element {
	t.Helper()
	doc := etree.NewDocument()
	if err := doc.ReadFromString(xml); err != nil {
		t.Fatalf("makeElement: %v", err)
	}
	return doc.Root()
}

func TestCollectReferencedRIds_REmbed(t *testing.T) {
	// <a:blip r:embed="rId7"/> — typical embedded image reference.
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId7"/>`)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 1 || rids[0] != "rId7" {
		t.Errorf("expected [rId7], got %v", rids)
	}
}

func TestCollectReferencedRIds_RId(t *testing.T) {
	// <w:hyperlink r:id="rId3"> — external hyperlink.
	el := makeElement(t, `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3"><w:r><w:t>click</w:t></w:r></w:hyperlink>`)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 1 || rids[0] != "rId3" {
		t.Errorf("expected [rId3], got %v", rids)
	}
}

func TestCollectReferencedRIds_RLink(t *testing.T) {
	// <a:blip r:link="rId9"/> — linked (not embedded) image.
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rId9"/>`)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 1 || rids[0] != "rId9" {
		t.Errorf("expected [rId9], got %v", rids)
	}
}

func TestCollectReferencedRIds_Nested(t *testing.T) {
	// rId deep inside table → cell → paragraph → run → drawing → blipFill → blip.
	xml := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:tr><w:tc><w:p><w:r>
			<w:drawing><a:graphic><a:graphicData>
				<pic:blipFill xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<a:blip r:embed="rId42"/>
				</pic:blipFill>
			</a:graphicData></a:graphic></w:drawing>
		</w:r></w:p></w:tc></w:tr>
	</w:tbl>`
	el := makeElement(t, xml)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 1 || rids[0] != "rId42" {
		t.Errorf("expected [rId42], got %v", rids)
	}
}

func TestCollectReferencedRIds_NoDuplicates(t *testing.T) {
	// Two elements referencing the same rId → single entry in result.
	xml1 := `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"/>`
	xml2 := `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"/>`
	el1 := makeElement(t, xml1)
	el2 := makeElement(t, xml2)
	rids := collectReferencedRIds([]*etree.Element{el1, el2})
	if len(rids) != 1 {
		t.Errorf("expected 1 unique rId, got %d: %v", len(rids), rids)
	}
}

func TestCollectReferencedRIds_Multiple(t *testing.T) {
	// Two different rIds from two elements.
	xml1 := `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"/>`
	xml2 := `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId8"/>`
	el1 := makeElement(t, xml1)
	el2 := makeElement(t, xml2)
	rids := collectReferencedRIds([]*etree.Element{el1, el2})
	if len(rids) != 2 {
		t.Fatalf("expected 2 rIds, got %d: %v", len(rids), rids)
	}
	// Verify both present (order: rId5 first since el1 is processed first).
	want := map[string]bool{"rId5": true, "rId8": true}
	for _, rid := range rids {
		if !want[rid] {
			t.Errorf("unexpected rId: %s", rid)
		}
	}
}

func TestCollectReferencedRIds_Empty(t *testing.T) {
	rids := collectReferencedRIds(nil)
	if len(rids) != 0 {
		t.Errorf("expected empty, got %v", rids)
	}
}

func TestCollectReferencedRIds_EmptySlice(t *testing.T) {
	rids := collectReferencedRIds([]*etree.Element{})
	if len(rids) != 0 {
		t.Errorf("expected empty, got %v", rids)
	}
}

func TestCollectReferencedRIds_NoRelAttrs(t *testing.T) {
	// Element with only non-rel attributes.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>hello</w:t></w:r></w:p>`)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 0 {
		t.Errorf("expected empty, got %v", rids)
	}
}

func TestCollectReferencedRIds_EmptyValue(t *testing.T) {
	// r:embed with empty value — should be skipped.
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed=""/>`)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 0 {
		t.Errorf("expected empty for blank r:embed, got %v", rids)
	}
}

func TestCollectReferencedRIds_MixedAttrs(t *testing.T) {
	// Element with both r:embed and r:link on the same element.
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId10" r:link="rId11"/>`)
	rids := collectReferencedRIds([]*etree.Element{el})
	if len(rids) != 2 {
		t.Fatalf("expected 2 rIds, got %d: %v", len(rids), rids)
	}
	want := map[string]bool{"rId10": true, "rId11": true}
	for _, rid := range rids {
		if !want[rid] {
			t.Errorf("unexpected rId: %s", rid)
		}
	}
}

// ---------------------------------------------------------------------------
// Phase 1: remapRIds tests
// ---------------------------------------------------------------------------

func TestRemapRIds_REmbed(t *testing.T) {
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId7"/>`)
	remapRIds([]*etree.Element{el}, map[string]string{"rId7": "rId15"})
	val := attrValue(el, "r", "embed")
	if val != "rId15" {
		t.Errorf("expected rId15, got %s", val)
	}
}

func TestRemapRIds_RId(t *testing.T) {
	el := makeElement(t, `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3"/>`)
	remapRIds([]*etree.Element{el}, map[string]string{"rId3": "rId20"})
	val := attrValue(el, "r", "id")
	if val != "rId20" {
		t.Errorf("expected rId20, got %s", val)
	}
}

func TestRemapRIds_RLink(t *testing.T) {
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rId9"/>`)
	remapRIds([]*etree.Element{el}, map[string]string{"rId9": "rId30"})
	val := attrValue(el, "r", "link")
	if val != "rId30" {
		t.Errorf("expected rId30, got %s", val)
	}
}

func TestRemapRIds_AllTypes(t *testing.T) {
	// Single document with all three attribute types at different nesting levels.
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:hyperlink r:id="rId1">
			<w:r><w:drawing>
				<a:blip r:embed="rId2" r:link="rId3"/>
			</w:drawing></w:r>
		</w:hyperlink>
	</w:p>`
	el := makeElement(t, xml)
	ridMap := map[string]string{
		"rId1": "rId10",
		"rId2": "rId20",
		"rId3": "rId30",
	}
	remapRIds([]*etree.Element{el}, ridMap)

	// Verify hyperlink r:id.
	hl := findDescendant(el, "w", "hyperlink")
	if hl == nil {
		t.Fatal("hyperlink not found")
	}
	if v := attrValue(hl, "r", "id"); v != "rId10" {
		t.Errorf("hyperlink r:id: expected rId10, got %s", v)
	}

	// Verify blip r:embed and r:link.
	blip := findDescendant(el, "a", "blip")
	if blip == nil {
		t.Fatal("blip not found")
	}
	if v := attrValue(blip, "r", "embed"); v != "rId20" {
		t.Errorf("blip r:embed: expected rId20, got %s", v)
	}
	if v := attrValue(blip, "r", "link"); v != "rId30" {
		t.Errorf("blip r:link: expected rId30, got %s", v)
	}
}

func TestRemapRIds_UnknownLeftAlone(t *testing.T) {
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId99"/>`)
	// Map does not contain rId99 → should be left unchanged.
	remapRIds([]*etree.Element{el}, map[string]string{"rId1": "rId10"})
	val := attrValue(el, "r", "embed")
	if val != "rId99" {
		t.Errorf("expected rId99 unchanged, got %s", val)
	}
}

func TestRemapRIds_NestedElements(t *testing.T) {
	// Deeply nested blip inside table structure.
	xml := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:tr><w:tc><w:p><w:r>
			<w:drawing><a:graphic><a:graphicData>
				<a:blip r:embed="rId42"/>
			</a:graphicData></a:graphic></w:drawing>
		</w:r></w:p></w:tc></w:tr>
	</w:tbl>`
	el := makeElement(t, xml)
	remapRIds([]*etree.Element{el}, map[string]string{"rId42": "rId100"})

	blip := findDescendant(el, "a", "blip")
	if blip == nil {
		t.Fatal("blip not found")
	}
	if v := attrValue(blip, "r", "embed"); v != "rId100" {
		t.Errorf("expected rId100, got %s", v)
	}
}

func TestRemapRIds_EmptyMap(t *testing.T) {
	el := makeElement(t, `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"/>`)
	// Empty map → nothing changes.
	remapRIds([]*etree.Element{el}, map[string]string{})
	val := attrValue(el, "r", "embed")
	if val != "rId5" {
		t.Errorf("expected rId5 unchanged, got %s", val)
	}
}

func TestRemapRIds_NilElements(t *testing.T) {
	// Should not panic on nil/empty input.
	remapRIds(nil, map[string]string{"rId1": "rId2"})
	remapRIds([]*etree.Element{}, map[string]string{"rId1": "rId2"})
}

func TestRemapRIds_NonRelAttrsUntouched(t *testing.T) {
	// Ensure non-rel attributes with similar values are NOT touched.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:pPr><w:pStyle w:val="rId1"/></w:pPr></w:p>`)
	remapRIds([]*etree.Element{el}, map[string]string{"rId1": "rId99"})
	// w:val="rId1" should be untouched (it's not a rel attr).
	pStyle := findDescendant(el, "w", "pStyle")
	if pStyle == nil {
		t.Fatal("pStyle not found")
	}
	if v := attrValue(pStyle, "w", "val"); v != "rId1" {
		t.Errorf("w:val should be unchanged, got %s", v)
	}
}

func TestRemapRIds_MultipleElements(t *testing.T) {
	// Two separate root elements, each with an rId.
	xml1 := `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>`
	xml2 := `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>`
	el1 := makeElement(t, xml1)
	el2 := makeElement(t, xml2)
	ridMap := map[string]string{"rId1": "rId10", "rId2": "rId20"}
	remapRIds([]*etree.Element{el1, el2}, ridMap)

	if v := attrValue(el1, "r", "embed"); v != "rId10" {
		t.Errorf("el1 r:embed: expected rId10, got %s", v)
	}
	if v := attrValue(el2, "r", "id"); v != "rId20" {
		t.Errorf("el2 r:id: expected rId20, got %s", v)
	}
}

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

// attrValue returns the value of a namespaced attribute on el, or "".
func attrValue(el *etree.Element, space, key string) string {
	for _, attr := range el.Attr {
		if attr.Space == space && attr.Key == key {
			return attr.Value
		}
	}
	return ""
}

// findDescendant performs a depth-first search for the first descendant
// matching the given space and tag.
func findDescendant(el *etree.Element, space, tag string) *etree.Element {
	stack := el.ChildElements()
	for len(stack) > 0 {
		cur := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if cur.Space == space && cur.Tag == tag {
			return cur
		}
		stack = append(stack, cur.ChildElements()...)
	}
	return nil
}

// ---------------------------------------------------------------------------
// Phase 2: test infrastructure helpers
// ---------------------------------------------------------------------------

// newTestTargetParts creates a minimal WmlPackage + StoryPart for use as the
// "target" in importRelationship tests. The StoryPart has empty Rels with
// baseURI "/word".
func newTestTargetParts(t *testing.T) (*parts.WmlPackage, *parts.StoryPart) {
	t.Helper()
	opcPkg := opc.NewOpcPackage(nil)
	wmlPkg := parts.NewWmlPackage(opcPkg)
	opcPkg.SetAppPackage(wmlPkg)

	// Minimal StoryPart: XmlPart wrapping a dummy element.
	el := etree.NewElement("w:document")
	xp := opc.NewXmlPartFromElement("/word/document.xml", "application/xml", el, opcPkg)
	xp.SetRels(opc.NewRelationships("/word"))
	sp := parts.NewStoryPart(xp)
	return wmlPkg, sp
}

// newTestResourceImporter creates a ResourceImporter for tests.
func newTestResourceImporter(t *testing.T, source *Document) (*ResourceImporter, *parts.StoryPart) {
	t.Helper()
	wmlPkg, targetSP := newTestTargetParts(t)
	ri := newResourceImporter(source, nil, wmlPkg)
	return ri, targetSP
}

// newTestImagePart creates a tiny ImagePart with the given 1-byte blob for testing.
// Uses explicit meta to avoid needing a real image.
func newTestImagePart(t *testing.T, partName opc.PackURI, blob []byte) *parts.ImagePart {
	t.Helper()
	ip := parts.NewImagePartWithMeta(partName, "image/png", blob, 1, 1, 72, 72, "test.png")
	return ip
}

// newTestGenericPart creates a BasePart with the given partname and blob.
func newTestGenericPart(t *testing.T, partName opc.PackURI, contentType string, blob []byte) *opc.BasePart {
	t.Helper()
	return opc.NewBasePart(partName, contentType, blob, nil)
}

// ---------------------------------------------------------------------------
// Phase 2: partNameTemplate tests
// ---------------------------------------------------------------------------

func TestPartNameTemplate_WithDigits(t *testing.T) {
	got := partNameTemplate("/word/media/image3.png")
	want := "/word/media/image%d.png"
	if got != want {
		t.Errorf("partNameTemplate: got %q, want %q", got, want)
	}
}

func TestPartNameTemplate_NoDigits(t *testing.T) {
	got := partNameTemplate("/word/charts/chart.xml")
	want := "/word/charts/chart%d.xml"
	if got != want {
		t.Errorf("partNameTemplate: got %q, want %q", got, want)
	}
}

func TestPartNameTemplate_MultipleDigits(t *testing.T) {
	got := partNameTemplate("/word/media/image123.jpeg")
	want := "/word/media/image%d.jpeg"
	if got != want {
		t.Errorf("partNameTemplate: got %q, want %q", got, want)
	}
}

func TestPartNameTemplate_SingleDigit(t *testing.T) {
	got := partNameTemplate("/word/charts/chart1.xml")
	want := "/word/charts/chart%d.xml"
	if got != want {
		t.Errorf("partNameTemplate: got %q, want %q", got, want)
	}
}

func TestPartNameTemplate_NoExtension(t *testing.T) {
	got := partNameTemplate("/word/embeddings/oleObject1")
	want := "/word/embeddings/oleObject%d"
	if got != want {
		t.Errorf("partNameTemplate: got %q, want %q", got, want)
	}
}

// ---------------------------------------------------------------------------
// Phase 2: importRelationship tests
// ---------------------------------------------------------------------------

func TestImportRelationship_ExternalHyperlink(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)
	_ = wmlPkg

	srcRel := &opc.Relationship{
		RID:        "rId1",
		RelType:    opc.RTHyperlink,
		TargetRef:  "https://example.com",
		IsExternal: true,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId, err := importRelationship(srcRel, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if rId == "" {
		t.Fatal("expected non-empty rId")
	}

	// Verify the relationship was created in targetSP.
	rel := targetSP.Rels().GetByRID(rId)
	if rel == nil {
		t.Fatalf("relationship %s not found in target", rId)
	}
	if !rel.IsExternal {
		t.Error("expected external relationship")
	}
	if rel.TargetRef != "https://example.com" {
		t.Errorf("expected TargetRef https://example.com, got %s", rel.TargetRef)
	}
	if rel.RelType != opc.RTHyperlink {
		t.Errorf("expected RelType RTHyperlink, got %s", rel.RelType)
	}
}

func TestImportRelationship_ExternalHyperlinkDedup(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	srcRel := &opc.Relationship{
		RID:        "rId1",
		RelType:    opc.RTHyperlink,
		TargetRef:  "https://example.com",
		IsExternal: true,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId1, err := importRelationship(srcRel, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("first call: %v", err)
	}
	// Same URL again — should return the same rId.
	srcRel2 := &opc.Relationship{
		RID:        "rId5",
		RelType:    opc.RTHyperlink,
		TargetRef:  "https://example.com",
		IsExternal: true,
	}
	rId2, err := importRelationship(srcRel2, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("second call: %v", err)
	}
	if rId1 != rId2 {
		t.Errorf("expected same rId for same URL, got %s and %s", rId1, rId2)
	}
}

func TestImportRelationship_Image(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	// Create a source ImagePart with known content.
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47} // fake PNG header
	srcIP := newTestImagePart(t, "/word/media/image1.png", imgBlob)

	srcRel := &opc.Relationship{
		RID:        "rId7",
		RelType:    opc.RTImage,
		TargetRef:  "media/image1.png",
		TargetPart: srcIP,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId, err := importRelationship(srcRel, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if rId == "" {
		t.Fatal("expected non-empty rId")
	}

	// Verify: relationship exists in targetSP and points to an ImagePart.
	rel := targetSP.Rels().GetByRID(rId)
	if rel == nil {
		t.Fatalf("relationship %s not found", rId)
	}
	if rel.IsExternal {
		t.Error("expected internal relationship")
	}
	if rel.RelType != opc.RTImage {
		t.Errorf("expected RTImage, got %s", rel.RelType)
	}
	targetIP, ok := rel.TargetPart.(*parts.ImagePart)
	if !ok {
		t.Fatalf("expected *ImagePart, got %T", rel.TargetPart)
	}
	// Verify blob was copied.
	gotBlob, err := targetIP.Blob()
	if err != nil {
		t.Fatalf("reading target blob: %v", err)
	}
	if len(gotBlob) != len(imgBlob) {
		t.Errorf("blob length: got %d, want %d", len(gotBlob), len(imgBlob))
	}
}

func TestImportRelationship_ImageDedup(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	// Same blob → same SHA-256 → should dedup to one ImagePart.
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47}
	srcIP1 := newTestImagePart(t, "/word/media/image1.png", imgBlob)
	srcIP2 := newTestImagePart(t, "/word/media/image2.png", imgBlob) // same blob, different partname

	srcRel1 := &opc.Relationship{
		RID: "rId7", RelType: opc.RTImage,
		TargetRef: "media/image1.png", TargetPart: srcIP1,
	}
	srcRel2 := &opc.Relationship{
		RID: "rId8", RelType: opc.RTImage,
		TargetRef: "media/image2.png", TargetPart: srcIP2,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId1, err := importRelationship(srcRel1, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("first import: %v", err)
	}
	rId2, err := importRelationship(srcRel2, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("second import: %v", err)
	}

	// Both rIds should point to the same ImagePart (deduped by SHA-256).
	rel1 := targetSP.Rels().GetByRID(rId1)
	rel2 := targetSP.Rels().GetByRID(rId2)
	if rel1.TargetPart != rel2.TargetPart {
		t.Error("expected same ImagePart after SHA-256 dedup")
	}
	// rIds should also be the same since GetOrAdd returns existing rel.
	if rId1 != rId2 {
		t.Errorf("expected same rId for deduped image, got %s and %s", rId1, rId2)
	}

	// Only one ImagePart in WmlPackage.
	if wmlPkg.ImageParts().Len() != 1 {
		t.Errorf("expected 1 ImagePart in WmlPackage, got %d", wmlPkg.ImageParts().Len())
	}
}

func TestImportRelationship_InternalNilTargetPart(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	srcRel := &opc.Relationship{
		RID:        "rId1",
		RelType:    opc.RTImage,
		TargetRef:  "media/image1.png",
		IsExternal: false,
		TargetPart: nil, // oops — broken rel
	}
	importedParts := map[opc.PackURI]opc.Part{}

	_, err := importRelationship(srcRel, targetSP, wmlPkg, importedParts)
	if err == nil {
		t.Fatal("expected error for nil TargetPart")
	}
}

func TestImportRelationship_GenericPart(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	chartBlob := []byte("<c:chartSpace/>")
	srcPart := newTestGenericPart(t, "/word/charts/chart1.xml", "application/xml", chartBlob)
	srcPart.SetRels(opc.NewRelationships("/word/charts"))

	srcRel := &opc.Relationship{
		RID:        "rId10",
		RelType:    opc.RTChart,
		TargetRef:  "charts/chart1.xml",
		TargetPart: srcPart,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId, err := importRelationship(srcRel, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Verify relationship.
	rel := targetSP.Rels().GetByRID(rId)
	if rel == nil {
		t.Fatalf("relationship %s not found", rId)
	}
	if rel.RelType != opc.RTChart {
		t.Errorf("expected RTChart, got %s", rel.RelType)
	}
	if rel.IsExternal {
		t.Error("expected internal relationship")
	}

	// Verify blob was copied.
	gotBlob, err := rel.TargetPart.Blob()
	if err != nil {
		t.Fatalf("reading blob: %v", err)
	}
	if string(gotBlob) != string(chartBlob) {
		t.Errorf("blob mismatch: got %q, want %q", gotBlob, chartBlob)
	}

	// Verify part was added to OpcPackage.
	newPN := rel.TargetPart.PartName()
	if _, ok := wmlPkg.OpcPackage.PartByName(newPN); !ok {
		t.Errorf("new part %s not found in OpcPackage", newPN)
	}

	// Verify dedup map was populated.
	if _, ok := importedParts["/word/charts/chart1.xml"]; !ok {
		t.Error("expected importedParts to contain source partname")
	}
}

func TestImportRelationship_GenericPartDedup(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	// Same source part referenced by two different rIds.
	chartBlob := []byte("<c:chartSpace/>")
	srcPart := newTestGenericPart(t, "/word/charts/chart1.xml", "application/xml", chartBlob)
	srcPart.SetRels(opc.NewRelationships("/word/charts"))

	srcRel1 := &opc.Relationship{
		RID: "rId10", RelType: opc.RTChart,
		TargetRef: "charts/chart1.xml", TargetPart: srcPart,
	}
	srcRel2 := &opc.Relationship{
		RID: "rId11", RelType: opc.RTChart,
		TargetRef: "charts/chart1.xml", TargetPart: srcPart,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId1, err := importRelationship(srcRel1, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("first import: %v", err)
	}
	rId2, err := importRelationship(srcRel2, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("second import: %v", err)
	}

	// Both should point to the same target part (dedup by PartName).
	rel1 := targetSP.Rels().GetByRID(rId1)
	rel2 := targetSP.Rels().GetByRID(rId2)
	if rel1.TargetPart != rel2.TargetPart {
		t.Error("expected same target Part after generic part dedup")
	}

	// Same rId because GetOrAdd returns existing rel with same (relType, part).
	if rId1 != rId2 {
		t.Errorf("expected same rId for deduped generic part, got %s and %s", rId1, rId2)
	}

	// Only 1 entry in importedParts.
	if len(importedParts) != 1 {
		t.Errorf("expected 1 importedParts entry, got %d", len(importedParts))
	}
}

func TestImportRelationship_GenericPartNewPartname(t *testing.T) {
	wmlPkg, targetSP := newTestTargetParts(t)

	// Pre-populate target with /word/charts/chart1.xml so the new part gets chart2.
	existingPart := opc.NewBasePart("/word/charts/chart1.xml", "application/xml", []byte("existing"), wmlPkg.OpcPackage)
	wmlPkg.OpcPackage.AddPart(existingPart)

	srcPart := newTestGenericPart(t, "/word/charts/chart1.xml", "application/xml", []byte("new"))
	srcPart.SetRels(opc.NewRelationships("/word/charts"))

	srcRel := &opc.Relationship{
		RID: "rId10", RelType: opc.RTChart,
		TargetRef: "charts/chart1.xml", TargetPart: srcPart,
	}
	importedParts := map[opc.PackURI]opc.Part{}

	rId, err := importRelationship(srcRel, targetSP, wmlPkg, importedParts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	rel := targetSP.Rels().GetByRID(rId)
	newPN := rel.TargetPart.PartName()
	// Should NOT be chart1.xml (already taken) — should be chart2.xml.
	if newPN == "/word/charts/chart1.xml" {
		t.Error("expected new partname, got the same as existing")
	}
	if newPN != "/word/charts/chart2.xml" {
		t.Errorf("expected /word/charts/chart2.xml, got %s", newPN)
	}
}

// ---------------------------------------------------------------------------
// Phase 3: prepareContentElements tests
// ---------------------------------------------------------------------------

func TestPrepareContentElements_SkipsSectPr(t *testing.T) {
	source := mustNewDoc(t)
	// Default template has a body-level <w:sectPr>. Add a paragraph so
	// body is not empty after sectPr is stripped.
	source.AddParagraph("hello")

	ri, targetSP := newTestResourceImporter(t, source)

	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// None of the prepared elements should be <w:sectPr>.
	for _, el := range prep.elements {
		if el.Space == "w" && el.Tag == "sectPr" {
			t.Error("sectPr should have been filtered out")
		}
	}

	// We should have at least 1 element (the paragraph(s)).
	if len(prep.elements) == 0 {
		t.Error("expected at least 1 element after filtering sectPr")
	}
}

// ---------------------------------------------------------------------------
// sanitizeForInsertion tests
// ---------------------------------------------------------------------------

func TestSanitize_RemovesSectPrFromPPr(t *testing.T) {
	// Paragraph with pPr containing sectPr (with headerReference) and jc.
	// After sanitize: sectPr gone, jc preserved.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:pPr><w:jc w:val="center"/><w:sectPr><w:headerReference w:type="default" r:id="rId5"/></w:sectPr></w:pPr><w:r><w:t>text</w:t></w:r></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	// sectPr must be gone from pPr.
	pPr := el.FindElement("./w:pPr")
	if pPr == nil {
		t.Fatal("pPr disappeared")
	}
	for _, child := range pPr.ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			t.Error("sectPr should have been removed from pPr")
		}
	}
	// jc must survive.
	if pPr.FindElement("./w:jc") == nil {
		t.Error("jc should have been preserved in pPr")
	}
	// Run text must survive.
	if el.FindElement(".//w:t") == nil {
		t.Error("run text should have been preserved")
	}
}

func TestSanitize_SectPrOutsidePPrUntouched(t *testing.T) {
	// A sectPr directly inside tblPr (hypothetical) — must NOT be removed
	// because the rule is sectPr inside pPr only.
	el := makeElement(t, `<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sectPr/></w:tblPr>`)

	sanitizeForInsertion([]*etree.Element{el})

	if el.FindElement("./w:sectPr") == nil {
		t.Error("sectPr outside pPr should not have been removed")
	}
}

func TestSanitize_NoPPr(t *testing.T) {
	// Paragraph without pPr — should be a no-op.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>hello</w:t></w:r></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	if el.FindElement(".//w:t") == nil {
		t.Error("paragraph content should remain intact")
	}
}

func TestSanitize_RemovesCommentMarkers(t *testing.T) {
	// Paragraph with commentRangeStart, commentRangeEnd, and commentReference.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:commentRangeStart w:id="0"/><w:r><w:t>text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	for _, tag := range []string{"commentRangeStart", "commentRangeEnd", "commentReference"} {
		if findDescendant(el, "w", tag) != nil {
			t.Errorf("%s should have been removed", tag)
		}
	}
	if el.FindElement(".//w:t") == nil {
		t.Error("run text should have been preserved")
	}
}

func TestSanitize_RemovesBookmarkMarkers(t *testing.T) {
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bookmarkStart w:id="1" w:name="bm1"/><w:r><w:t>text</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	if findDescendant(el, "w", "bookmarkStart") != nil {
		t.Error("bookmarkStart should have been removed")
	}
	if findDescendant(el, "w", "bookmarkEnd") != nil {
		t.Error("bookmarkEnd should have been removed")
	}
	if el.FindElement(".//w:t") == nil {
		t.Error("run text should have been preserved")
	}
}

func TestSanitize_PreservesNonAnnotationElements(t *testing.T) {
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	if el.FindElement(".//w:b") == nil {
		t.Error("bold formatting should be preserved")
	}
	if el.FindElement(".//w:t") == nil {
		t.Error("text should be preserved")
	}
}

func TestSanitize_CombinedMarkup(t *testing.T) {
	// Paragraph with BOTH sectPr in pPr AND comment markers — all removed in one pass.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId5"/></w:sectPr></w:pPr><w:commentRangeStart w:id="0"/><w:bookmarkStart w:id="1" w:name="bm"/><w:r><w:t>text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:bookmarkEnd w:id="1"/></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	if findDescendant(el, "w", "sectPr") != nil {
		t.Error("sectPr should have been removed")
	}
	if findDescendant(el, "w", "commentRangeStart") != nil {
		t.Error("commentRangeStart should have been removed")
	}
	if findDescendant(el, "w", "bookmarkStart") != nil {
		t.Error("bookmarkStart should have been removed")
	}
	if el.FindElement(".//w:t") == nil {
		t.Error("run text should have been preserved")
	}
}

func TestSanitize_PreservesFootnoteAndEndnoteReferences(t *testing.T) {
	// Phase 4: footnoteReference/endnoteReference are no longer stripped
	// by sanitize — they are imported and remapped by ResourceImporter.
	// Stripping them would silently drop footnotes from the output.
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>text</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/></w:r></w:p>`)

	sanitizeForInsertion([]*etree.Element{el})

	if findDescendant(el, "w", "footnoteReference") == nil {
		t.Error("footnoteReference should be preserved (imported by Phase 4)")
	}
	if findDescendant(el, "w", "endnoteReference") == nil {
		t.Error("endnoteReference should be preserved (imported by Phase 4)")
	}
	if el.FindElement(".//w:t") == nil {
		t.Error("run text should have been preserved")
	}
}

func TestPrepareContentElements_StripsParagraphLevelSectPr(t *testing.T) {
	// Integration test: source with a paragraph-level sectPr containing
	// headerReference r:id. After prepareContentElements, the sectPr must be
	// gone and the headerReference r:id must NOT appear in any prepared element.
	source := mustNewDoc(t)

	// Inject a paragraph with pPr/sectPr/headerReference into source body.
	body := source.element.Body().RawElement()
	pEl := etree.NewElement("w:p")
	pEl.CreateAttr("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
	pPr := pEl.CreateElement("w:pPr")
	sectPr := pPr.CreateElement("w:sectPr")
	hdrRef := sectPr.CreateElement("w:headerReference")
	hdrRef.CreateAttr("w:type", "default")
	hdrRef.CreateAttr("r:id", "rId5")
	rEl := pEl.CreateElement("w:r")
	tEl := rEl.CreateElement("w:t")
	tEl.SetText("paragraph with sectPr")
	body.AddChild(pEl)

	ri, targetSP := newTestResourceImporter(t, source)

	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Verify no sectPr anywhere in prepared elements.
	for _, el := range prep.elements {
		if findDescendant(el, "w", "sectPr") != nil {
			t.Error("paragraph-level sectPr should have been stripped")
		}
	}

	// Verify no rId5 in any attribute (the headerReference rId should not
	// have been collected or remapped).
	for _, el := range prep.elements {
		stack := []*etree.Element{el}
		for len(stack) > 0 {
			cur := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			for _, attr := range cur.Attr {
				if attr.Value == "rId5" {
					t.Errorf("rId5 from headerReference should not appear in prepared content")
				}
			}
			stack = append(stack, cur.ChildElements()...)
		}
	}
}

// ---------------------------------------------------------------------------
// renumberDrawingIDs tests
// ---------------------------------------------------------------------------

func TestRenumberDrawingIDs_DocPrAndCNvPr(t *testing.T) {
	// Element tree with two bare numeric id attributes (wp:docPr and pic:cNvPr).
	el := makeElement(t, `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1"/><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="img.png"/></pic:nvPicPr></pic:pic></wp:inline></w:drawing></w:r></w:p>`)

	counter := 99
	nextID := func() int { counter++; return counter }

	renumberDrawingIDs([]*etree.Element{el}, nextID)

	// Both should be renumbered to unique values (100 and 101, order depends on DFS).
	docPr := findDescendant(el, "wp", "docPr")
	if docPr == nil {
		t.Fatal("docPr not found")
	}
	docPrID := docPr.SelectAttrValue("id", "")

	cNvPr := findDescendant(el, "pic", "cNvPr")
	if cNvPr == nil {
		t.Fatal("cNvPr not found")
	}
	cNvPrID := cNvPr.SelectAttrValue("id", "")

	// Both must be renumbered (not original values) and unique.
	if docPrID == "1" {
		t.Error("docPr id was not renumbered")
	}
	if cNvPrID == "0" {
		t.Error("cNvPr id was not renumbered")
	}
	if docPrID == cNvPrID {
		t.Errorf("docPr and cNvPr got the same id: %s", docPrID)
	}
}

func TestRenumberDrawingIDs_SkipsNamespacedId(t *testing.T) {
	// w:id attribute (namespaced) should NOT be renumbered.
	el := makeElement(t, `<w:bookmarkStart xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="42" w:name="bm"/>`)

	nextID := func() int { return 999 }
	renumberDrawingIDs([]*etree.Element{el}, nextID)

	// w:id should remain "42" (not renumbered).
	if v := attrValue(el, "w", "id"); v != "42" {
		t.Errorf("w:id changed to %s, expected 42 (should not be renumbered)", v)
	}
}

func TestRenumberDrawingIDs_SkipsNonNumericId(t *testing.T) {
	// Bare id with non-numeric value should NOT be renumbered.
	el := makeElement(t, `<foo id="abc" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)

	nextID := func() int { return 999 }
	renumberDrawingIDs([]*etree.Element{el}, nextID)

	if v := el.SelectAttrValue("id", ""); v != "abc" {
		t.Errorf("non-numeric id changed to %s, expected abc", v)
	}
}

func TestRenumberDrawingIDs_MultipleElements(t *testing.T) {
	el1 := makeElement(t, `<wp:docPr xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" id="1" name="P1"/>`)
	el2 := makeElement(t, `<wp:docPr xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" id="1" name="P2"/>`)

	counter := 0
	nextID := func() int { counter++; return counter }

	renumberDrawingIDs([]*etree.Element{el1, el2}, nextID)

	if v := el1.SelectAttrValue("id", ""); v != "1" {
		t.Errorf("el1 id: expected 1, got %s", v)
	}
	if v := el2.SelectAttrValue("id", ""); v != "2" {
		t.Errorf("el2 id: expected 2, got %s", v)
	}
}

func TestPrepareContentElements_EmptyBody(t *testing.T) {
	source := mustNewDoc(t)

	// Remove all children from body except sectPr to simulate empty content.
	body := source.element.Body().RawElement()
	var toRemove []*etree.Element
	for _, child := range body.ChildElements() {
		if !(child.Space == "w" && child.Tag == "sectPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		body.RemoveChild(child)
	}

	ri, targetSP := newTestResourceImporter(t, source)

	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(prep.elements) != 0 {
		t.Errorf("expected empty elements, got %d", len(prep.elements))
	}
}

func TestPrepareContentElements_PreservesOrder(t *testing.T) {
	source := mustNewDoc(t)

	// Clear body (except sectPr) and add elements in known order.
	body := source.element.Body().RawElement()
	var toRemove []*etree.Element
	for _, child := range body.ChildElements() {
		if !(child.Space == "w" && child.Tag == "sectPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		body.RemoveChild(child)
	}

	// Remove sectPr temporarily, add elements, then re-add sectPr at end.
	var sectPr *etree.Element
	for _, child := range body.ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			sectPr = child
			break
		}
	}
	if sectPr != nil {
		body.RemoveChild(sectPr)
	}

	// Insert: p("first"), tbl, p("second").
	p1 := body.CreateElement("w:p")
	r1 := p1.CreateElement("w:r")
	t1 := r1.CreateElement("w:t")
	t1.SetText("first")

	tbl := body.CreateElement("w:tbl")
	tbl.CreateElement("w:tr")

	p2 := body.CreateElement("w:p")
	r2 := p2.CreateElement("w:r")
	t2 := r2.CreateElement("w:t")
	t2.SetText("second")

	// Re-add sectPr at end (standard OOXML position).
	if sectPr != nil {
		body.AddChild(sectPr)
	}

	ri, targetSP := newTestResourceImporter(t, source)

	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if len(prep.elements) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(prep.elements))
	}
	if prep.elements[0].Tag != "p" {
		t.Errorf("element 0: expected p, got %s", prep.elements[0].Tag)
	}
	if prep.elements[1].Tag != "tbl" {
		t.Errorf("element 1: expected tbl, got %s", prep.elements[1].Tag)
	}
	if prep.elements[2].Tag != "p" {
		t.Errorf("element 2: expected p, got %s", prep.elements[2].Tag)
	}
}

func TestPrepareContentElements_DeepCopy(t *testing.T) {
	source := mustNewDoc(t)
	source.AddParagraph("original text")

	ri, targetSP := newTestResourceImporter(t, source)

	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Modify the prepared copy.
	for _, el := range prep.elements {
		for _, child := range el.ChildElements() {
			if child.Tag == "r" {
				for _, tc := range child.ChildElements() {
					if tc.Tag == "t" {
						tc.SetText("MODIFIED")
					}
				}
			}
		}
	}

	// Source body should be unmodified.
	srcBody := source.element.Body().RawElement()
	for _, child := range srcBody.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			for _, r := range child.ChildElements() {
				if r.Tag == "r" {
					for _, tc := range r.ChildElements() {
						if tc.Tag == "t" && tc.Text() == "MODIFIED" {
							t.Error("source body was modified — deep copy failed")
						}
					}
				}
			}
		}
	}
}

func TestPrepareContentElements_RemapsRIds(t *testing.T) {
	source := mustNewDoc(t)

	// Clear body and inject a paragraph with r:embed reference.
	body := source.element.Body().RawElement()
	var toRemove []*etree.Element
	for _, child := range body.ChildElements() {
		if !(child.Space == "w" && child.Tag == "sectPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		body.RemoveChild(child)
	}

	// Create: <w:p><w:r><w:drawing><a:blip r:embed="rId7"/></w:drawing></w:r></w:p>
	// Remove sectPr, add our paragraph, re-add sectPr.
	var sectPr *etree.Element
	for _, child := range body.ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			sectPr = child
			break
		}
	}
	if sectPr != nil {
		body.RemoveChild(sectPr)
	}

	p := body.CreateElement("w:p")
	r := p.CreateElement("w:r")
	drawing := r.CreateElement("w:drawing")
	blip := drawing.CreateElement("a:blip")
	blip.CreateAttr("r:embed", "rId7")

	if sectPr != nil {
		body.AddChild(sectPr)
	}

	// Add a source relationship: rId7 → image part.
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47}
	srcIP := newTestImagePart(t, "/word/media/image1.png", imgBlob)
	source.Part().Rels().Load("rId7", opc.RTImage, "media/image1.png", srcIP, false)

	ri, targetSP := newTestResourceImporter(t, source)

	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(prep.elements) == 0 {
		t.Fatal("expected at least 1 element")
	}

	// The r:embed value should no longer be "rId7" — it should be remapped
	// to whatever rId was assigned in the target.
	blipCopy := findDescendant(prep.elements[0], "a", "blip")
	if blipCopy == nil {
		t.Fatal("blip not found in prepared elements")
	}
	newRId := attrValue(blipCopy, "r", "embed")
	if newRId == "" {
		t.Fatal("r:embed is empty after remap")
	}
	if newRId == "rId7" {
		t.Error("r:embed was not remapped — still rId7")
	}

	// Verify the target has a relationship for the new rId.
	rel := targetSP.Rels().GetByRID(newRId)
	if rel == nil {
		t.Fatalf("target relationship %s not found", newRId)
	}
	if rel.RelType != opc.RTImage {
		t.Errorf("expected RTImage, got %s", rel.RelType)
	}

	// Source blip should still reference "rId7" (deep copy).
	srcBlip := findDescendant(p, "a", "blip")
	if srcBlip == nil {
		t.Fatal("source blip not found")
	}
	if v := attrValue(srcBlip, "r", "embed"); v != "rId7" {
		t.Errorf("source r:embed changed to %s, expected rId7", v)
	}
}

func TestPrepareContentElements_OrphanedRIdSkipped(t *testing.T) {
	source := mustNewDoc(t)

	// Inject paragraph with r:embed pointing to non-existent relationship.
	body := source.element.Body().RawElement()
	var toRemove []*etree.Element
	for _, child := range body.ChildElements() {
		if !(child.Space == "w" && child.Tag == "sectPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		body.RemoveChild(child)
	}

	var sectPr *etree.Element
	for _, child := range body.ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			sectPr = child
			break
		}
	}
	if sectPr != nil {
		body.RemoveChild(sectPr)
	}

	p := body.CreateElement("w:p")
	r := p.CreateElement("w:r")
	drawing := r.CreateElement("w:drawing")
	blip := drawing.CreateElement("a:blip")
	blip.CreateAttr("r:embed", "rId999") // no such relationship in source

	if sectPr != nil {
		body.AddChild(sectPr)
	}

	ri, targetSP := newTestResourceImporter(t, source)

	// Should NOT error — orphaned references are skipped silently.
	prep, err := prepareContentElements(source, targetSP, ri)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(prep.elements) != 1 {
		t.Fatalf("expected 1 element, got %d", len(prep.elements))
	}

	// The r:embed should remain "rId999" (no remap happened).
	blipCopy := findDescendant(prep.elements[0], "a", "blip")
	if blipCopy == nil {
		t.Fatal("blip not found")
	}
	if v := attrValue(blipCopy, "r", "embed"); v != "rId999" {
		t.Errorf("orphaned r:embed changed to %s, expected rId999", v)
	}
}
