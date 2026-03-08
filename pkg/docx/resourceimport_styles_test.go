package docx

import (
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// collectStyleIdsFromElements tests
// --------------------------------------------------------------------------

func TestCollectStyleIdsFromElements_ParagraphStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
		`<w:r><w:t>Title</w:t></w:r>` +
		`</w:p>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}

	ids := collectStyleIdsFromElements([]*etree.Element{el})
	if len(ids) != 1 {
		t.Fatalf("expected 1 styleId, got %d", len(ids))
	}
	if ids[0] != "Heading1" {
		t.Errorf("expected Heading1, got %s", ids[0])
	}
}

func TestCollectStyleIdsFromElements_RunStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>bold</w:t></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectStyleIdsFromElements([]*etree.Element{el})
	if len(ids) != 1 || ids[0] != "Strong" {
		t.Errorf("expected [Strong], got %v", ids)
	}
}

func TestCollectStyleIdsFromElements_TableStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>` +
		`</w:tbl>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectStyleIdsFromElements([]*etree.Element{el})
	if len(ids) != 1 || ids[0] != "TableGrid" {
		t.Errorf("expected [TableGrid], got %v", ids)
	}
}

func TestCollectStyleIdsFromElements_Multiple(t *testing.T) {
	t.Parallel()
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>` +
		`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>` +
		`<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>x</w:t></w:r></w:p>` +
		`<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr></w:tbl>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectStyleIdsFromElements(el.ChildElements())
	if len(ids) != 4 {
		t.Fatalf("expected 4 unique styleIds, got %d: %v", len(ids), ids)
	}
	// Verify document order.
	expected := []string{"Heading1", "Normal", "Strong", "TableGrid"}
	for i, e := range expected {
		if ids[i] != e {
			t.Errorf("ids[%d]: expected %s, got %s", i, e, ids[i])
		}
	}
}

func TestCollectStyleIdsFromElements_Dedup(t *testing.T) {
	t.Parallel()
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>` +
		`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectStyleIdsFromElements(el.ChildElements())
	if len(ids) != 1 {
		t.Fatalf("expected 1 unique styleId, got %d", len(ids))
	}
}

func TestCollectStyleIdsFromElements_Empty(t *testing.T) {
	t.Parallel()
	ids := collectStyleIdsFromElements(nil)
	if len(ids) != 0 {
		t.Errorf("expected 0 styleIds for nil elements, got %d", len(ids))
	}
}

// --------------------------------------------------------------------------
// collectStyleClosure tests
// --------------------------------------------------------------------------

// buildStylesXml builds a minimal <w:styles> element containing the given
// style definitions. Each def is a raw XML string for a <w:style> element.
func buildStylesXml(defs ...string) *oxml.CT_Styles {
	xml := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
	for _, d := range defs {
		xml += d
	}
	xml += `</w:styles>`
	el, _ := oxml.ParseXml([]byte(xml))
	return &oxml.CT_Styles{Element: oxml.WrapElement(el)}
}

func TestCollectStyleClosure_BasedOnChain(t *testing.T) {
	t.Parallel()
	// Chain: CustomBody → BodyText → Normal (3 levels).
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`,
		`<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/></w:style>`,
		`<w:style w:type="paragraph" w:styleId="CustomBody"><w:name w:val="Custom Body"/><w:basedOn w:val="BodyText"/></w:style>`,
	)

	// Minimal ResourceImporter with mocked source styles access.
	ri := &ResourceImporter{
		styleMap: map[string]string{},
	}
	// We can't easily mock sourceStyles(), so test collectStyleClosure
	// by calling it directly via a helper that takes CT_Styles.
	closure := collectStyleClosureFrom(srcStyles, []string{"CustomBody"})

	if len(closure) != 3 {
		t.Fatalf("expected 3 styles in closure, got %d", len(closure))
	}
	// BFS order: CustomBody, BodyText, Normal.
	ids := make([]string, len(closure))
	for i, s := range closure {
		ids[i] = s.StyleId()
	}
	expected := []string{"CustomBody", "BodyText", "Normal"}
	for i, e := range expected {
		if ids[i] != e {
			t.Errorf("closure[%d]: expected %s, got %s", i, e, ids[i])
		}
	}
	_ = ri // silence unused
}

func TestCollectStyleClosure_WithLink(t *testing.T) {
	t.Parallel()
	// Paragraph style with linked character style.
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:link w:val="Heading1Char"/></w:style>`,
		`<w:style w:type="character" w:styleId="Heading1Char"><w:name w:val="Heading 1 Char"/><w:link w:val="Heading1"/></w:style>`,
	)

	closure := collectStyleClosureFrom(srcStyles, []string{"Heading1"})

	if len(closure) != 2 {
		t.Fatalf("expected 2 styles (paragraph + linked char), got %d", len(closure))
	}
}

func TestCollectStyleClosure_WithNext(t *testing.T) {
	t.Parallel()
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:next w:val="Normal"/></w:style>`,
		`<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`,
	)

	closure := collectStyleClosureFrom(srcStyles, []string{"Heading1"})

	if len(closure) != 2 {
		t.Fatalf("expected 2 styles (heading + next), got %d", len(closure))
	}
}

func TestCollectStyleClosure_OrphanedReference(t *testing.T) {
	t.Parallel()
	// Reference to a style that doesn't exist in source.
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Custom"><w:name w:val="Custom"/><w:basedOn w:val="NonExistent"/></w:style>`,
	)

	closure := collectStyleClosureFrom(srcStyles, []string{"Custom"})

	// Only Custom found; NonExistent silently skipped.
	if len(closure) != 1 {
		t.Fatalf("expected 1 style, got %d", len(closure))
	}
}

func TestCollectStyleClosure_CircularDependency(t *testing.T) {
	t.Parallel()
	// Circular: A → B → A (shouldn't happen in valid OOXML, but must not loop).
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/></w:style>`,
		`<w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/></w:style>`,
	)

	closure := collectStyleClosureFrom(srcStyles, []string{"A"})

	if len(closure) != 2 {
		t.Fatalf("expected 2 styles (no infinite loop), got %d", len(closure))
	}
}

// collectStyleClosureFrom is a test helper that runs BFS closure without
// needing a full ResourceImporter. Mirrors collectStyleClosure logic.
func collectStyleClosureFrom(srcStyles *oxml.CT_Styles, seedIds []string) []*oxml.CT_Style {
	queue := make([]string, len(seedIds))
	copy(queue, seedIds)
	visited := map[string]bool{}
	var result []*oxml.CT_Style

	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		if visited[id] {
			continue
		}
		visited[id] = true

		s := srcStyles.GetByID(id)
		if s == nil {
			continue
		}
		result = append(result, s)

		if v, _ := s.BasedOnVal(); v != "" {
			queue = append(queue, v)
		}
		if v, _ := s.NextVal(); v != "" {
			queue = append(queue, v)
		}
		if link := s.RawElement().FindElement("w:link"); link != nil {
			if v := link.SelectAttrValue("w:val", ""); v != "" {
				queue = append(queue, v)
			}
		}
	}
	return result
}

// --------------------------------------------------------------------------
// remapAll style tests
// --------------------------------------------------------------------------

func TestRemapAll_PStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="OldStyle"/></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"OldStyle": "NewStyle"},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	pStyle := el.FindElement("//pStyle")
	if pStyle == nil {
		t.Fatal("pStyle element not found after remap")
	}
	got := pStyle.SelectAttrValue("w:val", "")
	if got != "NewStyle" {
		t.Errorf("expected pStyle remapped to NewStyle, got %s", got)
	}
}

func TestRemapAll_RStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"Strong": "BoldChar"},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	rStyle := el.FindElement("//rStyle")
	got := rStyle.SelectAttrValue("w:val", "")
	if got != "BoldChar" {
		t.Errorf("expected rStyle remapped to BoldChar, got %s", got)
	}
}

func TestRemapAll_TblStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>` +
		`</w:tbl>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"TableGrid": "CustomTable"},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	tblStyle := el.FindElement("//tblStyle")
	got := tblStyle.SelectAttrValue("w:val", "")
	if got != "CustomTable" {
		t.Errorf("expected tblStyle remapped to CustomTable, got %s", got)
	}
}

func TestRemapAll_StyleNotInMap(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Normal"/></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"Other": "Renamed"},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	pStyle := el.FindElement("//pStyle")
	got := pStyle.SelectAttrValue("w:val", "")
	if got != "Normal" {
		t.Errorf("expected pStyle to stay Normal, got %s", got)
	}
}

func TestRemapAll_CombinedNumIdAndStyle(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr>` +
		`<w:pStyle w:val="ListParagraph"/>` +
		`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>` +
		`</w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{5: 42},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"ListParagraph": "CustomList"},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	pStyle := el.FindElement("//pStyle")
	if got := pStyle.SelectAttrValue("w:val", ""); got != "CustomList" {
		t.Errorf("pStyle: expected CustomList, got %s", got)
	}
	numId := el.FindElement("//numId")
	if got := numId.SelectAttrValue("w:val", ""); got != "42" {
		t.Errorf("numId: expected 42, got %s", got)
	}
}

// --------------------------------------------------------------------------
// remapNumIdsInElement tests
// --------------------------------------------------------------------------

func TestRemapNumIdsInElement(t *testing.T) {
	t.Parallel()
	xml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:styleId="ListBullet">` +
		`<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>` +
		`</w:style>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{3: 99},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{},
	}

	ri.remapNumIdsInElement(el)

	numId := el.FindElement("//numId")
	if numId == nil {
		t.Fatal("numId not found")
	}
	if got := numId.SelectAttrValue("w:val", ""); got != "99" {
		t.Errorf("expected numId remapped to 99, got %s", got)
	}
}

func TestRemapNumIdsInElement_EmptyMap(t *testing.T) {
	t.Parallel()
	xml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:styleId="ListBullet">` +
		`<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>` +
		`</w:style>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{},
	}

	ri.remapNumIdsInElement(el)

	numId := el.FindElement("//numId")
	if got := numId.SelectAttrValue("w:val", ""); got != "3" {
		t.Errorf("expected numId to stay 3, got %s", got)
	}
}

// --------------------------------------------------------------------------
// remapStyleRefsInElement tests
// --------------------------------------------------------------------------

func TestRemapStyleRefsInElement(t *testing.T) {
	t.Parallel()
	xml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:styleId="Child">` +
		`<w:basedOn w:val="Parent"/>` +
		`<w:next w:val="Normal"/>` +
		`<w:link w:val="ChildChar"/>` +
		`</w:style>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{},
		absNumIdMap: map[int]int{},
		styleMap: map[string]string{
			"Parent":    "RenamedParent",
			"Normal":    "Normal",
			"ChildChar": "RenamedChildChar",
		},
	}

	ri.remapStyleRefsInElement(el)

	basedOn := el.FindElement("basedOn")
	if got := basedOn.SelectAttrValue("w:val", ""); got != "RenamedParent" {
		t.Errorf("basedOn: expected RenamedParent, got %s", got)
	}
	next := el.FindElement("next")
	if got := next.SelectAttrValue("w:val", ""); got != "Normal" {
		t.Errorf("next: expected Normal (identity), got %s", got)
	}
	link := el.FindElement("link")
	if got := link.SelectAttrValue("w:val", ""); got != "RenamedChildChar" {
		t.Errorf("link: expected RenamedChildChar, got %s", got)
	}
}

func TestRemapStyleRefsInElement_NotInMap(t *testing.T) {
	t.Parallel()
	xml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:styleId="Custom">` +
		`<w:basedOn w:val="Unknown"/>` +
		`</w:style>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{"Other": "Mapped"},
	}

	ri.remapStyleRefsInElement(el)

	basedOn := el.FindElement("basedOn")
	if got := basedOn.SelectAttrValue("w:val", ""); got != "Unknown" {
		t.Errorf("expected basedOn to stay Unknown, got %s", got)
	}
}

func TestRemapStyleRefsInElement_EmptyMap(t *testing.T) {
	t.Parallel()
	xml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:styleId="Custom">` +
		`<w:basedOn w:val="Parent"/>` +
		`</w:style>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{},
	}

	ri.remapStyleRefsInElement(el)

	basedOn := el.FindElement("basedOn")
	if got := basedOn.SelectAttrValue("w:val", ""); got != "Parent" {
		t.Errorf("expected basedOn to stay Parent, got %s", got)
	}
}

// --------------------------------------------------------------------------
// mergePropertiesDeep tests
// --------------------------------------------------------------------------

func TestMergePropertiesDeep_AddsMissingChildren(t *testing.T) {
	t.Parallel()
	// dst has jc; src has ind. After merge, dst should have both.
	dstXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>`
	srcXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:ind w:left="720"/>` +
		`</w:pPr>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	mergePropertiesDeep(dst, src)

	if findChild(dst, "w", "jc") == nil {
		t.Error("jc should still be present")
	}
	ind := findChild(dst, "w", "ind")
	if ind == nil {
		t.Fatal("ind should have been added from src")
	}
	if got := ind.SelectAttrValue("w:left", ""); got != "720" {
		t.Errorf("expected ind w:left=720, got %s", got)
	}
}

func TestMergePropertiesDeep_DstTakesPrecedence(t *testing.T) {
	t.Parallel()
	// Both have jc with different values. dst value must win.
	dstXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>`
	srcXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="left"/>` +
		`</w:pPr>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	mergePropertiesDeep(dst, src)

	jc := findChild(dst, "w", "jc")
	if jc == nil {
		t.Fatal("jc should be present")
	}
	if got := jc.SelectAttrValue("w:val", ""); got != "center" {
		t.Errorf("expected center (dst wins), got %s", got)
	}
}

func TestMergePropertiesDeep_MergesAttributes(t *testing.T) {
	t.Parallel()
	// rFonts: dst has w:ascii, src has w:hAnsi. Result should have both.
	dstXml := `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rFonts w:ascii="Arial"/>` +
		`</w:rPr>`
	srcXml := `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rFonts w:hAnsi="Times"/>` +
		`</w:rPr>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	mergePropertiesDeep(dst, src)

	rf := findChild(dst, "w", "rFonts")
	if rf == nil {
		t.Fatal("rFonts should be present")
	}
	if got := rf.SelectAttrValue("w:ascii", ""); got != "Arial" {
		t.Errorf("w:ascii: expected Arial, got %s", got)
	}
	if got := rf.SelectAttrValue("w:hAnsi", ""); got != "Times" {
		t.Errorf("w:hAnsi: expected Times, got %s", got)
	}
}

func TestMergePropertiesDeep_EmptySrc(t *testing.T) {
	t.Parallel()
	dstXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>`
	srcXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>` +
		``
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	mergePropertiesDeep(dst, src)

	if findChild(dst, "w", "jc") == nil {
		t.Error("jc should still be present")
	}
	if len(dst.ChildElements()) != 1 {
		t.Errorf("expected 1 child, got %d", len(dst.ChildElements()))
	}
}

// --------------------------------------------------------------------------
// overridePropertiesDeep tests
// --------------------------------------------------------------------------

func TestOverridePropertiesDeep_SrcOverridesDst(t *testing.T) {
	t.Parallel()
	// Both have jc. src (derived style) should override dst (base style).
	dstXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="left"/>` +
		`</w:pPr>`
	srcXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	overridePropertiesDeep(dst, src)

	jc := findChild(dst, "w", "jc")
	if got := jc.SelectAttrValue("w:val", ""); got != "center" {
		t.Errorf("expected center (src overrides), got %s", got)
	}
}

func TestOverridePropertiesDeep_AddsMissing(t *testing.T) {
	t.Parallel()
	dstXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="left"/>` +
		`</w:pPr>`
	srcXml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:ind w:left="360"/>` +
		`</w:pPr>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	overridePropertiesDeep(dst, src)

	if findChild(dst, "w", "jc") == nil {
		t.Error("jc should still be present")
	}
	ind := findChild(dst, "w", "ind")
	if ind == nil {
		t.Fatal("ind should be added from src")
	}
	if got := ind.SelectAttrValue("w:left", ""); got != "360" {
		t.Errorf("expected w:left=360, got %s", got)
	}
}

// --------------------------------------------------------------------------
// mergeAttrs tests
// --------------------------------------------------------------------------

func TestMergeAttrs_AddsNonExisting(t *testing.T) {
	t.Parallel()
	dstXml := `<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Arial"/>`
	srcXml := `<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:hAnsi="Times" w:cs="Noto"/>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	mergeAttrs(dst, src)

	if got := dst.SelectAttrValue("w:ascii", ""); got != "Arial" {
		t.Errorf("w:ascii: expected Arial, got %s", got)
	}
	if got := dst.SelectAttrValue("w:hAnsi", ""); got != "Times" {
		t.Errorf("w:hAnsi: expected Times (added), got %s", got)
	}
	if got := dst.SelectAttrValue("w:cs", ""); got != "Noto" {
		t.Errorf("w:cs: expected Noto (added), got %s", got)
	}
}

func TestMergeAttrs_DoesNotOverwrite(t *testing.T) {
	t.Parallel()
	dstXml := `<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="24"/>`
	srcXml := `<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="48"/>`
	dst, _ := oxml.ParseXml([]byte(dstXml))
	src, _ := oxml.ParseXml([]byte(srcXml))

	mergeAttrs(dst, src)

	if got := dst.SelectAttrValue("w:val", ""); got != "24" {
		t.Errorf("w:val: expected 24 (dst wins), got %s", got)
	}
}

// --------------------------------------------------------------------------
// resolveStyleChain tests (via resolveStyleChainFrom helper)
// --------------------------------------------------------------------------

// resolveStyleChainFrom is a test helper that resolves a style chain from
// a CT_Styles object without requiring a full Document/ResourceImporter.
func resolveStyleChainFrom(styles *oxml.CT_Styles, style *oxml.CT_Style) (pPr, rPr *etree.Element) {
	var chain []*oxml.CT_Style
	visited := map[string]bool{}
	current := style
	for current != nil {
		id := current.StyleId()
		if visited[id] {
			break
		}
		visited[id] = true
		chain = append(chain, current)
		basedOn, _ := current.BasedOnVal()
		if basedOn == "" {
			break
		}
		current = styles.GetByID(basedOn)
	}
	for i := len(chain) - 1; i >= 0; i-- {
		raw := chain[i].RawElement()
		if p := findChild(raw, "w", "pPr"); p != nil {
			if pPr == nil {
				pPr = p.Copy()
			} else {
				overridePropertiesDeep(pPr, p)
			}
		}
		if r := findChild(raw, "w", "rPr"); r != nil {
			if rPr == nil {
				rPr = r.Copy()
			} else {
				overridePropertiesDeep(rPr, r)
			}
		}
	}
	if pPr != nil {
		removeChild(pPr, "w", "pStyle")
	}
	if rPr != nil {
		removeChild(rPr, "w", "rStyle")
	}
	return
}

func TestResolveStyleChain_SingleStyle(t *testing.T) {
	t.Parallel()
	styles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Custom">` +
			`<w:name w:val="Custom"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	)
	style := styles.GetByID("Custom")
	pPr, rPr := resolveStyleChainFrom(styles, style)

	if pPr == nil {
		t.Fatal("pPr should not be nil")
	}
	if findChild(pPr, "w", "jc") == nil {
		t.Error("resolved pPr should have jc")
	}
	if rPr == nil {
		t.Fatal("rPr should not be nil")
	}
	if findChild(rPr, "w", "b") == nil {
		t.Error("resolved rPr should have b")
	}
}

func TestResolveStyleChain_DerivedOverridesBase(t *testing.T) {
	t.Parallel()
	styles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Base">` +
			`<w:name w:val="Base"/>` +
			`<w:pPr><w:jc w:val="left"/><w:ind w:left="720"/></w:pPr>` +
			`<w:rPr><w:sz w:val="20"/></w:rPr>` +
			`</w:style>`,
		`<w:style w:type="paragraph" w:styleId="Derived">` +
			`<w:name w:val="Derived"/>` +
			`<w:basedOn w:val="Base"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`<w:rPr><w:b/></w:rPr>` +
			`</w:style>`,
	)
	style := styles.GetByID("Derived")
	pPr, rPr := resolveStyleChainFrom(styles, style)

	if pPr == nil {
		t.Fatal("pPr should not be nil")
	}
	// jc: derived = center should override base = left.
	jc := findChild(pPr, "w", "jc")
	if jc == nil {
		t.Fatal("jc should be present")
	}
	if got := jc.SelectAttrValue("w:val", ""); got != "center" {
		t.Errorf("jc: expected center (derived overrides), got %s", got)
	}
	// ind: only in base, should be inherited.
	ind := findChild(pPr, "w", "ind")
	if ind == nil {
		t.Fatal("ind should be inherited from base")
	}
	if got := ind.SelectAttrValue("w:left", ""); got != "720" {
		t.Errorf("ind: expected 720, got %s", got)
	}

	if rPr == nil {
		t.Fatal("rPr should not be nil")
	}
	// sz from base, b from derived.
	if findChild(rPr, "w", "sz") == nil {
		t.Error("sz should be inherited from base")
	}
	if findChild(rPr, "w", "b") == nil {
		t.Error("b should be present from derived")
	}
}

func TestResolveStyleChain_ThreeLevels(t *testing.T) {
	t.Parallel()
	styles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Root">` +
			`<w:name w:val="Root"/>` +
			`<w:pPr><w:jc w:val="left"/></w:pPr>` +
			`<w:rPr><w:sz w:val="20"/><w:rFonts w:ascii="Times"/></w:rPr>` +
			`</w:style>`,
		`<w:style w:type="paragraph" w:styleId="Middle">` +
			`<w:name w:val="Middle"/>` +
			`<w:basedOn w:val="Root"/>` +
			`<w:rPr><w:sz w:val="24"/></w:rPr>` +
			`</w:style>`,
		`<w:style w:type="paragraph" w:styleId="Leaf">` +
			`<w:name w:val="Leaf"/>` +
			`<w:basedOn w:val="Middle"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
	)
	style := styles.GetByID("Leaf")
	pPr, rPr := resolveStyleChainFrom(styles, style)

	// jc: Root=left, Leaf=center → center wins
	jc := findChild(pPr, "w", "jc")
	if got := jc.SelectAttrValue("w:val", ""); got != "center" {
		t.Errorf("jc: expected center, got %s", got)
	}

	// sz: Root=20, Middle=24 → 24 wins
	sz := findChild(rPr, "w", "sz")
	if got := sz.SelectAttrValue("w:val", ""); got != "24" {
		t.Errorf("sz: expected 24 (middle overrides root), got %s", got)
	}

	// rFonts: only in Root → inherited
	rf := findChild(rPr, "w", "rFonts")
	if rf == nil {
		t.Error("rFonts should be inherited from Root")
	}
}

func TestResolveStyleChain_CircularProtection(t *testing.T) {
	t.Parallel()
	styles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="A">` +
			`<w:name w:val="A"/>` +
			`<w:basedOn w:val="B"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
		`<w:style w:type="paragraph" w:styleId="B">` +
			`<w:name w:val="B"/>` +
			`<w:basedOn w:val="A"/>` +
			`<w:pPr><w:ind w:left="360"/></w:pPr>` +
			`</w:style>`,
	)
	style := styles.GetByID("A")
	pPr, _ := resolveStyleChainFrom(styles, style)

	// Should terminate without infinite loop and produce merged result.
	if pPr == nil {
		t.Fatal("pPr should not be nil")
	}
	if findChild(pPr, "w", "jc") == nil {
		t.Error("jc should be present")
	}
}

func TestResolveStyleChain_StripsPStyleRStyle(t *testing.T) {
	t.Parallel()
	// pPr in style definition sometimes contains pStyle (for inheritance).
	// The resolved result should strip it — pStyle is not a direct attr.
	styles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Custom">` +
			`<w:name w:val="Custom"/>` +
			`<w:pPr><w:pStyle w:val="Custom"/><w:jc w:val="center"/></w:pPr>` +
			`<w:rPr><w:rStyle w:val="CustomChar"/><w:b/></w:rPr>` +
			`</w:style>`,
	)
	style := styles.GetByID("Custom")
	pPr, rPr := resolveStyleChainFrom(styles, style)

	if findChild(pPr, "w", "pStyle") != nil {
		t.Error("pStyle should be stripped from resolved pPr")
	}
	if findChild(rPr, "w", "rStyle") != nil {
		t.Error("rStyle should be stripped from resolved rPr")
	}
	// Actual properties should still be there.
	if findChild(pPr, "w", "jc") == nil {
		t.Error("jc should still be present")
	}
	if findChild(rPr, "w", "b") == nil {
		t.Error("b should still be present")
	}
}

// --------------------------------------------------------------------------
// expandDirectFormatting tests
// --------------------------------------------------------------------------

// newExpandTestRI creates a ResourceImporter for expand tests.
// It mocks sourceStyles by using the provided CT_Styles through a minimal
// Document setup.
func newExpandTestRI(expandMap map[string]*oxml.CT_Style) *ResourceImporter {
	return &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{},
		expandStyles:  expandMap,
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}
}

func TestExpandDirectFormatting_EmptyMap(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
		`<w:r><w:t>text</w:t></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := newExpandTestRI(map[string]*oxml.CT_Style{})
	ri.expandDirectFormatting([]*etree.Element{el})

	// No expandStyles → nothing should change.
	pPr := findChild(el, "w", "pPr")
	if len(pPr.ChildElements()) != 1 {
		t.Errorf("pPr should still have only pStyle, got %d children", len(pPr.ChildElements()))
	}
}

func TestExpandDirectFormatting_ParagraphStyleExpansion(t *testing.T) {
	t.Parallel()
	// Source style: Heading1 with jc=center and bold rPr.
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1">` +
			`<w:name w:val="heading 1"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>` +
			`</w:style>`,
	)
	heading1 := srcStyles.GetByID("Heading1")

	// Document paragraph referencing Heading1 with no direct formatting.
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
		`<w:r><w:t>Title</w:t></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := newExpandTestRI(map[string]*oxml.CT_Style{
		"Heading1": heading1,
	})
	// Mock sourceDoc with styles access.
	ri.sourceDoc = &Document{}
	ri.sourceDoc = nil // sourceStyles will fail, but we call resolveStyleChain via expand

	// Since we can't easily mock sourceDoc.part.Styles(), test the lower level.
	// Manually resolve and apply.
	pPr := findChild(el, "w", "pPr")

	// Simulate what expandParagraphStyle does when resolveStyleChain returns properties.
	resolvedPPr, resolvedRPr := resolveStyleChainFrom(srcStyles, heading1)
	if resolvedPPr != nil {
		mergePropertiesDeep(pPr, resolvedPPr)
	}
	if resolvedRPr != nil {
		existingRPr := findChild(pPr, "w", "rPr")
		if existingRPr == nil {
			existingRPr = etree.NewElement("w:rPr")
			pPr.AddChild(existingRPr)
		}
		mergePropertiesDeep(existingRPr, resolvedRPr)
	}

	// Verify jc was added.
	jc := findChild(pPr, "w", "jc")
	if jc == nil {
		t.Fatal("jc should have been expanded from style")
	}
	if got := jc.SelectAttrValue("w:val", ""); got != "center" {
		t.Errorf("jc: expected center, got %s", got)
	}

	// Verify rPr was created with b and sz.
	rPr := findChild(pPr, "w", "rPr")
	if rPr == nil {
		t.Fatal("rPr should have been created")
	}
	if findChild(rPr, "w", "b") == nil {
		t.Error("rPr should have b from style")
	}
	if sz := findChild(rPr, "w", "sz"); sz == nil {
		t.Error("rPr should have sz from style")
	} else if got := sz.SelectAttrValue("w:val", ""); got != "28" {
		t.Errorf("sz: expected 28, got %s", got)
	}
}

func TestExpandDirectFormatting_DirectFormattingTakesPrecedence(t *testing.T) {
	t.Parallel()
	// Style has jc=center, but paragraph already has jc=right (direct).
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1">` +
			`<w:name w:val="heading 1"/>` +
			`<w:pPr><w:jc w:val="center"/><w:ind w:left="720"/></w:pPr>` +
			`</w:style>`,
	)
	heading1 := srcStyles.GetByID("Heading1")

	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="right"/></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))
	pPr := findChild(el, "w", "pPr")

	resolvedPPr, _ := resolveStyleChainFrom(srcStyles, heading1)
	mergePropertiesDeep(pPr, resolvedPPr)

	// jc: direct=right should win over style=center.
	jc := findChild(pPr, "w", "jc")
	if got := jc.SelectAttrValue("w:val", ""); got != "right" {
		t.Errorf("jc: expected right (direct wins), got %s", got)
	}
	// ind: only in style, should be added.
	ind := findChild(pPr, "w", "ind")
	if ind == nil {
		t.Fatal("ind should be added from style")
	}
	if got := ind.SelectAttrValue("w:left", ""); got != "720" {
		t.Errorf("ind: expected 720, got %s", got)
	}
}

func TestExpandDirectFormatting_RunStyleExpansion(t *testing.T) {
	t.Parallel()
	// Character style: Strong with bold + red color.
	srcStyles := buildStylesXml(
		`<w:style w:type="character" w:styleId="Strong">` +
			`<w:name w:val="Strong"/>` +
			`<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>` +
			`</w:style>`,
	)
	strong := srcStyles.GetByID("Strong")

	xml := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPr><w:rStyle w:val="Strong"/></w:rPr>` +
		`<w:t>bold text</w:t>` +
		`</w:r>`
	el, _ := oxml.ParseXml([]byte(xml))
	rPr := findChild(el, "w", "rPr")

	_, resolvedRPr := resolveStyleChainFrom(srcStyles, strong)
	mergePropertiesDeep(rPr, resolvedRPr)

	if findChild(rPr, "w", "b") == nil {
		t.Error("b should be expanded from style")
	}
	color := findChild(rPr, "w", "color")
	if color == nil {
		t.Fatal("color should be expanded from style")
	}
	if got := color.SelectAttrValue("w:val", ""); got != "FF0000" {
		t.Errorf("color: expected FF0000, got %s", got)
	}
}

func TestExpandDirectFormatting_RunDirectFormattingPreserved(t *testing.T) {
	t.Parallel()
	srcStyles := buildStylesXml(
		`<w:style w:type="character" w:styleId="Strong">` +
			`<w:name w:val="Strong"/>` +
			`<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>` +
			`</w:style>`,
	)
	strong := srcStyles.GetByID("Strong")

	// Run already has sz=36 (direct), style has sz=24.
	xml := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPr><w:rStyle w:val="Strong"/><w:sz w:val="36"/></w:rPr>` +
		`<w:t>big bold</w:t>` +
		`</w:r>`
	el, _ := oxml.ParseXml([]byte(xml))
	rPr := findChild(el, "w", "rPr")

	_, resolvedRPr := resolveStyleChainFrom(srcStyles, strong)
	mergePropertiesDeep(rPr, resolvedRPr)

	// sz: direct=36 should win over style=24.
	sz := findChild(rPr, "w", "sz")
	if got := sz.SelectAttrValue("w:val", ""); got != "36" {
		t.Errorf("sz: expected 36 (direct wins), got %s", got)
	}
	// b should still be added from style.
	if findChild(rPr, "w", "b") == nil {
		t.Error("b should be expanded from style")
	}
}

func TestExpandDirectFormatting_InheritedStyleChain(t *testing.T) {
	t.Parallel()
	// Base has font, Derived overrides size and adds bold.
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Base">` +
			`<w:name w:val="Base"/>` +
			`<w:pPr><w:spacing w:after="200"/></w:pPr>` +
			`<w:rPr><w:rFonts w:ascii="Times"/><w:sz w:val="20"/></w:rPr>` +
			`</w:style>`,
		`<w:style w:type="paragraph" w:styleId="Derived">` +
			`<w:name w:val="Derived"/>` +
			`<w:basedOn w:val="Base"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>` +
			`</w:style>`,
	)
	derived := srcStyles.GetByID("Derived")

	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Derived"/></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))
	pPr := findChild(el, "w", "pPr")

	resolvedPPr, resolvedRPr := resolveStyleChainFrom(srcStyles, derived)
	if resolvedPPr != nil {
		mergePropertiesDeep(pPr, resolvedPPr)
	}
	if resolvedRPr != nil {
		rPr := etree.NewElement("w:rPr")
		pPr.AddChild(rPr)
		mergePropertiesDeep(rPr, resolvedRPr)
	}

	// spacing from Base should be inherited.
	spacing := findChild(pPr, "w", "spacing")
	if spacing == nil {
		t.Fatal("spacing should be inherited from Base")
	}
	if got := spacing.SelectAttrValue("w:after", ""); got != "200" {
		t.Errorf("spacing after: expected 200, got %s", got)
	}

	// jc from Derived.
	jc := findChild(pPr, "w", "jc")
	if jc == nil {
		t.Fatal("jc should come from Derived")
	}

	// rPr: sz=28 (Derived overrides Base's 20), b from Derived, rFonts from Base.
	rPr := findChild(pPr, "w", "rPr")
	if rPr == nil {
		t.Fatal("rPr should exist")
	}
	sz := findChild(rPr, "w", "sz")
	if got := sz.SelectAttrValue("w:val", ""); got != "28" {
		t.Errorf("sz: expected 28 (Derived overrides), got %s", got)
	}
	if findChild(rPr, "w", "b") == nil {
		t.Error("b should come from Derived")
	}
	rf := findChild(rPr, "w", "rFonts")
	if rf == nil {
		t.Fatal("rFonts should be inherited from Base")
	}
	if got := rf.SelectAttrValue("w:ascii", ""); got != "Times" {
		t.Errorf("rFonts ascii: expected Times, got %s", got)
	}
}

func TestExpandDirectFormatting_StyleNotInExpandMap(t *testing.T) {
	t.Parallel()
	// Paragraph with a style NOT in expandStyles — expandParagraphStyle
	// checks the map and skips, no sourceDoc access needed.
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Normal"/></w:pPr>` +
		`<w:r><w:t>text</w:t></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1">` +
			`<w:name w:val="heading 1"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
	)

	// expandDirectFormatting with "Heading1" in map but paragraph uses "Normal"
	// → no match → expandParagraphStyle returns before calling resolveStyleChain.
	ri := newExpandTestRI(map[string]*oxml.CT_Style{
		"Heading1": srcStyles.GetByID("Heading1"),
	})
	// sourceDoc is nil, but expandParagraphStyle will bail out before
	// calling resolveStyleChain because "Normal" is not in expandStyles.
	ri.expandDirectFormatting([]*etree.Element{el})

	// Normal is not in expandStyles, so nothing should change.
	pPr := findChild(el, "w", "pPr")
	if len(pPr.ChildElements()) != 1 {
		t.Errorf("pPr should still have only pStyle, got %d children", len(pPr.ChildElements()))
	}
}

func TestExpandDirectFormatting_NoPPr(t *testing.T) {
	t.Parallel()
	// Paragraph with no pPr — expandParagraphStyle returns immediately
	// (no sourceDoc access needed).
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:t>text</w:t></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1">` +
			`<w:name w:val="heading 1"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
	)

	ri := newExpandTestRI(map[string]*oxml.CT_Style{
		"Heading1": srcStyles.GetByID("Heading1"),
	})
	ri.expandDirectFormatting([]*etree.Element{el})

	// No pPr → nothing should be created.
	if findChild(el, "w", "pPr") != nil {
		t.Error("pPr should not be created when paragraph has no pPr")
	}
}

func TestExpandDirectFormatting_MultipleElements(t *testing.T) {
	t.Parallel()
	// Body with two paragraphs, one with expand style, one without.
	// Test uses resolveStyleChainFrom + mergePropertiesDeep to simulate
	// what expandDirectFormatting does internally (avoids needing a full
	// Document for sourceStyles access).
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1">` +
			`<w:name w:val="heading 1"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
	)
	heading1 := srcStyles.GetByID("Heading1")

	p1Xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>`
	p2Xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>`
	p1, _ := oxml.ParseXml([]byte(p1Xml))
	p2, _ := oxml.ParseXml([]byte(p2Xml))

	// Apply expand to Heading1 paragraph only.
	expandMap := map[string]*oxml.CT_Style{"Heading1": heading1}
	pPr1 := findChild(p1, "w", "pPr")
	pStyleEl := findChild(pPr1, "w", "pStyle")
	sid := pStyleEl.SelectAttrValue("w:val", "")
	if _, ok := expandMap[sid]; ok {
		resolvedPPr, _ := resolveStyleChainFrom(srcStyles, heading1)
		if resolvedPPr != nil {
			mergePropertiesDeep(pPr1, resolvedPPr)
		}
	}

	// Heading1 paragraph should have jc expanded.
	if findChild(pPr1, "w", "jc") == nil {
		t.Error("Heading1 paragraph should have jc expanded")
	}

	// Normal paragraph should NOT have jc.
	pPr2 := findChild(p2, "w", "pPr")
	if findChild(pPr2, "w", "jc") != nil {
		t.Error("Normal paragraph should not have jc added")
	}
}

func TestExpandDirectFormatting_NestedInTable(t *testing.T) {
	t.Parallel()
	// Paragraph inside a table cell — expansion should handle nested structures.
	srcStyles := buildStylesXml(
		`<w:style w:type="paragraph" w:styleId="Heading1">` +
			`<w:name w:val="heading 1"/>` +
			`<w:pPr><w:jc w:val="center"/></w:pPr>` +
			`</w:style>`,
	)
	heading1 := srcStyles.GetByID("Heading1")

	xml := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:tr><w:tc>` +
		`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Cell</w:t></w:r></w:p>` +
		`</w:tc></w:tr>` +
		`</w:tbl>`
	el, _ := oxml.ParseXml([]byte(xml))

	// Find the paragraph inside the table and apply expansion manually.
	p := el.FindElement("//p")
	if p == nil {
		t.Fatal("paragraph should exist in table")
	}
	pPr := findChild(p, "w", "pPr")
	pStyleEl := findChild(pPr, "w", "pStyle")
	sid := pStyleEl.SelectAttrValue("w:val", "")
	if sid == "Heading1" {
		resolvedPPr, _ := resolveStyleChainFrom(srcStyles, heading1)
		if resolvedPPr != nil {
			mergePropertiesDeep(pPr, resolvedPPr)
		}
	}

	if findChild(pPr, "w", "jc") == nil {
		t.Error("paragraph in table should have jc expanded")
	}
}

// --------------------------------------------------------------------------
// removeChild tests
// --------------------------------------------------------------------------

func TestRemoveChild_Exists(t *testing.T) {
	t.Parallel()
	xml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pStyle w:val="Normal"/><w:jc w:val="center"/>` +
		`</w:pPr>`
	el, _ := oxml.ParseXml([]byte(xml))

	removeChild(el, "w", "pStyle")
	if findChild(el, "w", "pStyle") != nil {
		t.Error("pStyle should have been removed")
	}
	if findChild(el, "w", "jc") == nil {
		t.Error("jc should still be present")
	}
}

func TestRemoveChild_DoesNotExist(t *testing.T) {
	t.Parallel()
	xml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>`
	el, _ := oxml.ParseXml([]byte(xml))

	removeChild(el, "w", "pStyle") // should not panic
	if findChild(el, "w", "jc") == nil {
		t.Error("jc should still be present")
	}
}

// --------------------------------------------------------------------------
// copyStyleToTarget — semiHidden/unhideWhenUsed tests (Step 5)
// --------------------------------------------------------------------------

func TestCopyStyleToTarget_RenamedStyle_AddsSemiHidden(t *testing.T) {
	t.Parallel()
	// When a style is copied under a new ID (ForceCopyStyles rename),
	// semiHidden and unhideWhenUsed must be added to prevent clutter
	// in Word's Style Gallery.
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	srcStyles, _ := source.part.Styles()
	srcStyleXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="Heading1">` +
		`<w:name w:val="heading 1"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcStyleEl, _ := oxml.ParseXml([]byte(srcStyleXml))
	srcStyles.RawElement().AddChild(srcStyleEl)
	srcStyle := srcStyles.GetByID("Heading1")
	if srcStyle == nil {
		t.Fatal("source style Heading1 not found")
	}

	ri := newResourceImporter(source, target, target.wmlPkg, KeepSourceFormatting,
		ImportFormatOptions{ForceCopyStyles: true})

	if err := ri.copyStyleToTarget(srcStyle, "Heading1_0"); err != nil {
		t.Fatalf("copyStyleToTarget: %v", err)
	}

	// Verify the clone in target has semiHidden and unhideWhenUsed.
	tgtStyles, _ := target.part.Styles()
	copied := tgtStyles.GetByID("Heading1_0")
	if copied == nil {
		t.Fatal("copied style Heading1_0 not found in target")
	}
	raw := copied.RawElement()
	if findChild(raw, "w", "semiHidden") == nil {
		t.Error("expected semiHidden on renamed style")
	}
	if findChild(raw, "w", "unhideWhenUsed") == nil {
		t.Error("expected unhideWhenUsed on renamed style")
	}
}

func TestCopyStyleToTarget_SameId_NoSemiHidden(t *testing.T) {
	t.Parallel()
	// When a style is copied under the SAME ID (new style, not a rename),
	// semiHidden/unhideWhenUsed must NOT be added.
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	srcStyles, _ := source.part.Styles()
	srcStyleXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="MyCustom">` +
		`<w:name w:val="My Custom"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcStyleEl, _ := oxml.ParseXml([]byte(srcStyleXml))
	srcStyles.RawElement().AddChild(srcStyleEl)
	srcStyle := srcStyles.GetByID("MyCustom")

	ri := newResourceImporter(source, target, target.wmlPkg, UseDestinationStyles,
		ImportFormatOptions{})

	if err := ri.copyStyleToTarget(srcStyle, "MyCustom"); err != nil {
		t.Fatalf("copyStyleToTarget: %v", err)
	}

	tgtStyles, _ := target.part.Styles()
	copied := tgtStyles.GetByID("MyCustom")
	if copied == nil {
		t.Fatal("copied style MyCustom not found in target")
	}
	raw := copied.RawElement()
	if findChild(raw, "w", "semiHidden") != nil {
		t.Error("semiHidden should NOT be present for same-ID copy")
	}
	if findChild(raw, "w", "unhideWhenUsed") != nil {
		t.Error("unhideWhenUsed should NOT be present for same-ID copy")
	}
}

func TestCopyStyleToTarget_RenamedStyle_PreservesExistingSemiHidden(t *testing.T) {
	t.Parallel()
	// If the source style already has semiHidden, copyStyleToTarget must not
	// add a duplicate element.
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	srcStyles, _ := source.part.Styles()
	srcStyleXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="HiddenStyle">` +
		`<w:name w:val="Hidden Style"/>` +
		`<w:semiHidden/>` +
		`<w:unhideWhenUsed/>` +
		`</w:style>`
	srcStyleEl, _ := oxml.ParseXml([]byte(srcStyleXml))
	srcStyles.RawElement().AddChild(srcStyleEl)
	srcStyle := srcStyles.GetByID("HiddenStyle")

	ri := newResourceImporter(source, target, target.wmlPkg, KeepSourceFormatting,
		ImportFormatOptions{ForceCopyStyles: true})

	if err := ri.copyStyleToTarget(srcStyle, "HiddenStyle_0"); err != nil {
		t.Fatalf("copyStyleToTarget: %v", err)
	}

	tgtStyles, _ := target.part.Styles()
	copied := tgtStyles.GetByID("HiddenStyle_0")
	if copied == nil {
		t.Fatal("copied style not found")
	}
	raw := copied.RawElement()

	// Count semiHidden elements — must be exactly 1 (no duplicate).
	count := 0
	for _, child := range raw.ChildElements() {
		if child.Space == "w" && child.Tag == "semiHidden" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("expected exactly 1 semiHidden, got %d", count)
	}
}

func TestCopyStyleToTarget_RenamedStyle_DisplayNameUpdated(t *testing.T) {
	t.Parallel()
	// Verify display name gets " (imported)" suffix.
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	srcStyles, _ := source.part.Styles()
	srcStyleXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="Title">` +
		`<w:name w:val="Title"/>` +
		`<w:pPr><w:jc w:val="center"/></w:pPr>` +
		`</w:style>`
	srcStyleEl, _ := oxml.ParseXml([]byte(srcStyleXml))
	srcStyles.RawElement().AddChild(srcStyleEl)
	srcStyle := srcStyles.GetByID("Title")

	ri := newResourceImporter(source, target, target.wmlPkg, KeepSourceFormatting,
		ImportFormatOptions{ForceCopyStyles: true})

	if err := ri.copyStyleToTarget(srcStyle, "Title_0"); err != nil {
		t.Fatalf("copyStyleToTarget: %v", err)
	}

	tgtStyles, _ := target.part.Styles()
	copied := tgtStyles.GetByID("Title_0")
	if copied == nil {
		t.Fatal("copied style not found")
	}
	nameEl := findChild(copied.RawElement(), "w", "name")
	if nameEl == nil {
		t.Fatal("name element not found")
	}
	got := nameEl.SelectAttrValue("w:val", "")
	if got != "Title (imported)" {
		t.Errorf("display name = %q, want %q", got, "Title (imported)")
	}
}

// --------------------------------------------------------------------------
// uniqueStyleId tests (Step 5)
// --------------------------------------------------------------------------

func TestUniqueStyleId_BasicSuffix(t *testing.T) {
	t.Parallel()
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	ri := newResourceImporter(source, target, target.wmlPkg, KeepSourceFormatting,
		ImportFormatOptions{ForceCopyStyles: true})

	got := ri.uniqueStyleId("Heading1")
	if got != "Heading1_0" {
		t.Errorf("uniqueStyleId = %q, want %q", got, "Heading1_0")
	}
}

func TestUniqueStyleId_SkipsExisting(t *testing.T) {
	t.Parallel()
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	// Pre-populate target with Heading1_0.
	tgtStyles, _ := target.part.Styles()
	existing := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="Heading1_0"><w:name w:val="h1_0"/></w:style>`
	el, _ := oxml.ParseXml([]byte(existing))
	tgtStyles.RawElement().AddChild(el)

	ri := newResourceImporter(source, target, target.wmlPkg, KeepSourceFormatting,
		ImportFormatOptions{ForceCopyStyles: true})

	got := ri.uniqueStyleId("Heading1")
	if got != "Heading1_1" {
		t.Errorf("uniqueStyleId = %q, want %q (should skip existing _0)", got, "Heading1_1")
	}
}

func TestUniqueStyleId_SkipsStyleMapCollision(t *testing.T) {
	t.Parallel()
	target := mustNewDoc(t)
	source := mustNewDoc(t)

	ri := newResourceImporter(source, target, target.wmlPkg, KeepSourceFormatting,
		ImportFormatOptions{ForceCopyStyles: true})
	// Simulate a previous import that mapped something to "Custom_0".
	ri.styleMap["Custom_0"] = "Custom_0"

	got := ri.uniqueStyleId("Custom")
	if got != "Custom_1" {
		t.Errorf("uniqueStyleId = %q, want %q (should skip styleMap collision)", got, "Custom_1")
	}
}

// --------------------------------------------------------------------------
// remapAll early exit fix — regression test
// --------------------------------------------------------------------------

func TestRemapAll_StyleOnlyNoNumIds(t *testing.T) {
	t.Parallel()
	// This tests the critical fix: remapAll must process styles even when
	// numIdMap is empty.
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:pStyle w:val="Custom"/></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},          // empty!
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"Custom": "Renamed"},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	pStyle := el.FindElement("//pStyle")
	got := pStyle.SelectAttrValue("w:val", "")
	if got != "Renamed" {
		t.Errorf("expected pStyle remapped to Renamed (even with empty numIdMap), got %s", got)
	}
}
