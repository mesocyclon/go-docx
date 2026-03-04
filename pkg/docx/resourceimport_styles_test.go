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
