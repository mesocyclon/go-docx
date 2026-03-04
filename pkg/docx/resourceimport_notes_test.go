package docx

import (
	"strconv"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// collectNoteRefsFromElements tests
// --------------------------------------------------------------------------

func TestCollectNoteRefsFromElements_Footnotes(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
		`<w:footnoteReference w:id="1"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNoteRefsFromElements([]*etree.Element{el}, "footnoteReference")
	if len(ids) != 1 || ids[0] != 1 {
		t.Errorf("expected [1], got %v", ids)
	}
}

func TestCollectNoteRefsFromElements_Endnotes(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:endnoteReference w:id="3"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNoteRefsFromElements([]*etree.Element{el}, "endnoteReference")
	if len(ids) != 1 || ids[0] != 3 {
		t.Errorf("expected [3], got %v", ids)
	}
}

func TestCollectNoteRefsFromElements_Multiple(t *testing.T) {
	t.Parallel()
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
		`<w:p><w:r><w:footnoteReference w:id="5"/></w:r></w:p>` +
		`<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNoteRefsFromElements(el.ChildElements(), "footnoteReference")
	if len(ids) != 3 {
		t.Fatalf("expected 3 ids, got %d: %v", len(ids), ids)
	}
	// Document order.
	expected := []int{1, 5, 2}
	for i, e := range expected {
		if ids[i] != e {
			t.Errorf("ids[%d]: expected %d, got %d", i, e, ids[i])
		}
	}
}

func TestCollectNoteRefsFromElements_Dedup(t *testing.T) {
	t.Parallel()
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
		`<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNoteRefsFromElements(el.ChildElements(), "footnoteReference")
	if len(ids) != 1 {
		t.Errorf("expected 1 unique id, got %d", len(ids))
	}
}

func TestCollectNoteRefsFromElements_SkipZeroAndNegative(t *testing.T) {
	t.Parallel()
	// id=-1 (separator) and id=0 (continuationSeparator) must be skipped.
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:footnoteReference w:id="-1"/></w:r></w:p>` +
		`<w:p><w:r><w:footnoteReference w:id="0"/></w:r></w:p>` +
		`<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNoteRefsFromElements(el.ChildElements(), "footnoteReference")
	if len(ids) != 1 || ids[0] != 1 {
		t.Errorf("expected [1] (skip -1 and 0), got %v", ids)
	}
}

func TestCollectNoteRefsFromElements_Empty(t *testing.T) {
	t.Parallel()
	ids := collectNoteRefsFromElements(nil, "footnoteReference")
	if len(ids) != 0 {
		t.Errorf("expected empty, got %v", ids)
	}
}

func TestCollectNoteRefsFromElements_WrongTag(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:footnoteReference w:id="1"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	// Looking for endnotes but only footnotes present.
	ids := collectNoteRefsFromElements([]*etree.Element{el}, "endnoteReference")
	if len(ids) != 0 {
		t.Errorf("expected empty for wrong refTag, got %v", ids)
	}
}

// --------------------------------------------------------------------------
// findNoteById tests
// --------------------------------------------------------------------------

func buildNotesXml(tag string, ids ...int) *etree.Element {
	xml := `<w:` + tag + `s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
	for _, id := range ids {
		xml += `<w:` + tag + ` w:id="` + itoa(id) + `"><w:p><w:r><w:t>Note ` + itoa(id) + `</w:t></w:r></w:p></w:` + tag + `>`
	}
	xml += `</w:` + tag + `s>`
	el, _ := oxml.ParseXml([]byte(xml))
	return el
}

func itoa(i int) string {
	return strconv.Itoa(i)
}

func TestFindNoteById_Found(t *testing.T) {
	t.Parallel()
	el := buildNotesXml("footnote", -1, 0, 1, 3)

	note := findNoteById(el, "footnote", 3)
	if note == nil {
		t.Fatal("expected to find footnote id=3")
	}
	if note.SelectAttrValue("w:id", "") != "3" {
		t.Errorf("expected id=3, got %s", note.SelectAttrValue("w:id", ""))
	}
}

func TestFindNoteById_NotFound(t *testing.T) {
	t.Parallel()
	el := buildNotesXml("footnote", -1, 0, 1)

	if findNoteById(el, "footnote", 99) != nil {
		t.Error("expected nil for nonexistent id")
	}
}

func TestFindNoteById_Endnote(t *testing.T) {
	t.Parallel()
	el := buildNotesXml("endnote", -1, 0, 2)

	note := findNoteById(el, "endnote", 2)
	if note == nil {
		t.Fatal("expected to find endnote id=2")
	}
}

// --------------------------------------------------------------------------
// nextNoteId tests
// --------------------------------------------------------------------------

func TestNextNoteId_WithExisting(t *testing.T) {
	t.Parallel()
	el := buildNotesXml("footnote", -1, 0, 1, 3)

	got := nextNoteId(el, "footnote")
	if got != 4 {
		t.Errorf("expected 4 (max=3 + 1), got %d", got)
	}
}

func TestNextNoteId_SeparatorsOnly(t *testing.T) {
	t.Parallel()
	el := buildNotesXml("footnote", -1, 0)

	got := nextNoteId(el, "footnote")
	if got != 1 {
		t.Errorf("expected 1 (first user footnote), got %d", got)
	}
}

func TestNextNoteId_Empty(t *testing.T) {
	t.Parallel()
	xml := `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := oxml.ParseXml([]byte(xml))

	got := nextNoteId(el, "footnote")
	if got != 1 {
		t.Errorf("expected 1 on empty, got %d", got)
	}
}

// --------------------------------------------------------------------------
// setNoteId tests
// --------------------------------------------------------------------------

func TestSetNoteId(t *testing.T) {
	t.Parallel()
	xml := `<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1"/>`
	el, _ := oxml.ParseXml([]byte(xml))

	setNoteId(el, 42)
	if el.SelectAttrValue("w:id", "") != "42" {
		t.Errorf("expected id=42, got %s", el.SelectAttrValue("w:id", ""))
	}
}

// --------------------------------------------------------------------------
// noteBodyElements tests
// --------------------------------------------------------------------------

func TestNoteBodyElements(t *testing.T) {
	t.Parallel()
	xml := `<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1">` +
		`<w:p><w:r><w:t>Text</w:t></w:r></w:p>` +
		`<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
		`</w:footnote>`
	el, _ := oxml.ParseXml([]byte(xml))

	body := noteBodyElements(el)
	if len(body) != 2 {
		t.Fatalf("expected 2 body elements, got %d", len(body))
	}
	if body[0].Tag != "p" {
		t.Errorf("expected first element to be p, got %s", body[0].Tag)
	}
	if body[1].Tag != "tbl" {
		t.Errorf("expected second element to be tbl, got %s", body[1].Tag)
	}
}

// --------------------------------------------------------------------------
// remapAll footnote/endnote reference tests
// --------------------------------------------------------------------------

func TestRemapAll_FootnoteReference(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:footnoteReference w:id="1"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{},
		footnoteIdMap: map[int]int{1: 5},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	ref := el.FindElement("//footnoteReference")
	if ref == nil {
		t.Fatal("footnoteReference not found")
	}
	if got := ref.SelectAttrValue("w:id", ""); got != "5" {
		t.Errorf("expected id remapped to 5, got %s", got)
	}
}

func TestRemapAll_EndnoteReference(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:endnoteReference w:id="3"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{},
		footnoteIdMap: map[int]int{},
		endnoteIdMap:  map[int]int{3: 10},
	}

	ri.remapAll([]*etree.Element{el})

	ref := el.FindElement("//endnoteReference")
	if got := ref.SelectAttrValue("w:id", ""); got != "10" {
		t.Errorf("expected id remapped to 10, got %s", got)
	}
}

func TestRemapAll_FootnoteRefNotInMap(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:r><w:footnoteReference w:id="7"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{},
		footnoteIdMap: map[int]int{1: 5},
		endnoteIdMap:  map[int]int{},
	}

	ri.remapAll([]*etree.Element{el})

	ref := el.FindElement("//footnoteReference")
	if got := ref.SelectAttrValue("w:id", ""); got != "7" {
		t.Errorf("expected id to stay 7, got %s", got)
	}
}

func TestRemapAll_CombinedAllResourceTypes(t *testing.T) {
	t.Parallel()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr>` +
		`<w:pStyle w:val="Custom"/>` +
		`<w:numPr><w:numId w:val="3"/></w:numPr>` +
		`</w:pPr>` +
		`<w:r><w:footnoteReference w:id="1"/></w:r>` +
		`<w:r><w:endnoteReference w:id="2"/></w:r>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:      map[int]int{3: 99},
		absNumIdMap:   map[int]int{},
		styleMap:      map[string]string{"Custom": "Renamed"},
		footnoteIdMap: map[int]int{1: 50},
		endnoteIdMap:  map[int]int{2: 60},
	}

	ri.remapAll([]*etree.Element{el})

	if got := el.FindElement("//pStyle").SelectAttrValue("w:val", ""); got != "Renamed" {
		t.Errorf("pStyle: expected Renamed, got %s", got)
	}
	if got := el.FindElement("//numId").SelectAttrValue("w:val", ""); got != "99" {
		t.Errorf("numId: expected 99, got %s", got)
	}
	if got := el.FindElement("//footnoteReference").SelectAttrValue("w:id", ""); got != "50" {
		t.Errorf("footnoteRef: expected 50, got %s", got)
	}
	if got := el.FindElement("//endnoteReference").SelectAttrValue("w:id", ""); got != "60" {
		t.Errorf("endnoteRef: expected 60, got %s", got)
	}
}

// --------------------------------------------------------------------------
// Regression: note body pipeline must sanitize annotations (bug #1)
// --------------------------------------------------------------------------

func TestNoteBodyPipeline_SanitizesAnnotations(t *testing.T) {
	t.Parallel()
	// Footnote body containing bookmarkStart/End — these carry source-scoped
	// w:id values and must be stripped, same as for body content.
	xml := `<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1">` +
		`<w:p>` +
		`<w:bookmarkStart w:id="0" w:name="test"/>` +
		`<w:r><w:t>footnote with bookmark</w:t></w:r>` +
		`<w:bookmarkEnd w:id="0"/>` +
		`</w:p>` +
		`</w:footnote>`
	el, _ := oxml.ParseXml([]byte(xml))
	bodyEls := noteBodyElements(el)

	sanitizeForInsertion(bodyEls)

	if findDescendant(el, "w", "bookmarkStart") != nil {
		t.Error("bookmarkStart should have been stripped from footnote body")
	}
	if findDescendant(el, "w", "bookmarkEnd") != nil {
		t.Error("bookmarkEnd should have been stripped from footnote body")
	}
	if el.FindElement(".//w:t") == nil {
		t.Error("run text should have been preserved")
	}
}

func TestNoteBodyPipeline_SanitizesCommentRefs(t *testing.T) {
	t.Parallel()
	xml := `<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1">` +
		`<w:p>` +
		`<w:commentRangeStart w:id="3"/>` +
		`<w:r><w:t>commented text</w:t></w:r>` +
		`<w:commentRangeEnd w:id="3"/>` +
		`<w:r><w:commentReference w:id="3"/></w:r>` +
		`</w:p>` +
		`</w:footnote>`
	el, _ := oxml.ParseXml([]byte(xml))
	bodyEls := noteBodyElements(el)

	sanitizeForInsertion(bodyEls)

	if findDescendant(el, "w", "commentRangeStart") != nil {
		t.Error("commentRangeStart should have been stripped")
	}
	if findDescendant(el, "w", "commentReference") != nil {
		t.Error("commentReference should have been stripped")
	}
}

// --------------------------------------------------------------------------
// Regression: note body pipeline must renumber drawing IDs (bug #2)
// --------------------------------------------------------------------------

func TestNoteBodyPipeline_RenumbersDrawingIDs(t *testing.T) {
	t.Parallel()
	// Footnote body containing an inline drawing with bare numeric id
	// attribute (wp:docPr). Without renumbering, duplicate ids corrupt
	// the target document.
	xml := `<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
		` w:id="1">` +
		`<w:p><w:r>` +
		`<w:drawing><wp:inline><wp:docPr id="5" name="Pic1"/></wp:inline></w:drawing>` +
		`</w:r></w:p>` +
		`</w:footnote>`
	el, _ := oxml.ParseXml([]byte(xml))
	bodyEls := noteBodyElements(el)

	counter := 100
	renumberDrawingIDs(bodyEls, func() int {
		counter++
		return counter
	})

	docPr := findDescendant(el, "wp", "docPr")
	if docPr == nil {
		t.Fatal("wp:docPr not found after renumbering")
	}
	got := docPr.SelectAttrValue("id", "")
	if got == "5" {
		t.Error("drawing id should have been renumbered, still 5")
	}
	if got != "101" {
		t.Errorf("expected renumbered id 101, got %s", got)
	}
}
