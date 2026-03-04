package docx

import (
	"strconv"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// collectNumIdsFromElements tests
// --------------------------------------------------------------------------

func TestCollectNumIdsFromElements_Basic(t *testing.T) {
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>` +
		`<w:r><w:t>List item</w:t></w:r>` +
		`</w:p>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}

	ids := collectNumIdsFromElements([]*etree.Element{el})
	if len(ids) != 1 {
		t.Fatalf("expected 1 numId, got %d", len(ids))
	}
	if ids[0] != 5 {
		t.Errorf("expected numId=5, got %d", ids[0])
	}
}

func TestCollectNumIdsFromElements_Dedup(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:p>` +
		`<w:p><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:p>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNumIdsFromElements(el.ChildElements())
	if len(ids) != 1 {
		t.Fatalf("expected 1 unique numId, got %d", len(ids))
	}
}

func TestCollectNumIdsFromElements_SkipZero(t *testing.T) {
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNumIdsFromElements([]*etree.Element{el})
	if len(ids) != 0 {
		t.Errorf("expected 0 numIds (numId=0 is skip), got %d", len(ids))
	}
}

func TestCollectNumIdsFromElements_Multiple(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>` +
		`<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:p>` +
		`<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>` +
		`</w:body>`
	el, _ := oxml.ParseXml([]byte(xml))

	ids := collectNumIdsFromElements(el.ChildElements())
	if len(ids) != 2 {
		t.Fatalf("expected 2 unique numIds, got %d", len(ids))
	}
	if ids[0] != 1 || ids[1] != 7 {
		t.Errorf("expected [1, 7], got %v", ids)
	}
}

func TestCollectNumIdsFromElements_Empty(t *testing.T) {
	ids := collectNumIdsFromElements(nil)
	if len(ids) != 0 {
		t.Errorf("expected 0 numIds for nil elements, got %d", len(ids))
	}
}

// --------------------------------------------------------------------------
// remapAll numId tests
// --------------------------------------------------------------------------

func TestRemapAll_NumId(t *testing.T) {
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{5: 42},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{},
	}

	ri.remapAll([]*etree.Element{el})

	numIdEl := el.FindElement("//numId")
	if numIdEl == nil {
		t.Fatal("numId element not found after remap")
	}
	got := numIdEl.SelectAttrValue("w:val", "")
	if got != "42" {
		t.Errorf("expected numId remapped to 42, got %s", got)
	}
}

func TestRemapAll_NumId_NotInMap(t *testing.T) {
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{5: 42},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{},
	}

	ri.remapAll([]*etree.Element{el})

	numIdEl := el.FindElement("//numId")
	got := numIdEl.SelectAttrValue("w:val", "")
	if got != "99" {
		t.Errorf("expected numId to stay 99, got %s", got)
	}
}

func TestRemapAll_EmptyMap(t *testing.T) {
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr>` +
		`</w:p>`
	el, _ := oxml.ParseXml([]byte(xml))

	ri := &ResourceImporter{
		numIdMap:    map[int]int{},
		absNumIdMap: map[int]int{},
		styleMap:    map[string]string{},
	}

	ri.remapAll([]*etree.Element{el})

	numIdEl := el.FindElement("//numId")
	got := numIdEl.SelectAttrValue("w:val", "")
	if got != "5" {
		t.Errorf("expected numId to stay 5, got %s", got)
	}
}

// --------------------------------------------------------------------------
// Helper tests
// --------------------------------------------------------------------------

func TestRegenerateNsid(t *testing.T) {
	xml := `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">` +
		`<w:nsid w:val="12345678"/>` +
		`</w:abstractNum>`
	el, _ := oxml.ParseXml([]byte(xml))

	oldNsid := el.FindElement("//nsid").SelectAttrValue("w:val", "")
	regenerateNsid(el)
	newNsid := el.FindElement("//nsid").SelectAttrValue("w:val", "")

	if newNsid == oldNsid {
		t.Error("expected nsid to change after regeneration")
	}
	if len(newNsid) != 8 {
		t.Errorf("expected 8-char hex nsid, got %q", newNsid)
	}
}

func TestRegenerateNsid_NoExisting(t *testing.T) {
	xml := `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0"/>`
	el, _ := oxml.ParseXml([]byte(xml))

	regenerateNsid(el)

	var nsidEl *etree.Element
	for _, child := range el.ChildElements() {
		if child.Tag == "nsid" {
			nsidEl = child
			break
		}
	}
	if nsidEl == nil {
		t.Fatal("expected nsid element to be created")
	}
	val := nsidEl.SelectAttrValue("w:val", "")
	if len(val) != 8 {
		t.Errorf("expected 8-char hex nsid, got %q", val)
	}
}

func TestCopyLvlOverrides(t *testing.T) {
	srcXml := `<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="1">` +
		`<w:abstractNumId w:val="0"/>` +
		`<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>` +
		`<w:lvlOverride w:ilvl="1"><w:startOverride w:val="1"/></w:lvlOverride>` +
		`</w:num>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcNum := &oxml.CT_Num{Element: oxml.WrapElement(srcEl)}

	tgtXml := `<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="5">` +
		`<w:abstractNumId w:val="3"/>` +
		`</w:num>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtNum := &oxml.CT_Num{Element: oxml.WrapElement(tgtEl)}

	copyLvlOverrides(srcNum, tgtNum)

	overrides := tgtNum.LvlOverrideList()
	if len(overrides) != 2 {
		t.Fatalf("expected 2 lvlOverrides, got %d", len(overrides))
	}
	ilvl0, _ := overrides[0].Ilvl()
	ilvl1, _ := overrides[1].Ilvl()
	if ilvl0 != 0 || ilvl1 != 1 {
		t.Errorf("expected ilvl 0 and 1, got %d and %d", ilvl0, ilvl1)
	}
}

func TestSetAbstractNumId(t *testing.T) {
	el := etree.NewElement("abstractNum")
	el.Space = "w"
	el.CreateAttr("w:abstractNumId", "0")

	setAbstractNumId(el, 42)

	got := el.SelectAttrValue("w:abstractNumId", "")
	if got != "42" {
		t.Errorf("expected abstractNumId=42, got %s", got)
	}
}

func TestRandomNsid(t *testing.T) {
	nsid := randomNsid()
	if len(nsid) != 8 {
		t.Errorf("expected 8-char nsid, got %d chars: %q", len(nsid), nsid)
	}
	if _, err := strconv.ParseUint(nsid, 16, 32); err != nil {
		t.Errorf("expected valid hex, got %q: %v", nsid, err)
	}

	nsid2 := randomNsid()
	if nsid == nsid2 {
		t.Error("expected different nsids from two calls")
	}
}
