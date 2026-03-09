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

// --------------------------------------------------------------------------
// extractLevelSignatures tests
// --------------------------------------------------------------------------

func TestExtractLevelSignatures_Basic(t *testing.T) {
	t.Parallel()
	xml := `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
		`<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>` +
		`<w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>` +
		`</w:abstractNum>`
	el, _ := oxml.ParseXml([]byte(xml))

	sigs := extractLevelSignatures(el)
	if len(sigs) != 3 {
		t.Fatalf("expected 3 levels, got %d", len(sigs))
	}
	if s := sigs["0"]; s.numFmt != "decimal" || s.lvlText != "%1." {
		t.Errorf("level 0: got %+v", s)
	}
	if s := sigs["1"]; s.numFmt != "lowerLetter" || s.lvlText != "%2)" {
		t.Errorf("level 1: got %+v", s)
	}
	if s := sigs["2"]; s.numFmt != "lowerRoman" || s.lvlText != "%3." {
		t.Errorf("level 2: got %+v", s)
	}
}

func TestExtractLevelSignatures_Empty(t *testing.T) {
	t.Parallel()
	xml := `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">` +
		`<w:nsid w:val="AABBCCDD"/>` +
		`</w:abstractNum>`
	el, _ := oxml.ParseXml([]byte(xml))

	sigs := extractLevelSignatures(el)
	if len(sigs) != 0 {
		t.Errorf("expected empty map, got %d entries", len(sigs))
	}
}

func TestExtractLevelSignatures_MissingNumFmt(t *testing.T) {
	t.Parallel()
	xml := `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:lvlText w:val="%1."/></w:lvl>` +
		`</w:abstractNum>`
	el, _ := oxml.ParseXml([]byte(xml))

	sigs := extractLevelSignatures(el)
	if s := sigs["0"]; s.numFmt != "" {
		t.Errorf("expected empty numFmt, got %q", s.numFmt)
	}
}

// --------------------------------------------------------------------------
// abstractNumsCompatible tests
// --------------------------------------------------------------------------

func makeAbstractNum(t *testing.T, xml string) *etree.Element {
	t.Helper()
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	return el
}

func TestAbstractNumsCompatible_SingleLevel_Same(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`)
	if !abstractNumsCompatible(src, tgt) {
		t.Error("single level, same numFmt + lvlText → should be compatible")
	}
}

func TestAbstractNumsCompatible_SingleLevel_DifferentFmt(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>`)
	if abstractNumsCompatible(src, tgt) {
		t.Error("decimal vs bullet → should not be compatible")
	}
}

func TestAbstractNumsCompatible_SingleLevel_DifferentText(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum>`)
	if abstractNumsCompatible(src, tgt) {
		t.Error("same numFmt but different lvlText → should not be compatible")
	}
}

func TestAbstractNumsCompatible_MultiLevel_AllMatch(t *testing.T) {
	t.Parallel()
	lvls := `<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
		`<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>` +
		`<w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>`
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+lvls+`</w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+lvls+`</w:abstractNum>`)
	if !abstractNumsCompatible(src, tgt) {
		t.Error("3 levels, all match → should be compatible")
	}
}

func TestAbstractNumsCompatible_MultiLevel_Level2Differs(t *testing.T) {
	t.Parallel()
	srcLvls := `<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
		`<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>` +
		`<w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>`
	tgtLvls := `<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
		`<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>` +
		`<w:lvl w:ilvl="2"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%3."/></w:lvl>`
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+srcLvls+`</w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+tgtLvls+`</w:abstractNum>`)
	if abstractNumsCompatible(src, tgt) {
		t.Error("level 2 differs → should not be compatible")
	}
}

func TestAbstractNumsCompatible_DifferentLevelCount_OverlapOk(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>`+
		`<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>`+
		`<w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>`+
		`</w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>`+
		`</w:abstractNum>`)
	if !abstractNumsCompatible(src, tgt) {
		t.Error("overlap on level 0, matches → should be compatible")
	}
}

func TestAbstractNumsCompatible_NoLevels(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0"></w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1"></w:abstractNum>`)
	if abstractNumsCompatible(src, tgt) {
		t.Error("both empty → should not be compatible")
	}
}

func TestAbstractNumsCompatible_OneEmpty(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0"></w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`)
	if abstractNumsCompatible(src, tgt) {
		t.Error("one empty → should not be compatible (no overlap)")
	}
}

func TestAbstractNumsCompatible_NoOverlap(t *testing.T) {
	t.Parallel()
	src := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0">`+
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`)
	tgt := makeAbstractNum(t, `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="1">`+
		`<w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum>`)
	if abstractNumsCompatible(src, tgt) {
		t.Error("no overlapping ilvl → should not be compatible")
	}
}

// --------------------------------------------------------------------------
// findMatchingTargetNum tests
// --------------------------------------------------------------------------

// makeNumbering builds a CT_Numbering from XML string for tests.
func makeNumbering(t *testing.T, xml string) *oxml.CT_Numbering {
	t.Helper()
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	return &oxml.CT_Numbering{Element: oxml.WrapElement(el)}
}

func TestFindMatchingTargetNum_Match(t *testing.T) {
	srcXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	src := makeNumbering(t, srcXml)

	tgtXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	tgt := makeNumbering(t, tgtXml)

	ri := &ResourceImporter{
		numIdMap:    make(map[int]int),
		absNumIdMap: make(map[int]int),
		styleMap:    make(map[string]string),
	}

	got := ri.findMatchingTargetNum(0, src, tgt, buildAbsNumToNumIdMap(tgt))
	if got != 3 {
		t.Errorf("expected matching target numId=3, got %d", got)
	}
}

func TestFindMatchingTargetNum_NoMatch(t *testing.T) {
	srcXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	src := makeNumbering(t, srcXml)

	tgtXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	tgt := makeNumbering(t, tgtXml)

	ri := &ResourceImporter{
		numIdMap:    make(map[int]int),
		absNumIdMap: make(map[int]int),
		styleMap:    make(map[string]string),
	}

	got := ri.findMatchingTargetNum(0, src, tgt, buildAbsNumToNumIdMap(tgt))
	if got != 0 {
		t.Errorf("expected no match (0), got %d", got)
	}
}

func TestFindMatchingTargetNum_BulletMatch(t *testing.T) {
	srcXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="2">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="5"><w:abstractNumId w:val="2"/></w:num>` +
		`</w:numbering>`
	src := makeNumbering(t, srcXml)

	tgtXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:abstractNum w:abstractNumId="1">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>` +
		`</w:numbering>`
	tgt := makeNumbering(t, tgtXml)

	ri := &ResourceImporter{
		numIdMap:    make(map[int]int),
		absNumIdMap: make(map[int]int),
		styleMap:    make(map[string]string),
	}

	got := ri.findMatchingTargetNum(2, src, tgt, buildAbsNumToNumIdMap(tgt))
	if got != 2 {
		t.Errorf("expected matching bullet target numId=2, got %d", got)
	}
}

func TestFindMatchingTargetNum_EmptyTarget(t *testing.T) {
	srcXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	src := makeNumbering(t, srcXml)

	tgtXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	tgt := makeNumbering(t, tgtXml)

	ri := &ResourceImporter{
		numIdMap:    make(map[int]int),
		absNumIdMap: make(map[int]int),
		styleMap:    make(map[string]string),
	}

	got := ri.findMatchingTargetNum(0, src, tgt, buildAbsNumToNumIdMap(tgt))
	if got != 0 {
		t.Errorf("expected no match on empty target, got %d", got)
	}
}

func TestFindMatchingTargetNum_SrcAbsNumNotFound(t *testing.T) {
	srcXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	src := makeNumbering(t, srcXml)

	tgtXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	tgt := makeNumbering(t, tgtXml)

	ri := &ResourceImporter{
		numIdMap:    make(map[int]int),
		absNumIdMap: make(map[int]int),
		styleMap:    make(map[string]string),
	}

	// srcAbsId=99 doesn't exist in source
	got := ri.findMatchingTargetNum(99, src, tgt, buildAbsNumToNumIdMap(tgt))
	if got != 0 {
		t.Errorf("expected 0 for missing source abstractNum, got %d", got)
	}
}

// --------------------------------------------------------------------------
// buildAbsNumToNumIdMap tests
// --------------------------------------------------------------------------

func TestBuildAbsNumToNumIdMap(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0"/>` +
		`<w:abstractNum w:abstractNumId="5"/>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="2"><w:abstractNumId w:val="5"/></w:num>` +
		`<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	n := makeNumbering(t, xml)

	m := buildAbsNumToNumIdMap(n)
	if len(m) != 2 {
		t.Fatalf("expected 2 entries, got %d", len(m))
	}
	// absNumId=0 → first numId=1 (not 3)
	if m[0] != 1 {
		t.Errorf("absNumId=0: expected numId=1, got %d", m[0])
	}
	// absNumId=5 → numId=2
	if m[5] != 2 {
		t.Errorf("absNumId=5: expected numId=2, got %d", m[5])
	}
}

func TestBuildAbsNumToNumIdMap_Empty(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	n := makeNumbering(t, xml)

	m := buildAbsNumToNumIdMap(n)
	if len(m) != 0 {
		t.Errorf("expected empty map, got %d entries", len(m))
	}
}

// --------------------------------------------------------------------------
// findMatchingTargetNum cache path test
// --------------------------------------------------------------------------

func TestFindMatchingTargetNum_CacheHit(t *testing.T) {
	srcXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	src := makeNumbering(t, srcXml)

	tgtXml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0">` +
		`<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>` +
		`</w:abstractNum>` +
		`<w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	tgt := makeNumbering(t, tgtXml)

	ri := &ResourceImporter{
		numIdMap:    make(map[int]int),
		absNumIdMap: map[int]int{0: -7}, // negative sentinel = previously merged to numId=7
		styleMap:    make(map[string]string),
	}

	// Should hit cache, return 7 without scanning.
	got := ri.findMatchingTargetNum(0, src, tgt, buildAbsNumToNumIdMap(tgt))
	if got != 7 {
		t.Errorf("expected cached numId=7, got %d", got)
	}
}
