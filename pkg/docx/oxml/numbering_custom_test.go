package oxml

import (
	"testing"
)

func TestCT_Numbering_NextNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	if got := n.NextNumId(); got != 1 {
		t.Errorf("expected next numId=1 on empty, got %d", got)
	}
}

func TestCT_Numbering_AddNumWithAbstractNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	num, err := n.AddNumWithAbstractNumId(0)
	if err != nil {
		t.Fatalf("AddNumWithAbstractNumId: %v", err)
	}
	if num == nil {
		t.Fatal("expected num, got nil")
	}
	numId, err := num.NumId()
	if err != nil {
		t.Fatalf("numId error: %v", err)
	}
	if numId != 1 {
		t.Errorf("expected numId=1, got %d", numId)
	}

	// Check abstractNumId
	absNum, err := num.AbstractNumId()
	if err != nil {
		t.Fatalf("AbstractNumId error: %v", err)
	}
	absVal, err := absNum.Val()
	if err != nil {
		t.Fatalf("abstractNumId val error: %v", err)
	}
	if absVal != 0 {
		t.Errorf("expected abstractNumId=0, got %d", absVal)
	}

	// Add another
	num2, err := n.AddNumWithAbstractNumId(1)
	if err != nil {
		t.Fatalf("AddNumWithAbstractNumId: %v", err)
	}
	numId2, _ := num2.NumId()
	if numId2 != 2 {
		t.Errorf("expected numId=2, got %d", numId2)
	}
}

func TestCT_Numbering_NumHavingNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	num := n.NumHavingNumId(7)
	if num == nil {
		t.Fatal("expected num with numId=7, got nil")
	}

	if n.NumHavingNumId(999) != nil {
		t.Error("expected nil for nonexistent numId")
	}
}

func TestCT_Numbering_NextNumId_GapFilling(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	// Should find gap at 2
	if got := n.NextNumId(); got != 2 {
		t.Errorf("expected next numId=2 (gap), got %d", got)
	}
}

func TestNewNum(t *testing.T) {
	num, err := NewNum(5, 3)
	if err != nil {
		t.Fatalf("NewNum: %v", err)
	}
	numId, err := num.NumId()
	if err != nil {
		t.Fatalf("numId error: %v", err)
	}
	if numId != 5 {
		t.Errorf("expected numId=5, got %d", numId)
	}
	absNumId, err := num.AbstractNumId()
	if err != nil {
		t.Fatalf("AbstractNumId error: %v", err)
	}
	absVal, err := absNumId.Val()
	if err != nil {
		t.Fatalf("abstractNumId error: %v", err)
	}
	if absVal != 3 {
		t.Errorf("expected abstractNumId=3, got %d", absVal)
	}
}

func TestCT_NumPr_ValAccessors(t *testing.T) {
	xml := `<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:ilvl w:val="2"/>` +
		`<w:numId w:val="5"/>` +
		`</w:numPr>`
	el, _ := ParseXml([]byte(xml))
	np := &CT_NumPr{Element{e: el}}

	ilvl, err := np.IlvlVal()
	if err != nil {
		t.Fatalf("IlvlVal: %v", err)
	}
	if ilvl == nil || *ilvl != 2 {
		t.Errorf("expected ilvl=2, got %v", ilvl)
	}
	numId, err := np.NumIdVal()
	if err != nil {
		t.Fatalf("NumIdVal: %v", err)
	}
	if numId == nil || *numId != 5 {
		t.Errorf("expected numId=5, got %v", numId)
	}
}

func TestCT_NumPr_ValAccessors_Empty(t *testing.T) {
	xml := `<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := ParseXml([]byte(xml))
	np := &CT_NumPr{Element{e: el}}

	if iv, err := np.IlvlVal(); err != nil {
		t.Fatalf("IlvlVal: %v", err)
	} else if iv != nil {
		t.Error("expected nil ilvl on empty numPr")
	}
	if nid, err := np.NumIdVal(); err != nil {
		t.Fatalf("NumIdVal: %v", err)
	} else if nid != nil {
		t.Error("expected nil numId on empty numPr")
	}

	// Set and verify
	if err := np.SetIlvlVal(3); err != nil {
		t.Fatalf("SetIlvlVal: %v", err)
	}
	if err := np.SetNumIdVal(7); err != nil {
		t.Fatalf("SetNumIdVal: %v", err)
	}
	ilvl, err := np.IlvlVal()
	if err != nil {
		t.Fatalf("IlvlVal: %v", err)
	}
	if ilvl == nil || *ilvl != 3 {
		t.Errorf("expected ilvl=3, got %v", ilvl)
	}
	numId, err := np.NumIdVal()
	if err != nil {
		t.Fatalf("NumIdVal: %v", err)
	}
	if numId == nil || *numId != 7 {
		t.Errorf("expected numId=7, got %v", numId)
	}
}


func TestCT_Numbering_FindAbstractNum(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0"><w:nsid w:val="1A2B3C4D"/></w:abstractNum>` +
		`<w:abstractNum w:abstractNumId="3"><w:nsid w:val="AABBCCDD"/></w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	found := n.FindAbstractNum(3)
	if found == nil {
		t.Fatal("expected abstractNum with id=3, got nil")
	}
	if got := found.SelectAttrValue("w:abstractNumId", ""); got != "3" {
		t.Errorf("expected abstractNumId=3, got %s", got)
	}

	if n.FindAbstractNum(999) != nil {
		t.Error("expected nil for nonexistent abstractNumId")
	}
}

func TestCT_Numbering_NextAbstractNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0"/>` +
		`<w:abstractNum w:abstractNumId="2"/>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	if got := n.NextAbstractNumId(); got != 1 {
		t.Errorf("expected nextAbstractNumId=1 (gap), got %d", got)
	}
}

func TestCT_Numbering_NextAbstractNumId_Empty(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	if got := n.NextAbstractNumId(); got != 0 {
		t.Errorf("expected nextAbstractNumId=0 on empty, got %d", got)
	}
}

func TestCT_Numbering_InsertAbstractNum(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	absNum := OxmlElement("w:abstractNum")
	absNum.CreateAttr("w:abstractNumId", "5")
	n.InsertAbstractNum(absNum)

	children := el.ChildElements()
	if len(children) != 2 {
		t.Fatalf("expected 2 children, got %d", len(children))
	}
	if children[0].Tag != "abstractNum" {
		t.Errorf("expected first child to be abstractNum, got %s", children[0].Tag)
	}
	if children[1].Tag != "num" {
		t.Errorf("expected second child to be num, got %s", children[1].Tag)
	}
}

func TestCT_Numbering_AllAbstractNums(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:abstractNum w:abstractNumId="0"><w:nsid w:val="11111111"/></w:abstractNum>` +
		`<w:abstractNum w:abstractNumId="3"><w:nsid w:val="22222222"/></w:abstractNum>` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="2"><w:abstractNumId w:val="3"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	absNums := n.AllAbstractNums()
	if len(absNums) != 2 {
		t.Fatalf("expected 2 abstractNums, got %d", len(absNums))
	}
	if absNums[0].SelectAttrValue("w:abstractNumId", "") != "0" {
		t.Errorf("first abstractNum: expected id=0, got %s",
			absNums[0].SelectAttrValue("w:abstractNumId", ""))
	}
	if absNums[1].SelectAttrValue("w:abstractNumId", "") != "3" {
		t.Errorf("second abstractNum: expected id=3, got %s",
			absNums[1].SelectAttrValue("w:abstractNumId", ""))
	}
}

func TestAddLvlOverrideWithIlvl(t *testing.T) {
	t.Parallel()
	numEl := OxmlElement("w:num")
	numEl.CreateAttr("w:numId", "1")
	num := &CT_Num{Element{e: numEl}}

	lvl, err := num.AddLvlOverrideWithIlvl(3)
	if err != nil {
		t.Fatal(err)
	}
	ilvl, err := lvl.Ilvl()
	if err != nil {
		t.Fatalf("Ilvl: %v", err)
	}
	if ilvl != 3 {
		t.Errorf("ilvl = %d, want 3", ilvl)
	}
}

func TestAddStartOverrideWithVal(t *testing.T) {
	t.Parallel()
	numLvlEl := OxmlElement("w:lvlOverride")
	nl := &CT_NumLvl{Element{e: numLvlEl}}

	so, err := nl.AddStartOverrideWithVal(5)
	if err != nil {
		t.Fatal(err)
	}
	val, err := so.Val()
	if err != nil {
		t.Fatalf("Val: %v", err)
	}
	if val != 5 {
		t.Errorf("startOverride val = %d, want 5", val)
	}
}

func TestNewDecimalNumber_InvalidTag(t *testing.T) {
	t.Parallel()
	_, err := NewDecimalNumber("bad", 1)
	if err == nil {
		t.Error("expected error for invalid tag")
	}
}

func TestNewCtString_InvalidTag(t *testing.T) {
	t.Parallel()
	_, err := NewCtString("bad", "val")
	if err == nil {
		t.Error("expected error for invalid tag")
	}
}

func TestSetNumIdVal(t *testing.T) {
	t.Parallel()
	npEl := OxmlElement("w:numPr")
	np := &CT_NumPr{Element{e: npEl}}

	// Initially nil
	v, err := np.NumIdVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("expected nil")
	}

	if err := np.SetNumIdVal(5); err != nil {
		t.Fatal(err)
	}
	got, err := np.NumIdVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 5 {
		t.Errorf("NumIdVal() = %v, want 5", got)
	}
}

func TestSetIlvlVal(t *testing.T) {
	t.Parallel()
	npEl := OxmlElement("w:numPr")
	np := &CT_NumPr{Element{e: npEl}}

	v, err := np.IlvlVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("expected nil")
	}

	if err := np.SetIlvlVal(2); err != nil {
		t.Fatal(err)
	}
	got, err := np.IlvlVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 2 {
		t.Errorf("IlvlVal() = %v, want 2", got)
	}
}

func TestCT_Numbering_AllAbstractNums_Empty(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{e: el}}

	absNums := n.AllAbstractNums()
	if len(absNums) != 0 {
		t.Errorf("expected 0 abstractNums on numbering with only nums, got %d", len(absNums))
	}
}

func TestAbstractNumIdOf(t *testing.T) {
	tests := []struct {
		name string
		xml  string
		want int
	}{
		{
			name: "valid id",
			xml:  `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="5"/>`,
			want: 5,
		},
		{
			name: "zero id",
			xml:  `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="0"/>`,
			want: 0,
		},
		{
			name: "missing attr",
			xml:  `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			want: -1,
		},
		{
			name: "non-numeric",
			xml:  `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="abc"/>`,
			want: -1,
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			el, _ := ParseXml([]byte(tc.xml))
			got := AbstractNumIdOf(el)
			if got != tc.want {
				t.Errorf("AbstractNumIdOf() = %d, want %d", got, tc.want)
			}
		})
	}
}
