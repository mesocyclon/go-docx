package oxml

import (
	"fmt"
	"sort"

	"github.com/beevik/etree"
)

// ===========================================================================
// CT_Numbering — custom methods
// ===========================================================================

// AddNumWithAbstractNumId adds a new <w:num> referencing the given abstract
// numbering definition id. The new num is assigned the next available numId.
// Returns the newly created CT_Num.
func (n *CT_Numbering) AddNumWithAbstractNumId(abstractNumId int) (*CT_Num, error) {
	nextNumId := n.NextNumId()
	num, err := NewNum(nextNumId, abstractNumId)
	if err != nil {
		return nil, err
	}
	n.insertNum(num)
	return num, nil
}

// NumHavingNumId returns the <w:num> child with the given numId attribute,
// or nil if not found.
func (n *CT_Numbering) NumHavingNumId(numId int) *CT_Num {
	for _, num := range n.NumList() {
		id, err := num.NumId()
		if err == nil && id == numId {
			return num
		}
	}
	return nil
}

// NextNumId returns the first numId not used by any <w:num> element,
// starting at 1 and filling gaps.
func (n *CT_Numbering) NextNumId() int {
	var numIds []int
	for _, num := range n.NumList() {
		id, err := num.NumId()
		if err == nil {
			numIds = append(numIds, id)
		}
	}
	sort.Ints(numIds)
	idSet := make(map[int]bool, len(numIds))
	for _, id := range numIds {
		idSet[id] = true
	}
	for i := 1; i <= len(numIds)+1; i++ {
		if !idSet[i] {
			return i
		}
	}
	return len(numIds) + 1
}

// ===========================================================================
// CT_Num — custom methods
// ===========================================================================

// NewNum creates a new <w:num> element with the given numId and a child
// <w:abstractNumId> referencing abstractNumId.
func NewNum(numId, abstractNumId int) (*CT_Num, error) {
	el := OxmlElement("w:num")
	num := &CT_Num{Element{e: el}}
	if err := num.SetNumId(numId); err != nil {
		return nil, err
	}

	// Create <w:abstractNumId w:val="N"/>
	absEl := OxmlElement("w:abstractNumId")
	absNum := &CT_DecimalNumber{Element{e: absEl}}
	if err := absNum.SetVal(abstractNumId); err != nil {
		return nil, err
	}
	el.AddChild(absEl)

	return num, nil
}

// AddLvlOverrideWithIlvl adds a new <w:lvlOverride> child with the given ilvl attribute.
func (n *CT_Num) AddLvlOverrideWithIlvl(ilvl int) (*CT_NumLvl, error) {
	lvl := n.AddLvlOverride()
	if err := lvl.SetIlvl(ilvl); err != nil {
		return nil, err
	}
	return lvl, nil
}

// ===========================================================================
// CT_NumLvl — custom methods
// ===========================================================================

// AddStartOverrideWithVal adds a <w:startOverride> child element with the given val.
func (nl *CT_NumLvl) AddStartOverrideWithVal(val int) (*CT_DecimalNumber, error) {
	so := nl.GetOrAddStartOverride()
	if err := so.SetVal(val); err != nil {
		return nil, err
	}
	return so, nil
}

// ===========================================================================
// CT_NumPr — custom methods
// ===========================================================================

// NumIdVal returns the value of the w:numId/w:val attribute, or nil if not present.
func (np *CT_NumPr) NumIdVal() (*int, error) {
	numId := np.NumId()
	if numId == nil {
		return nil, nil
	}
	v, err := numId.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetNumIdVal sets the w:numId/w:val attribute, creating the element if needed.
func (np *CT_NumPr) SetNumIdVal(val int) error {
	if err := np.GetOrAddNumId().SetVal(val); err != nil {
		return err
	}
	return nil
}

// IlvlVal returns the value of the w:ilvl/w:val attribute, or nil if not present.
func (np *CT_NumPr) IlvlVal() (*int, error) {
	ilvl := np.Ilvl()
	if ilvl == nil {
		return nil, nil
	}
	v, err := ilvl.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetIlvlVal sets the w:ilvl/w:val attribute, creating the element if needed.
func (np *CT_NumPr) SetIlvlVal(val int) error {
	if err := np.GetOrAddIlvl().SetVal(val); err != nil {
		return err
	}
	return nil
}

// ===========================================================================
// CT_DecimalNumber — additional factory method
// ===========================================================================

// NewDecimalNumber creates a new element with the given namespace-prefixed tagname
// and val attribute set. Mirrors CT_DecimalNumber.new() from Python.
func NewDecimalNumber(nspTagname string, val int) (*CT_DecimalNumber, error) {
	el, err := TryOxmlElement(nspTagname)
	if err != nil {
		return nil, fmt.Errorf("NewDecimalNumber: %w", err)
	}
	dn := &CT_DecimalNumber{Element{e: el}}
	if err := dn.SetVal(val); err != nil {
		return nil, err
	}
	return dn, nil
}

// ===========================================================================
// CT_String — additional factory method
// ===========================================================================

// NewCtString creates a new element with the given namespace-prefixed tagname
// and val attribute set. Mirrors CT_String.new() from Python.
func NewCtString(nspTagname, val string) (*CT_String, error) {
	el, err := TryOxmlElement(nspTagname)
	if err != nil {
		return nil, fmt.Errorf("NewCtString: %w", err)
	}
	s := &CT_String{Element{e: el}}
	if err := s.SetVal(val); err != nil {
		return nil, err
	}
	return s, nil
}

// ===========================================================================
// CT_Numbering — raw etree helpers for abstractNum management
// ===========================================================================

// FindAbstractNum returns the <w:abstractNum> child element with the given
// w:abstractNumId attribute, or nil if not found.
func (n *CT_Numbering) FindAbstractNum(absNumId int) *etree.Element {
	target := fmt.Sprintf("%d", absNumId)
	for _, child := range n.e.ChildElements() {
		if child.Space == "w" && child.Tag == "abstractNum" {
			if child.SelectAttrValue("w:abstractNumId", "") == target {
				return child
			}
		}
	}
	return nil
}

// NextAbstractNumId returns the first abstractNumId not used by any
// <w:abstractNum> element, starting at 0 and filling gaps.
func (n *CT_Numbering) NextAbstractNumId() int {
	var ids []int
	for _, child := range n.e.ChildElements() {
		if child.Space == "w" && child.Tag == "abstractNum" {
			if v := child.SelectAttrValue("w:abstractNumId", ""); v != "" {
				if parsed, err := parseIntAttr(v); err == nil {
					ids = append(ids, parsed)
				}
			}
		}
	}
	sort.Ints(ids)
	idSet := make(map[int]bool, len(ids))
	for _, id := range ids {
		idSet[id] = true
	}
	for i := 0; i <= len(ids); i++ {
		if !idSet[i] {
			return i
		}
	}
	return len(ids)
}

// InsertAbstractNum inserts an <w:abstractNum> element before the first
// <w:num> element. If no <w:num> exists, the element is appended.
// This maintains the OOXML ordering requirement: all abstractNum elements
// must precede all num elements.
func (n *CT_Numbering) InsertAbstractNum(absNum *etree.Element) {
	// Find the first <w:num>.
	for _, child := range n.e.ChildElements() {
		if child.Space == "w" && child.Tag == "num" {
			insertBefore(n.e, absNum, child)
			return
		}
	}
	// No <w:num> found — append.
	n.e.AddChild(absNum)
}
