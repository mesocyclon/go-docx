package docx

import (
	"crypto/rand"
	"fmt"
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// resourceimport_num.go — Numbering import for ResourceImporter
//
// Scans the source document body for numId references, then copies the
// corresponding abstractNum and num definitions into the target document
// with fresh IDs. The numIdMap is populated for later use by remapAll.
// --------------------------------------------------------------------------

// importNumbering imports all numbering definitions referenced by the source
// document body into the target document. Creates new abstractNum and num
// entries with fresh IDs and populates ri.numIdMap and ri.absNumIdMap.
//
// Safe to call multiple times — only runs once (idempotent via numDone flag).
func (ri *ResourceImporter) importNumbering() error {
	if ri.numDone {
		return nil
	}
	ri.numDone = true

	// 1. Collect all numId values from source body.
	referencedNumIds := collectNumIdsFromBody(ri.sourceDoc)
	if len(referencedNumIds) == 0 {
		return nil
	}

	// 2. Get source numbering part.
	srcNP, err := ri.sourceDoc.part.NumberingPart()
	if err != nil {
		// No numbering part — numId references are orphaned, skip.
		return nil
	}
	srcNumbering, err := srcNP.Numbering()
	if err != nil {
		return nil
	}

	// 3. Ensure numbering part exists in target.
	tgtNP, err := ri.targetDoc.part.GetOrAddNumberingPart()
	if err != nil {
		return fmt.Errorf("docx: ensuring target numbering part: %w", err)
	}
	tgtNumbering, err := tgtNP.Numbering()
	if err != nil {
		return fmt.Errorf("docx: accessing target numbering element: %w", err)
	}

	// 4. For each referenced numId, import the num and its abstractNum.
	for _, srcNumId := range referencedNumIds {
		if _, done := ri.numIdMap[srcNumId]; done {
			continue
		}

		// 4a. Find <w:num> in source.
		srcNum := srcNumbering.NumHavingNumId(srcNumId)
		if srcNum == nil {
			continue
		}

		// 4b. Extract abstractNumId.
		srcAbsId, err := abstractNumIdOf(srcNum)
		if err != nil {
			continue
		}

		// 4c. Import abstractNum (if not already done).
		tgtAbsId, ok := ri.absNumIdMap[srcAbsId]
		if !ok {
			imported, err := ri.importAbstractNum(srcNumbering, tgtNumbering, srcAbsId)
			if err != nil {
				return err
			}
			tgtAbsId = imported
			ri.absNumIdMap[srcAbsId] = tgtAbsId
		}

		// 4d. Create new <w:num> in target.
		tgtNum, err := tgtNumbering.AddNumWithAbstractNumId(tgtAbsId)
		if err != nil {
			return fmt.Errorf("docx: adding num to target: %w", err)
		}
		tgtNumId, err := tgtNum.NumId()
		if err != nil {
			return fmt.Errorf("docx: reading new numId: %w", err)
		}

		// 4e. Copy <w:lvlOverride> elements if present.
		copyLvlOverrides(srcNum, tgtNum)

		ri.numIdMap[srcNumId] = tgtNumId
	}
	return nil
}

// importAbstractNum copies a single <w:abstractNum> from source to target
// with a fresh abstractNumId and regenerated nsid.
func (ri *ResourceImporter) importAbstractNum(
	srcNumbering, tgtNumbering *oxml.CT_Numbering,
	srcAbsId int,
) (int, error) {
	srcAbsNum := srcNumbering.FindAbstractNum(srcAbsId)
	if srcAbsNum == nil {
		return 0, fmt.Errorf("docx: abstractNum %d not found in source", srcAbsId)
	}
	clone := srcAbsNum.Copy()

	// Assign new abstractNumId.
	tgtAbsId := tgtNumbering.NextAbstractNumId()
	setAbstractNumId(clone, tgtAbsId)

	// Regenerate nsid — Word does this on paste to prevent merging
	// definitions with the same nsid.
	regenerateNsid(clone)

	// Insert before first <w:num> in target.
	tgtNumbering.InsertAbstractNum(clone)

	return tgtAbsId, nil
}

// --------------------------------------------------------------------------
// Body scanning
// --------------------------------------------------------------------------

// collectNumIdsFromBody scans the source document body for all unique
// w:numId values referenced in w:pPr/w:numPr/w:numId/@w:val.
// Returns a deduplicated slice in document order.
func collectNumIdsFromBody(sourceDoc *Document) []int {
	srcBody := sourceDoc.element.Body()
	if srcBody == nil {
		return nil
	}
	return collectNumIdsFromElements(srcBody.RawElement().ChildElements())
}

// collectNumIdsFromElements scans element trees for w:numId values.
func collectNumIdsFromElements(elements []*etree.Element) []int {
	seen := map[int]bool{}
	var result []int
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			if el.Space == "w" && el.Tag == "numId" {
				if v := el.SelectAttrValue("w:val", ""); v != "" {
					if numId, err := strconv.Atoi(v); err == nil && numId != 0 && !seen[numId] {
						seen[numId] = true
						result = append(result, numId)
					}
				}
			}
			// Reverse push so LIFO stack yields document order.
			children := el.ChildElements()
			for i := len(children) - 1; i >= 0; i-- {
				stack = append(stack, children[i])
			}
		}
	}
	return result
}

// --------------------------------------------------------------------------
// Helpers
// --------------------------------------------------------------------------

// abstractNumIdOf extracts the abstractNumId value from a CT_Num.
func abstractNumIdOf(num *oxml.CT_Num) (int, error) {
	absNumIdEl, err := num.AbstractNumId()
	if err != nil {
		return 0, err
	}
	return absNumIdEl.Val()
}

// setAbstractNumId sets the w:abstractNumId attribute on an <w:abstractNum>
// raw etree element.
func setAbstractNumId(absNum *etree.Element, id int) {
	absNum.CreateAttr("w:abstractNumId", strconv.Itoa(id))
}

// regenerateNsid replaces the w:nsid value in an <w:abstractNum> element
// with a fresh random 8-digit hex string. Word uses nsid to detect
// "same definition" for merging; regenerating prevents unintended merges.
func regenerateNsid(absNum *etree.Element) {
	for _, child := range absNum.ChildElements() {
		if child.Space == "w" && child.Tag == "nsid" {
			child.CreateAttr("w:val", randomNsid())
			return
		}
	}
	// No nsid element — create one.
	nsid := absNum.CreateElement("nsid")
	nsid.Space = "w"
	nsid.CreateAttr("w:val", randomNsid())
}

// randomNsid generates an 8-character uppercase hex string for use as
// a numbering definition nsid. Uses crypto/rand for uniqueness.
func randomNsid() string {
	b := make([]byte, 4)
	_, _ = rand.Read(b) // infallible since Go 1.22
	return fmt.Sprintf("%08X", b)
}

// copyLvlOverrides copies all <w:lvlOverride> children from srcNum to tgtNum.
func copyLvlOverrides(srcNum, tgtNum *oxml.CT_Num) {
	for _, override := range srcNum.LvlOverrideList() {
		clone := override.RawElement().Copy()
		tgtNum.RawElement().AddChild(clone)
	}
}
