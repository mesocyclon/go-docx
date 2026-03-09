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

	// 4. Build reverse map for KeepSourceNumbering=false merge path.
	// Maps target abstractNumId → first numId that references it.
	// Built once, O(N), avoids repeated O(A*N) scans in findMatchingTargetNum.
	var tgtAbsToNum map[int]int
	if !ri.opts.KeepSourceNumbering {
		tgtAbsToNum = buildAbsNumToNumIdMap(tgtNumbering)
	}

	// 5. For each referenced numId, import the num and its abstractNum.
	for _, srcNumId := range referencedNumIds {
		if _, done := ri.numIdMap[srcNumId]; done {
			continue
		}

		// 5a. Find <w:num> in source.
		srcNum := srcNumbering.NumHavingNumId(srcNumId)
		if srcNum == nil {
			continue
		}

		// 5b. Extract abstractNumId.
		srcAbsId, err := abstractNumIdOf(srcNum)
		if err != nil {
			continue
		}

		// 5b2. When KeepSourceNumbering is disabled, try to merge into
		// a matching target list before creating a new abstractNum.
		// If a compatible list definition exists (matching numFmt + lvlText
		// across all overlapping levels), reuse it — the source numbering
		// continues the target sequence.
		//
		// Uses absNumIdMap as cache: if we already resolved this srcAbsId
		// to a target numId via merge, reuse the same mapping.
		if !ri.opts.KeepSourceNumbering {
			if tgtNumId := ri.findMatchingTargetNum(srcAbsId, srcNumbering, tgtNumbering, tgtAbsToNum); tgtNumId > 0 {
				ri.numIdMap[srcNumId] = tgtNumId
				ri.absNumIdMap[srcAbsId] = -tgtNumId // negative sentinel: merged, not copied
				continue
			}
		}

		// 5c. Import abstractNum (if not already done).
		tgtAbsId, ok := ri.absNumIdMap[srcAbsId]
		if !ok {
			imported, err := ri.importAbstractNum(srcNumbering, tgtNumbering, srcAbsId)
			if err != nil {
				return err
			}
			tgtAbsId = imported
			ri.absNumIdMap[srcAbsId] = tgtAbsId
		}

		// 5d. Create new <w:num> in target.
		tgtNum, err := tgtNumbering.AddNumWithAbstractNumId(tgtAbsId)
		if err != nil {
			return fmt.Errorf("docx: adding num to target: %w", err)
		}
		tgtNumId, err := tgtNum.NumId()
		if err != nil {
			return fmt.Errorf("docx: reading new numId: %w", err)
		}

		// 5e. Copy <w:lvlOverride> elements if present.
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

// --------------------------------------------------------------------------
// Numbering merge (KeepSourceNumbering = false)
// --------------------------------------------------------------------------

// buildAbsNumToNumIdMap builds a reverse index: target abstractNumId → first
// numId referencing it. Built once per importNumbering call to avoid repeated
// O(A*N) scans when matching multiple source lists.
func buildAbsNumToNumIdMap(numbering *oxml.CT_Numbering) map[int]int {
	result := make(map[int]int)
	for _, num := range numbering.NumList() {
		absId, err := abstractNumIdOf(num)
		if err != nil {
			continue
		}
		if _, exists := result[absId]; exists {
			continue // keep first — stable ordering
		}
		numId, err := num.NumId()
		if err != nil {
			continue
		}
		result[absId] = numId
	}
	return result
}

// findMatchingTargetNum finds a target num whose abstractNum is compatible
// with the source abstractNum for list merging.
//
// Matching heuristic: compare numFmt and lvlText of every overlapping level
// (by ilvl) between source and target abstractNums. If all overlapping levels
// match, the definitions are considered compatible and the target numId is
// returned. Levels present in only one definition are ignored.
//
// tgtAbsToNum is a pre-built reverse map (abstractNumId → numId) for O(1)
// lookup instead of scanning NumList() per candidate.
//
// Returns 0 if no matching target num exists — the caller should fall through
// to the standard abstractNum import path (creating a separate list).
func (ri *ResourceImporter) findMatchingTargetNum(
	srcAbsId int,
	srcNumbering, tgtNumbering *oxml.CT_Numbering,
	tgtAbsToNum map[int]int,
) int {
	// Cache check: if we already resolved this srcAbsId via merge path,
	// the absNumIdMap contains a negative sentinel (-tgtNumId).
	if cached, ok := ri.absNumIdMap[srcAbsId]; ok && cached < 0 {
		return -cached
	}

	srcAbsNum := srcNumbering.FindAbstractNum(srcAbsId)
	if srcAbsNum == nil {
		return 0
	}
	if !hasLevels(srcAbsNum) {
		return 0
	}

	// Scan target abstractNums for a compatible multi-level definition.
	for _, tgtAbsNum := range tgtNumbering.AllAbstractNums() {
		if !abstractNumsCompatible(srcAbsNum, tgtAbsNum) {
			continue
		}
		tgtAbsId := oxml.AbstractNumIdOf(tgtAbsNum)
		if tgtAbsId < 0 {
			continue
		}
		// O(1) lookup via pre-built reverse map.
		if numId, ok := tgtAbsToNum[tgtAbsId]; ok {
			return numId
		}
	}
	return 0
}

// levelSignature holds the identity-defining properties of a single
// numbering level for merge compatibility checks.
type levelSignature struct {
	numFmt  string // w:numFmt/@w:val (e.g. "decimal", "bullet")
	lvlText string // w:lvlText/@w:val (e.g. "%1.", "•")
}

// extractLevelSignatures returns a map from ilvl string ("0", "1", ...)
// to levelSignature for each <w:lvl> in an abstractNum element.
func extractLevelSignatures(absNum *etree.Element) map[string]levelSignature {
	sigs := map[string]levelSignature{}
	for _, child := range absNum.ChildElements() {
		if child.Space != "w" || child.Tag != "lvl" {
			continue
		}
		ilvl := child.SelectAttrValue("w:ilvl", "")
		if ilvl == "" {
			continue
		}
		var sig levelSignature
		for _, lc := range child.ChildElements() {
			if lc.Space != "w" {
				continue
			}
			switch lc.Tag {
			case "numFmt":
				sig.numFmt = lc.SelectAttrValue("w:val", "")
			case "lvlText":
				sig.lvlText = lc.SelectAttrValue("w:val", "")
			}
		}
		sigs[ilvl] = sig
	}
	return sigs
}

// hasLevels reports whether an abstractNum element contains at least one
// <w:lvl> child.
func hasLevels(absNum *etree.Element) bool {
	for _, child := range absNum.ChildElements() {
		if child.Space == "w" && child.Tag == "lvl" {
			return true
		}
	}
	return false
}

// abstractNumsCompatible reports whether two abstractNum definitions are
// semantically compatible for list merging. Two definitions are compatible
// when every level present in BOTH has identical numFmt and lvlText values.
//
// Levels present in only one definition are ignored — Word handles missing
// levels gracefully by falling back to the definition that has them.
//
// Returns false if either has no levels or if no levels overlap.
func abstractNumsCompatible(src, tgt *etree.Element) bool {
	srcSigs := extractLevelSignatures(src)
	tgtSigs := extractLevelSignatures(tgt)

	if len(srcSigs) == 0 || len(tgtSigs) == 0 {
		return false
	}

	overlap := 0
	for ilvl, srcSig := range srcSigs {
		tgtSig, ok := tgtSigs[ilvl]
		if !ok {
			continue
		}
		overlap++
		if srcSig.numFmt != tgtSig.numFmt || srcSig.lvlText != tgtSig.lvlText {
			return false
		}
	}
	return overlap > 0
}
