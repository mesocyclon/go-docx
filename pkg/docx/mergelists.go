package docx

import (
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// mergeAdjacentLists scans the container for adjacent list paragraphs
// with different numId values and merges them when they reference
// compatible abstractNum definitions in the target document's numbering.
//
// After importNumbering + remapAll, all numIds in both target-original
// and inserted elements reference the target numbering part. Two adjacent
// list paragraphs with compatible abstractNums (same numFmt/lvlText)
// but different numId are logically the same list type — merging makes
// Word render them as one continuous list.
//
// The function also recurses into table cells.
func mergeAdjacentLists(container *etree.Element, doc *Document) {
	mc := buildMergeContext(doc)
	if mc == nil {
		return
	}
	mergeAdjacentListsInContainer(container, mc)
}

// mergeContext holds pre-computed data for the merge walk.
type mergeContext struct {
	// numId → abstractNumId
	absMap map[int]int
	// abstractNumId → *etree.Element (the <w:abstractNum> element)
	absNumDefs map[int]*etree.Element
}

// compatible reports whether two numIds reference compatible abstractNum
// definitions (same numFmt + lvlText across overlapping levels).
func (mc *mergeContext) compatible(numIdA, numIdB int) bool {
	absA, okA := mc.absMap[numIdA]
	absB, okB := mc.absMap[numIdB]
	if !okA || !okB {
		return false
	}
	if absA == absB {
		return true // same abstractNum — always compatible
	}
	defA, okA := mc.absNumDefs[absA]
	defB, okB := mc.absNumDefs[absB]
	if !okA || !okB {
		return false
	}
	return abstractNumsCompatible(defA, defB)
}

// buildMergeContext creates the pre-computed merge context from the
// target document's numbering part. Returns nil if no numbering exists.
func buildMergeContext(doc *Document) *mergeContext {
	np, err := doc.part.NumberingPart()
	if err != nil {
		return nil // no numbering part — nothing to merge
	}
	numbering, err := np.Numbering()
	if err != nil {
		return nil
	}
	nums := numbering.NumList()
	if len(nums) == 0 {
		return nil
	}

	absMap := make(map[int]int, len(nums))
	for _, num := range nums {
		numId, err := num.NumId()
		if err != nil {
			continue
		}
		absId, err := num.AbstractNumId()
		if err != nil {
			continue
		}
		val, err := absId.Val()
		if err != nil {
			continue
		}
		absMap[numId] = val
	}

	absNumDefs := make(map[int]*etree.Element)
	for _, absNum := range numbering.AllAbstractNums() {
		id := oxml.AbstractNumIdOf(absNum)
		if id >= 0 {
			absNumDefs[id] = absNum
		}
	}

	return &mergeContext{
		absMap:     absMap,
		absNumDefs: absNumDefs,
	}
}

// mergeAdjacentListsInContainer performs the merge walk on a single
// container element (body, header, footer, table cell).
func mergeAdjacentListsInContainer(container *etree.Element, mc *mergeContext) {
	prevNumId := 0
	for _, child := range container.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			curNumId := extractNumId(child)
			if curNumId != 0 && prevNumId != 0 && curNumId != prevNumId {
				if mc.compatible(prevNumId, curNumId) {
					setNumId(child, prevNumId)
					curNumId = prevNumId
				}
			}
			prevNumId = curNumId
			continue
		}
		// Recurse into tables: w:tbl → w:tr → w:tc (each tc is a container).
		if child.Space == "w" && child.Tag == "tbl" {
			mergeAdjacentListsInTable(child, mc)
		}
		// Non-list, non-table element resets tracking.
		prevNumId = 0
	}
}

// mergeAdjacentListsInTable recurses into table rows and cells.
func mergeAdjacentListsInTable(tbl *etree.Element, mc *mergeContext) {
	for _, tr := range tbl.ChildElements() {
		if tr.Space != "w" || tr.Tag != "tr" {
			continue
		}
		for _, tc := range tr.ChildElements() {
			if tc.Space != "w" || tc.Tag != "tc" {
				continue
			}
			mergeAdjacentListsInContainer(tc, mc)
		}
	}
}

// extractNumId returns the w:numId value from a paragraph's
// w:pPr/w:numPr/w:numId, or 0 if the paragraph is not a list item.
func extractNumId(p *etree.Element) int {
	pPr := findChild(p, "w", "pPr")
	if pPr == nil {
		return 0
	}
	numPr := findChild(pPr, "w", "numPr")
	if numPr == nil {
		return 0
	}
	numIdEl := findChild(numPr, "w", "numId")
	if numIdEl == nil {
		return 0
	}
	v := numIdEl.SelectAttrValue("w:val", "")
	if v == "" {
		return 0
	}
	id, err := strconv.Atoi(v)
	if err != nil {
		return 0
	}
	return id
}

// setNumId overwrites the w:numId w:val attribute in a paragraph's numPr.
func setNumId(p *etree.Element, numId int) {
	pPr := findChild(p, "w", "pPr")
	if pPr == nil {
		return
	}
	numPr := findChild(pPr, "w", "numPr")
	if numPr == nil {
		return
	}
	numIdEl := findChild(numPr, "w", "numId")
	if numIdEl == nil {
		return
	}
	numIdEl.CreateAttr("w:val", strconv.Itoa(numId))
}
