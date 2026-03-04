package oxml

import (
	"github.com/beevik/etree"
)

// --------------------------------------------------------------------------
// replacetable.go — paragraph splitting engine for block-element replacement
//
// Provides the algorithm for splitting a single <w:p> at tag boundaries,
// producing a sequence of Fragments (alternating paragraph segments and
// placeholders). The caller (docx layer) replaces placeholders with block
// elements (tables, content, etc.) and splices the result into the container.
//
// This file lives in oxml because it operates purely on etree elements.
// It knows nothing about containers, recursion, table widths, or what
// replaces the placeholders.
//
// Reuses collectTextAtoms and findOccurrences from replacetext.go.
// --------------------------------------------------------------------------

// FragmentKind classifies a fragment produced by SplitParagraphAtTags.
type FragmentKind int

const (
	// FragmentParagraph is a text segment represented as a <w:p> element.
	FragmentParagraph FragmentKind = iota
	// FragmentPlaceholder marks the position where a block element should be inserted.
	FragmentPlaceholder
)

// Fragment is one piece of a paragraph split at tag boundaries.
// A slice of Fragments alternates between paragraph segments and placeholders.
//
// Mirrors the InnerContentItem pattern (blkcntnr.go): exported struct,
// private fields, accessor methods.
type Fragment struct {
	kind FragmentKind
	el   *etree.Element
}

// Kind returns the fragment classification.
func (f Fragment) Kind() FragmentKind { return f.kind }

// IsParagraph reports whether this fragment is a text paragraph segment.
func (f Fragment) IsParagraph() bool { return f.kind == FragmentParagraph }

// IsPlaceholder reports whether this fragment marks an insertion point.
func (f Fragment) IsPlaceholder() bool { return f.kind == FragmentPlaceholder }

// Element returns the <w:p> element for paragraph fragments, or nil for placeholders.
func (f Fragment) Element() *etree.Element { return f.el }

// --- Private constructors (used only by SplitParagraphAtTags) ---

func paragraphFragment(el *etree.Element) Fragment {
	return Fragment{kind: FragmentParagraph, el: el}
}

func placeholderFragment() Fragment {
	return Fragment{kind: FragmentPlaceholder}
}

// --- Segment range ---

// segRange represents a byte range [start, end) in the concatenated paragraph text
// that belongs to a text segment (i.e. NOT a tag).
type segRange struct {
	start, end int
}

// buildSegRanges computes the text segment ranges from tag positions.
// For N tags there are N+1 segments (some may be empty).
func buildSegRanges(positions []int, oldLen, textLen int) []segRange {
	n := len(positions)
	segs := make([]segRange, n+1)
	segs[0] = segRange{start: 0, end: positions[0]}
	for i := 1; i < n; i++ {
		segs[i] = segRange{start: positions[i-1] + oldLen, end: positions[i]}
	}
	segs[n] = segRange{start: positions[n-1] + oldLen, end: textLen}
	return segs
}

// --- Atom-to-segment mapping ---

// atomSegInfo describes which portion of an atom's text belongs to a segment.
type atomSegInfo struct {
	segIndex  int
	textStart int // byte offset within atom.text
	textEnd   int // byte offset within atom.text
}

// atomSegContributions computes which segments an atom contributes to.
func atomSegContributions(a textAtom, segs []segRange) []atomSegInfo {
	aStart := a.startPos
	aEnd := aStart + len(a.text)
	var result []atomSegInfo
	for i, seg := range segs {
		oStart := max(aStart, seg.start)
		oEnd := min(aEnd, seg.end)
		if oStart < oEnd {
			result = append(result, atomSegInfo{
				segIndex:  i,
				textStart: oStart - aStart,
				textEnd:   oEnd - aStart,
			})
		}
	}
	return result
}

// --- SplitParagraphAtTags ---

// SplitParagraphAtTags splits a paragraph at all occurrences of the tag old.
// Returns a slice of Fragments alternating between paragraph segments and
// placeholders, or nil if the tag is not found.
//
// Empty text segments do not produce fragments, with one exception: if the
// original paragraph contains a sectPr, the last fragment is guaranteed to
// be a paragraph (possibly empty, carrying only the sectPr).
//
// Example for "Before [<T1>] between [<T2>] after":
//
//	[paragraph("Before "), placeholder, paragraph(" between "), placeholder, paragraph(" after")]
func SplitParagraphAtTags(pEl *etree.Element, old string) []Fragment {
	if old == "" {
		return nil
	}

	atoms, fullText := collectTextAtoms(pEl)
	positions := findOccurrences(fullText, old)
	if len(positions) == 0 {
		return nil
	}

	oldLen := len(old)
	segs := buildSegRanges(positions, oldLen, len(fullText))

	// Build atom contributions map: atom element pointer → contributions.
	atomContribs := make(map[*etree.Element][]atomSegInfo)
	for _, a := range atoms {
		cs := atomSegContributions(a, segs)
		if len(cs) > 0 {
			atomContribs[a.elem] = cs
		}
	}

	// Set of all atom elements — used to distinguish "atom entirely in tag"
	// from "non-atom child" in run splitting.
	allAtomElems := make(map[*etree.Element]bool, len(atoms))
	for _, a := range atoms {
		allAtomElems[a.elem] = true
	}

	// Group atoms by parent run for run-level processing.
	atomsByRun := make(map[*etree.Element][]textAtom)
	for _, a := range atoms {
		atomsByRun[a.run] = append(atomsByRun[a.run], a)
	}

	// Extract pPr and detect sectPr.
	pPrEl, sectPrEl := extractPPrAndSectPr(pEl)

	// Build per-segment children lists.
	numSegs := len(segs)
	segChildren := make([][]*etree.Element, numSegs)
	lastSeg := 0 // default segment for non-textual children before any run

	for _, child := range pEl.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			continue // handled separately
		}

		if child.Space == "w" && child.Tag == "r" {
			if len(atomsByRun[child]) == 0 {
				// Run has no text atoms (e.g. only <w:drawing/>).
				// Treat as a non-textual paragraph child: clone to lastSeg (§3 case 13).
				segChildren[lastSeg] = append(segChildren[lastSeg], child.Copy())
			} else {
				distributeRunToSegments(child, atomContribs, allAtomElems, atomsByRun[child], segs, segChildren, &lastSeg)
			}
			continue
		}

		if child.Space == "w" && child.Tag == "hyperlink" {
			distributeHyperlinkToSegments(child, atomContribs, allAtomElems, atomsByRun, segs, segChildren, &lastSeg)
			continue
		}

		// Non-textual paragraph child (bookmarkStart, commentRangeStart, etc.)
		// Assign to the segment of the nearest preceding run.
		segChildren[lastSeg] = append(segChildren[lastSeg], child.Copy())
	}

	// Assemble fragments.
	return assembleFragments(segChildren, pPrEl, sectPrEl, numSegs, len(positions))
}

// --- pPr / sectPr extraction ---

// extractPPrAndSectPr finds <w:pPr> in the paragraph and separates the sectPr.
// Returns (pPr-without-sectPr or nil, sectPr or nil).
// The returned elements are copies; the original paragraph is not modified.
func extractPPrAndSectPr(pEl *etree.Element) (pPr *etree.Element, sectPr *etree.Element) {
	for _, child := range pEl.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			pPr = child
			break
		}
	}
	if pPr == nil {
		return nil, nil
	}

	// Check for sectPr inside pPr.
	for _, child := range pPr.ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			sectPr = child.Copy()
			break
		}
	}

	// Clone pPr without sectPr.
	pPrCopy := pPr.Copy()
	for _, child := range pPrCopy.ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			pPrCopy.RemoveChild(child)
			break
		}
	}

	return pPrCopy, sectPr
}

// --- Run splitting ---

// distributeRunToSegments builds trimmed <w:r> elements and appends them to
// segChildren. Updates lastSeg to the last segment this run contributed to.
func distributeRunToSegments(
	rEl *etree.Element,
	atomContribs map[*etree.Element][]atomSegInfo,
	allAtomElems map[*etree.Element]bool,
	runAtoms []textAtom,
	segs []segRange,
	segChildren [][]*etree.Element,
	lastSeg *int,
) {
	numSegs := len(segs)

	// Quick check: does this run have any atoms in any segment?
	hasContrib := false
	for _, a := range runAtoms {
		if _, ok := atomContribs[a.elem]; ok {
			hasContrib = true
			break
		}
	}
	if !hasContrib {
		return
	}

	// Find rPr.
	var rPrEl *etree.Element
	for _, child := range rEl.ChildElements() {
		if child.Space == "w" && child.Tag == "rPr" {
			rPrEl = child
			break
		}
	}

	// Per-segment run builders. Created lazily.
	runs := make([]*etree.Element, numSegs)

	getOrCreateRun := func(segIdx int) *etree.Element {
		if runs[segIdx] == nil {
			runs[segIdx] = newRunElement(rPrEl)
		}
		return runs[segIdx]
	}

	// Track position context for classifying non-atom children (§5.1.2).
	//   lastAtomSeg: segment of the last atom that contributed to a segment (-1 = none yet).
	//   inTagGap: true when the run cursor is inside a tag region (atom's
	//     tail was consumed by a tag, or atom was entirely in a tag).
	// Non-atom children are assigned to lastAtomSeg when outside a tag,
	// and dropped when inside a tag gap.
	lastAtomSeg := -1
	inTagGap := false

	for _, child := range rEl.ChildElements() {
		if child.Space == "w" && child.Tag == "rPr" {
			continue
		}

		if allAtomElems[child] {
			// This is an atom element.
			contribs, hasSegs := atomContribs[child]
			if !hasSegs {
				// Atom entirely within a tag range → drop it.
				// The cursor is now inside the tag gap.
				inTagGap = true
				continue
			}
			a := findAtomByElem(runAtoms, child)
			for _, c := range contribs {
				if a.editable {
					// Editable atom (<w:t>) — create trimmed <w:t>.
					text := a.text[c.textStart:c.textEnd]
					if text != "" {
						r := getOrCreateRun(c.segIndex)
						tEl := OxmlElement("w:t")
						tEl.SetText(text)
						ensurePreserveSpace(tEl)
						r.AddChild(tEl)
					}
				} else {
					// Fixed atom (<w:tab/>, <w:br/>, etc.) — clone the
					// original element. Fixed atoms are indivisible (1 char).
					r := getOrCreateRun(c.segIndex)
					r.AddChild(a.elem.Copy())
				}
				lastAtomSeg = c.segIndex
			}
			// Check if the atom's tail extends into a tag region.
			// If the last contribution doesn't reach the end of the atom's text,
			// the tail is consumed by a tag → cursor enters a tag gap.
			lastContrib := contribs[len(contribs)-1]
			inTagGap = lastContrib.textEnd < len(a.text)
		} else {
			// Non-atom child (drawing, lastRenderedPageBreak, etc.).
			if inTagGap {
				// Inside a tag gap (between intersected atoms) → drop (§5.1.2).
				continue
			}
			// Outside tag gap → assign to the nearest preceding atom's segment.
			seg := lastAtomSeg
			if seg < 0 {
				// Before any atom: find the first segment this run contributes to.
				seg = firstContributingSeg(runAtoms, atomContribs)
			}
			if seg >= 0 {
				r := getOrCreateRun(seg)
				r.AddChild(child.Copy())
			}
		}
	}

	// Add non-empty runs to segChildren.
	for i, r := range runs {
		if r != nil && hasRunContent(r) {
			segChildren[i] = append(segChildren[i], r)
			*lastSeg = i
		}
	}
}

// newRunElement creates a new <w:r> with a cloned <w:rPr> (if non-nil).
func newRunElement(rPrEl *etree.Element) *etree.Element {
	r := OxmlElement("w:r")
	if rPrEl != nil {
		r.AddChild(rPrEl.Copy())
	}
	return r
}

// hasRunContent reports whether a <w:r> has any children beyond <w:rPr>.
func hasRunContent(rEl *etree.Element) bool {
	for _, child := range rEl.ChildElements() {
		if !(child.Space == "w" && child.Tag == "rPr") {
			return true
		}
	}
	return false
}

// findAtomByElem finds the textAtom with the given element pointer.
func findAtomByElem(atoms []textAtom, elem *etree.Element) textAtom {
	for _, a := range atoms {
		if a.elem == elem {
			return a
		}
	}
	return textAtom{} // should not happen
}

// firstContributingSeg returns the first segment index that any atom of the
// given run contributes to, or -1 if none.
func firstContributingSeg(runAtoms []textAtom, atomContribs map[*etree.Element][]atomSegInfo) int {
	for _, a := range runAtoms {
		if cs, ok := atomContribs[a.elem]; ok && len(cs) > 0 {
			return cs[0].segIndex
		}
	}
	return -1
}

// --- Hyperlink splitting ---

// distributeHyperlinkToSegments builds trimmed <w:hyperlink> elements for each
// segment and appends them to segChildren. If runs of the hyperlink span
// multiple segments, the hyperlink is cloned into each segment with only the
// relevant runs (all clones preserve the original r:id).
func distributeHyperlinkToSegments(
	hlEl *etree.Element,
	atomContribs map[*etree.Element][]atomSegInfo,
	allAtomElems map[*etree.Element]bool,
	atomsByRun map[*etree.Element][]textAtom,
	segs []segRange,
	segChildren [][]*etree.Element,
	lastSeg *int,
) {
	numSegs := len(segs)

	// Collect per-segment runs from within this hyperlink.
	hlSegRuns := make([][]*etree.Element, numSegs)

	for _, child := range hlEl.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			// Build runs for segments — use a temporary segChildren to collect
			// them separately from paragraph-level children.
			tempSeg := make([][]*etree.Element, numSegs)
			tempLastSeg := *lastSeg
			distributeRunToSegments(child, atomContribs, allAtomElems, atomsByRun[child], segs, tempSeg, &tempLastSeg)
			for i := 0; i < numSegs; i++ {
				hlSegRuns[i] = append(hlSegRuns[i], tempSeg[i]...)
			}
		}
	}

	// Build hyperlink for each segment that has runs.
	for segIdx := 0; segIdx < numSegs; segIdx++ {
		if len(hlSegRuns[segIdx]) == 0 {
			continue
		}
		hl := cloneHyperlinkShell(hlEl)
		for _, r := range hlSegRuns[segIdx] {
			hl.AddChild(r)
		}
		segChildren[segIdx] = append(segChildren[segIdx], hl)
		*lastSeg = segIdx
	}
}

// cloneHyperlinkShell creates a copy of a <w:hyperlink> element with all its
// attributes (including r:id) but no children.
func cloneHyperlinkShell(hlEl *etree.Element) *etree.Element {
	hl := etree.NewElement(hlEl.Tag)
	hl.Space = hlEl.Space
	for _, attr := range hlEl.Attr {
		if attr.Space != "" {
			hl.CreateAttr(attr.Space+":"+attr.Key, attr.Value)
		} else {
			hl.CreateAttr(attr.Key, attr.Value)
		}
	}
	return hl
}

// --- Fragment assembly ---

// assembleFragments builds the final fragment list from per-segment children.
// Inserts placeholders between segments and handles sectPr placement.
func assembleFragments(
	segChildren [][]*etree.Element,
	pPr *etree.Element,
	sectPr *etree.Element,
	numSegs int,
	numTags int,
) []Fragment {
	var frags []Fragment

	for i := 0; i < numSegs; i++ {
		hasContent := len(segChildren[i]) > 0

		if hasContent {
			p := buildSegmentParagraph(segChildren[i], pPr)
			frags = append(frags, paragraphFragment(p))
		}

		// Insert placeholder after each segment except the last.
		if i < numTags {
			frags = append(frags, placeholderFragment())
		}
	}

	// Handle sectPr: must go on the LAST paragraph fragment.
	if sectPr != nil {
		applySectPrToFragments(&frags, pPr, sectPr)
	}

	return frags
}

// buildSegmentParagraph creates a <w:p> element with the given children and
// a cloned pPr (if non-nil).
func buildSegmentParagraph(children []*etree.Element, pPr *etree.Element) *etree.Element {
	p := OxmlElement("w:p")
	if pPr != nil {
		p.AddChild(pPr.Copy())
	}
	for _, child := range children {
		p.AddChild(child)
	}
	return p
}

// applySectPrToFragments ensures the LAST fragment carries the sectPr.
// If the last fragment is already a paragraph, sectPr is added to it.
// Otherwise (last is placeholder), an empty <w:p> is appended.
//
// Why not "find the last paragraph anywhere"? Because placing sectPr on a
// paragraph that precedes a placeholder would put the section break BEFORE
// the table, pushing the table into the next section (wrong page layout,
// orientation, headers/footers). See plan §6.
func applySectPrToFragments(frags *[]Fragment, pPr *etree.Element, sectPr *etree.Element) {
	n := len(*frags)
	if n > 0 && (*frags)[n-1].IsParagraph() {
		// Last fragment is already a paragraph — add sectPr to it.
		addSectPrToParagraph((*frags)[n-1].el, sectPr)
	} else {
		// Last fragment is a placeholder (or slice is empty).
		// Create trailing empty <w:p> to carry the sectPr.
		emptyP := buildEmptySectPrParagraph(pPr, sectPr)
		*frags = append(*frags, paragraphFragment(emptyP))
	}
}

// addSectPrToParagraph inserts a <w:sectPr> into the paragraph's <w:pPr>.
// Creates <w:pPr> if the paragraph doesn't have one.
func addSectPrToParagraph(pEl *etree.Element, sectPr *etree.Element) {
	var pPr *etree.Element
	for _, child := range pEl.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			pPr = child
			break
		}
	}
	if pPr == nil {
		pPr = OxmlElement("w:pPr")
		// Insert pPr as first child.
		pEl.InsertChildAt(0, pPr)
	}
	pPr.AddChild(sectPr.Copy())
}

// buildEmptySectPrParagraph creates a <w:p> containing only <w:pPr><w:sectPr>.
func buildEmptySectPrParagraph(pPrTemplate *etree.Element, sectPr *etree.Element) *etree.Element {
	p := OxmlElement("w:p")
	var pPr *etree.Element
	if pPrTemplate != nil {
		pPr = pPrTemplate.Copy()
	} else {
		pPr = OxmlElement("w:pPr")
	}
	pPr.AddChild(sectPr.Copy())
	p.AddChild(pPr)
	return p
}
