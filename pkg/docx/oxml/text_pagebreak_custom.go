package oxml

import (
	"fmt"

	"github.com/beevik/etree"
)

// --- CT_LastRenderedPageBreak custom methods ---

// EnclosingP returns the w:p ancestor of this lastRenderedPageBreak.
// Handles both direct run children and runs inside hyperlinks.
func (pb *CT_LastRenderedPageBreak) EnclosingP() *CT_P {
	// Walk up: lastRenderedPageBreak → w:r → w:p (or w:r → w:hyperlink → w:p)
	parent := pb.e.Parent()
	if parent == nil {
		return nil
	}
	// parent is w:r
	grandparent := parent.Parent()
	if grandparent == nil {
		return nil
	}
	// If grandparent is w:hyperlink, go one more level
	if grandparent.Space == "w" && grandparent.Tag == "hyperlink" {
		ggp := grandparent.Parent()
		if ggp == nil {
			return nil
		}
		return &CT_P{Element{e: ggp}}
	}
	return &CT_P{Element{e: grandparent}}
}

// IsInHyperlink returns true when this page-break is embedded in a hyperlink run.
func (pb *CT_LastRenderedPageBreak) IsInHyperlink() bool {
	parent := pb.e.Parent() // w:r
	if parent == nil {
		return false
	}
	grandparent := parent.Parent()
	if grandparent == nil {
		return false
	}
	return grandparent.Space == "w" && grandparent.Tag == "hyperlink"
}

// PrecedesAllContent returns true when this page-break precedes all paragraph content.
// This is a common case occurring when the page breaks on an even paragraph boundary.
func (pb *CT_LastRenderedPageBreak) PrecedesAllContent() bool {
	if pb.IsInHyperlink() {
		return false
	}

	p := pb.EnclosingP()
	if p == nil {
		return false
	}

	// Check that this is in the first w:r of the paragraph
	// and no content-bearing sibling precedes it in that run
	parent := pb.e.Parent() // enclosing w:r
	if parent == nil {
		return false
	}

	// Is this the first run (ignoring w:pPr)?
	firstRun := firstRunInP(p.e)
	if firstRun == nil || firstRun != parent {
		return false
	}

	// Check no content-bearing elements precede this lrpb in the run
	for _, sibling := range parent.ChildElements() {
		if sibling == pb.e {
			return true // reached ourselves without seeing content
		}
		if isRunInnerContent(sibling) {
			return false
		}
	}
	return false
}

// FollowsAllContent returns true when this page-break is the last content in the paragraph.
func (pb *CT_LastRenderedPageBreak) FollowsAllContent() bool {
	if pb.IsInHyperlink() {
		return false
	}

	p := pb.EnclosingP()
	if p == nil {
		return false
	}

	parent := pb.e.Parent() // enclosing w:r

	// Is this the last run?
	lastRun := lastRunInP(p.e)
	if lastRun == nil || lastRun != parent {
		return false
	}

	// Check no content-bearing elements follow this lrpb in the run
	pastBreak := false
	for _, sibling := range parent.ChildElements() {
		if sibling == pb.e {
			pastBreak = true
			continue
		}
		if pastBreak && isRunInnerContent(sibling) {
			return false
		}
	}
	return pastBreak
}

// PrecedingFragmentP returns a clone of the enclosing w:p containing only the content
// before this page break. Only valid on the first rendered page-break in the paragraph.
func (pb *CT_LastRenderedPageBreak) PrecedingFragmentP() (*CT_P, error) {
	if pb.IsInHyperlink() {
		return pb.precedingFragInHlink()
	}
	return pb.precedingFragInRun()
}

// FollowingFragmentP returns a clone of the enclosing w:p containing only the content
// after this page break. Only valid on the first rendered page-break in the paragraph.
func (pb *CT_LastRenderedPageBreak) FollowingFragmentP() (*CT_P, error) {
	if pb.IsInHyperlink() {
		return pb.followingFragInHlink()
	}
	return pb.followingFragInRun()
}

// precedingFragInRun creates the preceding fragment when break is in a plain run.
func (pb *CT_LastRenderedPageBreak) precedingFragInRun() (*CT_P, error) {
	p := pb.EnclosingP()
	if p == nil {
		return nil, fmt.Errorf("no enclosing <w:p> found")
	}

	// Deep copy the paragraph
	cloneP := p.e.Copy()
	enclosingR := pb.e.Parent()

	// Find the corresponding run and lrpb in the clone
	cloneLrpb, cloneRun := findLrpbInClone(cloneP, enclosingR, pb.e)
	if cloneLrpb == nil || cloneRun == nil {
		return nil, fmt.Errorf("could not locate page break in clone")
	}

	// Remove all p inner-content following the enclosing run
	removeFollowingSiblings(cloneP, cloneRun)

	// Remove all run inner-content following lrpb, and the lrpb itself
	removeFollowingSiblings(cloneRun, cloneLrpb)
	cloneRun.RemoveChild(cloneLrpb)

	return &CT_P{Element{e: cloneP}}, nil
}

// followingFragInRun creates the following fragment when break is in a plain run.
func (pb *CT_LastRenderedPageBreak) followingFragInRun() (*CT_P, error) {
	p := pb.EnclosingP()
	if p == nil {
		return nil, fmt.Errorf("no enclosing <w:p> found")
	}

	cloneP := p.e.Copy()
	enclosingR := pb.e.Parent()

	cloneLrpb, cloneRun := findLrpbInClone(cloneP, enclosingR, pb.e)
	if cloneLrpb == nil || cloneRun == nil {
		return nil, fmt.Errorf("could not locate page break in clone")
	}

	// Remove all p inner-content preceding the enclosing run (but not w:pPr)
	removePrecedingContentSiblings(cloneP, cloneRun)

	// Remove all run inner-content preceding lrpb (but not w:rPr) and the lrpb itself
	removePrecedingRunContent(cloneRun, cloneLrpb)
	cloneRun.RemoveChild(cloneLrpb)

	return &CT_P{Element{e: cloneP}}, nil
}

// precedingFragInHlink creates the preceding fragment when break is inside a hyperlink.
func (pb *CT_LastRenderedPageBreak) precedingFragInHlink() (*CT_P, error) {
	p := pb.EnclosingP()
	if p == nil {
		return nil, fmt.Errorf("no enclosing <w:p> found")
	}

	cloneP := p.e.Copy()
	enclosingR := pb.e.Parent()
	enclosingHlink := enclosingR.Parent()

	// Find the hyperlink in clone by matching position
	cloneHlink := findMatchingChild(cloneP, enclosingHlink)
	if cloneHlink == nil {
		return nil, fmt.Errorf("could not locate hyperlink in clone")
	}

	// Find lrpb in clone hyperlink
	cloneLrpb := findLrpbInElement(cloneHlink)
	if cloneLrpb == nil {
		return nil, fmt.Errorf("could not locate page break in clone")
	}

	// Remove all p inner-content following the hyperlink
	removeFollowingSiblings(cloneP, cloneHlink)

	// Remove the page-break from inside the hyperlink
	// (entire hyperlink goes into preceding fragment)
	cloneLrpb.Parent().RemoveChild(cloneLrpb)

	return &CT_P{Element{e: cloneP}}, nil
}

// followingFragInHlink creates the following fragment when break is inside a hyperlink.
func (pb *CT_LastRenderedPageBreak) followingFragInHlink() (*CT_P, error) {
	p := pb.EnclosingP()
	if p == nil {
		return nil, fmt.Errorf("no enclosing <w:p> found")
	}

	cloneP := p.e.Copy()
	enclosingR := pb.e.Parent()
	enclosingHlink := enclosingR.Parent()

	cloneHlink := findMatchingChild(cloneP, enclosingHlink)
	if cloneHlink == nil {
		return nil, fmt.Errorf("could not locate hyperlink in clone")
	}

	// Remove all p inner-content preceding the hyperlink (but not pPr)
	removePrecedingContentSiblings(cloneP, cloneHlink)

	// Remove the entire hyperlink (it belongs to the preceding fragment)
	cloneP.RemoveChild(cloneHlink)

	return &CT_P{Element{e: cloneP}}, nil
}

// --- Helper functions ---

// firstRunInP returns the first w:r child of a w:p element.
func firstRunInP(p *etree.Element) *etree.Element {
	for _, child := range p.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			return child
		}
	}
	return nil
}

// lastRunInP returns the last w:r child of a w:p element.
func lastRunInP(p *etree.Element) *etree.Element {
	children := p.ChildElements()
	for i := len(children) - 1; i >= 0; i-- {
		if children[i].Space == "w" && children[i].Tag == "r" {
			return children[i]
		}
	}
	return nil
}

// isRunInnerContent returns true if the element is a content-bearing run child
// (w:br, w:cr, w:drawing, w:noBreakHyphen, w:ptab, w:t, w:tab).
func isRunInnerContent(e *etree.Element) bool {
	if e.Space != "w" {
		return false
	}
	switch e.Tag {
	case "br", "cr", "drawing", "noBreakHyphen", "ptab", "t", "tab":
		return true
	}
	return false
}

// findLrpbInClone locates the corresponding lastRenderedPageBreak in a cloned w:p
// by matching the positional path from the original.
func findLrpbInClone(cloneP, origRun, origLrpb *etree.Element) (*etree.Element, *etree.Element) {
	// Find the run index in original p
	origP := origRun.Parent()
	if origP == nil {
		return nil, nil
	}
	// Handle hyperlink wrapper
	if origP.Space == "w" && origP.Tag == "hyperlink" {
		origP = origP.Parent()
	}

	runIndex := elementIndex(origP, origRun)
	lrpbIndex := elementIndex(origRun, origLrpb)

	// Find the same indices in the clone
	children := cloneP.ChildElements()
	if runIndex >= len(children) {
		return nil, nil
	}
	cloneRun := children[runIndex]
	runChildren := cloneRun.ChildElements()
	if lrpbIndex >= len(runChildren) {
		return nil, nil
	}
	return runChildren[lrpbIndex], cloneRun
}

// elementIndex returns the index of child among parent's ChildElements.
func elementIndex(parent, child *etree.Element) int {
	for i, c := range parent.ChildElements() {
		if c == child {
			return i
		}
	}
	return -1
}

// findMatchingChild finds a child in clone that matches the original's position.
func findMatchingChild(cloneParent, origChild *etree.Element) *etree.Element {
	origParent := origChild.Parent()
	if origParent == nil {
		return nil
	}
	idx := elementIndex(origParent, origChild)
	children := cloneParent.ChildElements()
	if idx >= 0 && idx < len(children) {
		return children[idx]
	}
	return nil
}

// findLrpbInElement finds the first w:lastRenderedPageBreak inside any w:r child.
func findLrpbInElement(el *etree.Element) *etree.Element {
	for _, child := range el.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			for _, gc := range child.ChildElements() {
				if gc.Space == "w" && gc.Tag == "lastRenderedPageBreak" {
					return gc
				}
			}
		}
		// Also check nested
		if child.Space == "w" && child.Tag == "lastRenderedPageBreak" {
			return child
		}
	}
	return nil
}

// removeFollowingSiblings removes all element siblings after ref in parent.
func removeFollowingSiblings(parent, ref *etree.Element) {
	found := false
	var toRemove []*etree.Element
	for _, child := range parent.ChildElements() {
		if child == ref {
			found = true
			continue
		}
		if found {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		parent.RemoveChild(child)
	}
}

// removePrecedingContentSiblings removes all element siblings before ref in parent,
// except w:pPr elements.
func removePrecedingContentSiblings(parent, ref *etree.Element) {
	var toRemove []*etree.Element
	for _, child := range parent.ChildElements() {
		if child == ref {
			break
		}
		if child.Space == "w" && child.Tag == "pPr" {
			continue
		}
		toRemove = append(toRemove, child)
	}
	for _, child := range toRemove {
		parent.RemoveChild(child)
	}
}

// removePrecedingRunContent removes run inner-content before ref in a w:r element,
// preserving w:rPr.
func removePrecedingRunContent(run, ref *etree.Element) {
	var toRemove []*etree.Element
	for _, child := range run.ChildElements() {
		if child == ref {
			break
		}
		if child.Space == "w" && child.Tag == "rPr" {
			continue
		}
		toRemove = append(toRemove, child)
	}
	for _, child := range toRemove {
		run.RemoveChild(child)
	}
}
