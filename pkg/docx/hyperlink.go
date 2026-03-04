package docx

import (
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Hyperlink is a proxy for a <w:hyperlink> element.
//
// A hyperlink occurs as a child of a paragraph at the same level as a Run.
// It contains runs which hold the visible text of the hyperlink.
//
// Mirrors Python Hyperlink(Parented).
type Hyperlink struct {
	hyperlink *oxml.CT_Hyperlink
	part      *parts.StoryPart
}

// newHyperlink creates a new Hyperlink proxy.
func newHyperlink(elm *oxml.CT_Hyperlink, part *parts.StoryPart) *Hyperlink {
	return &Hyperlink{hyperlink: elm, part: part}
}

// Address returns the URL of the hyperlink (resolved from the relationship).
//
// When this hyperlink is an internal jump (e.g. TOC entry), the address is blank.
// The bookmark reference is in the Fragment property.
//
// Mirrors Python Hyperlink.address.
func (h *Hyperlink) Address() string {
	rId := h.hyperlink.RId()
	if rId == "" {
		return ""
	}
	rels := h.part.Rels()
	if rels == nil {
		return ""
	}
	rel := rels.GetByRID(rId)
	if rel == nil {
		return ""
	}
	return rel.TargetRef
}

// ContainsPageBreak returns true when the text of this hyperlink is broken
// across page boundaries.
//
// Mirrors Python Hyperlink.contains_page_break.
func (h *Hyperlink) ContainsPageBreak() bool {
	return len(h.hyperlink.HyperlinkLastRenderedPageBreaks()) > 0
}

// Fragment returns the URI fragment (e.g. bookmark reference) without the "#".
//
// Mirrors Python Hyperlink.fragment.
func (h *Hyperlink) Fragment() string {
	return h.hyperlink.Anchor()
}

// Runs returns the list of Run instances in this hyperlink.
//
// Mirrors Python Hyperlink.runs.
func (h *Hyperlink) Runs() []*Run {
	rList := h.hyperlink.RList()
	result := make([]*Run, len(rList))
	for i, r := range rList {
		result[i] = newRun(r, h.part)
	}
	return result
}

// Text returns the concatenated text of all runs in this hyperlink.
//
// Mirrors Python Hyperlink.text.
func (h *Hyperlink) Text() string {
	return h.hyperlink.HyperlinkText()
}

// URL returns the full URL (address + fragment) for convenience.
//
// Returns "" when there is no address portion, distinguishing external URIs
// from internal jump hyperlinks.
//
// Mirrors Python Hyperlink.url.
func (h *Hyperlink) URL() string {
	address := h.Address()
	if address == "" {
		return ""
	}
	fragment := h.Fragment()
	if fragment != "" {
		return address + "#" + fragment
	}
	return address
}
