package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// RenderedPageBreak represents a page-break inserted by Word during page-layout.
//
// These usually do not correspond to "hard" page-breaks; rather Word ran out of
// room on one page and needed to start another. Their position can change with
// printer, page-size, margins, and edits.
//
// Note: python-docx (and go-docx) never inserts these because it has no
// rendering function. They are useful only for text-extraction of existing documents.
//
// Mirrors Python RenderedPageBreak(Parented).
type RenderedPageBreak struct {
	lrpb *oxml.CT_LastRenderedPageBreak
	part *parts.StoryPart
}

// newRenderedPageBreak creates a new RenderedPageBreak proxy.
func newRenderedPageBreak(elm *oxml.CT_LastRenderedPageBreak, part *parts.StoryPart) *RenderedPageBreak {
	return &RenderedPageBreak{lrpb: elm, part: part}
}

// PrecedingParagraphFragment returns a "loose" paragraph containing the content
// preceding this page-break. Returns nil when no content precedes this break
// (common case: page breaks on an even paragraph boundary).
//
// The returned paragraph is divorced from the document body; changes to it
// will not be reflected in the document.
//
// Contains the entire hyperlink when this break occurs within a hyperlink.
//
// Mirrors Python RenderedPageBreak.preceding_paragraph_fragment.
func (rpb *RenderedPageBreak) PrecedingParagraphFragment() (*Paragraph, error) {
	if rpb.lrpb.PrecedesAllContent() {
		return nil, nil
	}

	fragP, err := rpb.lrpb.PrecedingFragmentP()
	if err != nil {
		return nil, fmt.Errorf("docx: building preceding fragment: %w", err)
	}
	return newParagraph(fragP, rpb.part), nil
}

// FollowingParagraphFragment returns a "loose" paragraph containing the content
// following this page-break. Returns nil when no content follows this break
// (unlikely but possible).
//
// The returned paragraph is divorced from the document body; changes to it
// will not be reflected in the document.
//
// Contains no portion of the hyperlink when this break occurs within a hyperlink.
//
// Mirrors Python RenderedPageBreak.following_paragraph_fragment.
func (rpb *RenderedPageBreak) FollowingParagraphFragment() (*Paragraph, error) {
	if rpb.lrpb.FollowsAllContent() {
		return nil, nil
	}

	fragP, err := rpb.lrpb.FollowingFragmentP()
	if err != nil {
		return nil, fmt.Errorf("docx: building following fragment: %w", err)
	}
	return newParagraph(fragP, rpb.part), nil
}
