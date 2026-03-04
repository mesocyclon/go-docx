package docx

import (
	"fmt"
	"io"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Run is a proxy object wrapping a <w:r> element.
//
// Mirrors Python Run(StoryChild).
type Run struct {
	r    *oxml.CT_R
	part *parts.StoryPart
}

// newRun creates a new Run proxy.
func newRun(r *oxml.CT_R, part *parts.StoryPart) *Run {
	return &Run{r: r, part: part}
}

// AddBreak adds a break element of the given type to this run.
//
// Mirrors Python Run.add_break. Maps break_type to (type_, clear) pairs.
func (run *Run) AddBreak(breakType enum.WdBreakType) error {
	type_, clear, err := breakTypeToAttrs(breakType)
	if err != nil {
		return err
	}
	br := run.r.NewDetachedBr()
	if type_ != "" {
		if err := br.SetType(type_); err != nil {
			return err
		}
	}
	if clear != "" {
		if err := br.SetClear(clear); err != nil {
			return err
		}
	}
	run.r.AttachBr(br)
	return nil
}

// breakTypeToAttrs maps WdBreakType to (type, clear) attribute values.
// Returns an error for break types not applicable to a run (e.g. section breaks).
func breakTypeToAttrs(bt enum.WdBreakType) (string, string, error) {
	switch bt {
	case enum.WdBreakTypeLine:
		return "", "", nil
	case enum.WdBreakTypePage:
		return "page", "", nil
	case enum.WdBreakTypeColumn:
		return "column", "", nil
	case enum.WdBreakTypeLineClearLeft:
		return "textWrapping", "left", nil
	case enum.WdBreakTypeLineClearRight:
		return "textWrapping", "right", nil
	case enum.WdBreakTypeLineClearAll:
		return "textWrapping", "all", nil
	default:
		return "", "", fmt.Errorf("docx: unsupported break type for Run.AddBreak: %d", bt)
	}
}

// AddPicture adds an inline picture to this run from an image stream and
// returns the InlineShape. Width and height are optional EMU dimensions;
// pass nil for native size or to compute proportionally.
//
// Mirrors Python Run.add_picture(image_path_or_stream, width, height).
func (run *Run) AddPicture(r io.ReadSeeker, width, height *int64) (*InlineShape, error) {
	if run.part == nil {
		return nil, fmt.Errorf("docx: run has no story part (required for image insertion)")
	}
	inline, err := run.part.NewPicInlineFromReader(r, width, height)
	if err != nil {
		return nil, fmt.Errorf("docx: creating pic inline from stream: %w", err)
	}
	run.r.AddDrawingWithInline(inline)
	return newInlineShape(inline, run.part), nil
}

// AddPictureFromPart adds an inline picture from a pre-built ImagePart.
// This is the lower-level API; prefer AddPicture for the standard flow.
func (run *Run) AddPictureFromPart(imgPart *parts.ImagePart, width, height *int64) (*InlineShape, error) {
	inline, err := run.part.NewPicInline(imgPart, width, height)
	if err != nil {
		return nil, fmt.Errorf("docx: creating pic inline: %w", err)
	}
	run.r.AddDrawingWithInline(inline)
	return newInlineShape(inline, run.part), nil
}

// AddTab adds a <w:tab/> element at the end of the run.
//
// Mirrors Python Run.add_tab.
func (run *Run) AddTab() {
	run.r.AddTab()
}

// AddText appends a <w:t> element with the given text to the run.
//
// Mirrors Python Run.add_text.
func (run *Run) AddText(text string) {
	run.r.AddTWithText(text)
}

// Bold returns the tri-state bold value (delegates to Font).
//
// Mirrors Python Run.bold (getter).
func (run *Run) Bold() *bool {
	return run.Font().Bold()
}

// SetBold sets the tri-state bold value (delegates to Font).
//
// Mirrors Python Run.bold (setter).
func (run *Run) SetBold(v *bool) error {
	return run.Font().SetBold(v)
}

// Clear removes all content from this run, preserving formatting.
//
// Mirrors Python Run.clear.
func (run *Run) Clear() {
	run.r.ClearContent()
}

// ContainsPageBreak returns true when rendered page-breaks occur in this run.
//
// Mirrors Python Run.contains_page_break.
func (run *Run) ContainsPageBreak() bool {
	return len(run.r.LastRenderedPageBreaks()) > 0
}

// Font returns the Font providing access to character formatting properties.
//
// Mirrors Python Run.font.
func (run *Run) Font() *Font {
	return newFont(run.r)
}

// Italic returns the tri-state italic value (delegates to Font).
//
// Mirrors Python Run.italic (getter).
func (run *Run) Italic() *bool {
	return run.Font().Italic()
}

// SetItalic sets the tri-state italic value (delegates to Font).
//
// Mirrors Python Run.italic (setter).
func (run *Run) SetItalic(v *bool) error {
	return run.Font().SetItalic(v)
}

// MarkCommentRange marks the range of runs from this run to lastRun as
// belonging to the comment identified by commentID.
//
// Mirrors Python Run.mark_comment_range.
func (run *Run) MarkCommentRange(lastRun *Run, commentID int) error {
	run.r.InsertCommentRangeStartAbove(commentID)
	lastRun.r.InsertCommentRangeEndAndReferenceBelow(commentID)
	return nil
}

// Style returns the character style applied to this run.
//
// Mirrors Python Run.style (getter).
func (run *Run) Style() (*oxml.CT_Style, error) {
	styleID, err := run.r.Style()
	if err != nil {
		return nil, err
	}
	return run.part.GetStyle(styleID, enum.WdStyleTypeCharacter)
}

// SetStyle sets the character style. style can be a string name or nil.
//
// Mirrors Python Run.style (setter).
func (run *Run) SetStyle(style StyleRef) error {
	return run.setStyleRaw(resolveStyleRef([]StyleRef{style}))
}

// setStyleRaw passes the raw style value to the parts layer.
func (run *Run) setStyleRaw(raw any) error {
	styleID, err := run.part.GetStyleID(raw, enum.WdStyleTypeCharacter)
	if err != nil {
		return err
	}
	return run.r.SetStyle(styleID)
}

// Text returns the textual content of this run.
//
// Mirrors Python Run.text (getter).
func (run *Run) Text() string {
	return run.r.RunText()
}

// SetText replaces all run content with elements representing the given text.
//
// Mirrors Python Run.text (setter).
func (run *Run) SetText(text string) {
	run.r.SetRunText(text)
}

// Underline returns the underline value (delegates to Font).
//
// Mirrors Python Run.underline (getter).
func (run *Run) Underline() (*UnderlineVal, error) {
	return run.Font().Underline()
}

// SetUnderline sets the underline value (delegates to Font).
// Pass nil to inherit.
//
// Mirrors Python Run.underline (setter).
func (run *Run) SetUnderline(v *UnderlineVal) error {
	return run.Font().SetUnderline(v)
}

// CT_R returns the underlying oxml element.
func (run *Run) CT_R() *oxml.CT_R { return run.r }

// RunContentItem represents one item from a run's inner content:
// either a string (accumulated text), *Drawing, or *RenderedPageBreak.
type RunContentItem struct {
	text              *string
	drawing           *Drawing
	renderedPageBreak *RenderedPageBreak
}

// IsText returns true if this item is accumulated text.
func (it *RunContentItem) IsText() bool { return it.text != nil }

// IsDrawing returns true if this item is a Drawing.
func (it *RunContentItem) IsDrawing() bool { return it.drawing != nil }

// IsRenderedPageBreak returns true if this item is a RenderedPageBreak.
func (it *RunContentItem) IsRenderedPageBreak() bool { return it.renderedPageBreak != nil }

// Text returns the text string, or "" if this item is not text.
func (it *RunContentItem) Text() string {
	if it.text != nil {
		return *it.text
	}
	return ""
}

// Drawing returns the Drawing, or nil if this item is not a drawing.
func (it *RunContentItem) Drawing() *Drawing { return it.drawing }

// RenderedPageBreak returns the RenderedPageBreak, or nil if not a page break.
func (it *RunContentItem) RenderedPageBreak() *RenderedPageBreak { return it.renderedPageBreak }

// IterInnerContent returns the content items in this run in the order they appear.
//
// Text-like elements (w:t, w:br, w:cr, w:tab, etc.) are accumulated into
// contiguous strings. Drawing and rendered-page-break elements are yielded
// individually, interrupting any accumulated text.
//
// Mirrors Python Run.iter_inner_content → yields str | Drawing | RenderedPageBreak.
func (run *Run) IterInnerContent() []*RunContentItem {
	var result []*RunContentItem
	for _, item := range run.r.InnerContentItems() {
		switch v := item.(type) {
		case string:
			s := v
			result = append(result, &RunContentItem{text: &s})
		case *oxml.CT_Drawing:
			result = append(result, &RunContentItem{drawing: newDrawing(v, run.part)})
		case *oxml.CT_LastRenderedPageBreak:
			result = append(result, &RunContentItem{renderedPageBreak: newRenderedPageBreak(v, run.part)})
		}
	}
	return result
}
