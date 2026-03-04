package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Paragraph is a proxy object wrapping a <w:p> element.
//
// Mirrors Python Paragraph(StoryChild).
type Paragraph struct {
	p    *oxml.CT_P
	part *parts.StoryPart
}

// newParagraph creates a new Paragraph proxy.
func newParagraph(p *oxml.CT_P, part *parts.StoryPart) *Paragraph {
	return &Paragraph{p: p, part: part}
}

// AddRun appends a run containing text and optionally styled with style.
// text may contain tab (\t) and newline (\n, \r) characters which are
// converted to their XML equivalents.
//
// Mirrors Python Paragraph.add_run.
func (para *Paragraph) AddRun(text string, style ...StyleRef) (*Run, error) {
	r := para.p.AddR()
	run := newRun(r, para.part)
	if text != "" {
		run.SetText(text)
	}
	if raw := resolveStyleRef(style); raw != nil {
		if err := run.setStyleRaw(raw); err != nil {
			return nil, fmt.Errorf("docx: setting run style: %w", err)
		}
	}
	return run, nil
}

// Alignment returns the paragraph alignment, or nil if not set (inherited).
//
// Mirrors Python Paragraph.alignment (getter).
func (para *Paragraph) Alignment() (*enum.WdParagraphAlignment, error) {
	return para.p.Alignment()
}

// SetAlignment sets the paragraph alignment. Passing nil removes the setting.
//
// Mirrors Python Paragraph.alignment (setter).
func (para *Paragraph) SetAlignment(v *enum.WdParagraphAlignment) error {
	return para.p.SetAlignment(v)
}

// Clear removes all content from this paragraph, preserving formatting.
//
// Mirrors Python Paragraph.clear.
func (para *Paragraph) Clear() {
	para.p.ClearContent()
}

// ContainsPageBreak returns true when one or more rendered page-breaks occur.
//
// Mirrors Python Paragraph.contains_page_break.
func (para *Paragraph) ContainsPageBreak() bool {
	return len(para.p.LastRenderedPageBreaks()) > 0
}

// Hyperlinks returns a Hyperlink for each hyperlink in this paragraph.
//
// Mirrors Python Paragraph.hyperlinks.
func (para *Paragraph) Hyperlinks() []*Hyperlink {
	var result []*Hyperlink
	for _, child := range para.p.RawElement().ChildElements() {
		if child.Space == "w" && child.Tag == "hyperlink" {
			hl := &oxml.CT_Hyperlink{Element: oxml.WrapElement(child)}
			result = append(result, newHyperlink(hl, para.part))
		}
	}
	return result
}

// InsertParagraphBefore creates and returns a new paragraph inserted directly
// before this one. If text is non-empty, it is placed in a single run.
//
// Mirrors Python Paragraph.insert_paragraph_before.
func (para *Paragraph) InsertParagraphBefore(text string, style ...StyleRef) (*Paragraph, error) {
	newP := para.p.AddPBefore()
	if newP == nil {
		return nil, fmt.Errorf("docx: cannot insert paragraph before (no parent)")
	}
	p := newParagraph(newP, para.part)
	if text != "" {
		if _, err := p.AddRun(text); err != nil {
			return nil, err
		}
	}
	if raw := resolveStyleRef(style); raw != nil {
		if err := p.setStyleRaw(raw); err != nil {
			return nil, err
		}
	}
	return p, nil
}

// IterInnerContent returns runs and hyperlinks in this paragraph in document order.
//
// Mirrors Python Paragraph.iter_inner_content.
func (para *Paragraph) IterInnerContent() []*InlineItem {
	var result []*InlineItem
	for _, item := range para.p.InnerContentElements() {
		switch v := item.(type) {
		case *oxml.CT_R:
			result = append(result, &InlineItem{run: newRun(v, para.part)})
		case *oxml.CT_Hyperlink:
			result = append(result, &InlineItem{hyperlink: newHyperlink(v, para.part)})
		}
	}
	return result
}

// ParagraphFormat returns the ParagraphFormat providing access to formatting
// properties like line spacing and indentation.
//
// Mirrors Python Paragraph.paragraph_format.
func (para *Paragraph) ParagraphFormat() *ParagraphFormat {
	return newParagraphFormatFromP(para.p)
}

// RenderedPageBreaks returns all rendered page-breaks in this paragraph.
//
// Mirrors Python Paragraph.rendered_page_breaks.
func (para *Paragraph) RenderedPageBreaks() []*RenderedPageBreak {
	var result []*RenderedPageBreak
	for _, lrpb := range para.p.LastRenderedPageBreaks() {
		result = append(result, newRenderedPageBreak(lrpb, para.part))
	}
	return result
}

// Runs returns all runs in this paragraph.
//
// Mirrors Python Paragraph.runs.
func (para *Paragraph) Runs() []*Run {
	var result []*Run
	for _, r := range para.p.RList() {
		result = append(result, newRun(r, para.part))
	}
	return result
}

// Style returns the paragraph style, delegating to the part for resolution.
//
// Mirrors Python Paragraph.style (getter).
func (para *Paragraph) Style() (*oxml.CT_Style, error) {
	styleID, err := para.p.Style()
	if err != nil {
		return nil, err
	}
	return para.part.GetStyle(styleID, enum.WdStyleTypeParagraph)
}

// SetStyle sets the paragraph style. style can be a string name or nil.
//
// Mirrors Python Paragraph.style (setter).
func (para *Paragraph) SetStyle(style StyleRef) error {
	return para.setStyleRaw(resolveStyleRef([]StyleRef{style}))
}

// setStyleRaw passes the raw style value (string or styledObject) to the parts layer.
func (para *Paragraph) setStyleRaw(raw any) error {
	styleID, err := para.part.GetStyleID(raw, enum.WdStyleTypeParagraph)
	if err != nil {
		return err
	}
	return para.p.SetStyle(styleID)
}

// Text returns the full textual content of this paragraph.
//
// Mirrors Python Paragraph.text (getter).
func (para *Paragraph) Text() string {
	return para.p.ParagraphText()
}

// SetText replaces all paragraph content with a single run containing text.
//
// Mirrors Python Paragraph.text (setter).
func (para *Paragraph) SetText(text string) error {
	para.Clear()
	_, err := para.AddRun(text)
	return err
}

// ReplaceText replaces all occurrences of old with new in the text of this
// paragraph. Works across run boundaries, including text inside hyperlinks.
// Preserves formatting and XML structure.
//
// Returns the number of replacements performed.
// Returns 0 if old == new (no-op optimization — XML is not modified even
// though occurrences may exist in the text).
func (para *Paragraph) ReplaceText(old, new string) int {
	return para.p.ReplaceText(old, new)
}

// CT_P returns the underlying oxml element.
func (para *Paragraph) CT_P() *oxml.CT_P { return para.p }
