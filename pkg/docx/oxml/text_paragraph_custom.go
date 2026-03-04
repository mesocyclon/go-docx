package oxml

import (
	"strings"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// --- CT_P custom methods ---

// AddPBefore creates a new <w:p> element inserted directly prior to this one.
// Returns nil if this paragraph has no parent element.
func (p *CT_P) AddPBefore() *CT_P {
	parent := p.e.Parent()
	if parent == nil {
		return nil
	}
	newP := OxmlElement("w:p")
	insertBefore(parent, newP, p.e)
	return &CT_P{Element{e: newP}}
}

// Alignment returns the paragraph alignment from pPr/jc, or nil if not set.
func (p *CT_P) Alignment() (*enum.WdParagraphAlignment, error) {
	pPr := p.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.JcVal()
}

// SetAlignment sets the paragraph alignment. Passing nil removes the jc element.
func (p *CT_P) SetAlignment(val *enum.WdParagraphAlignment) error {
	pPr := p.GetOrAddPPr()
	return pPr.SetJcVal(val)
}

// ClearContent removes all child elements except <w:pPr>.
func (p *CT_P) ClearContent() {
	var toRemove []*etree.Element
	for _, child := range p.e.ChildElements() {
		if !(child.Space == "w" && child.Tag == "pPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		p.e.RemoveChild(child)
	}
}

// InnerContentElements returns run and hyperlink children of the <w:p> element,
// in document order.
func (p *CT_P) InnerContentElements() []InlineItem {
	var result []InlineItem
	for _, child := range p.e.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			result = append(result, &CT_R{Element{e: child}})
		} else if child.Space == "w" && child.Tag == "hyperlink" {
			result = append(result, &CT_Hyperlink{Element{e: child}})
		}
	}
	return result
}

// LastRenderedPageBreaks returns all <w:lastRenderedPageBreak> descendants in this paragraph.
// Searches both direct runs and runs inside hyperlinks.
func (p *CT_P) LastRenderedPageBreaks() []*CT_LastRenderedPageBreak {
	var result []*CT_LastRenderedPageBreak
	for _, child := range p.e.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			for _, gc := range child.ChildElements() {
				if gc.Space == "w" && gc.Tag == "lastRenderedPageBreak" {
					result = append(result, &CT_LastRenderedPageBreak{Element{e: gc}})
				}
			}
		} else if child.Space == "w" && child.Tag == "hyperlink" {
			for _, run := range child.ChildElements() {
				if run.Space == "w" && run.Tag == "r" {
					for _, gc := range run.ChildElements() {
						if gc.Space == "w" && gc.Tag == "lastRenderedPageBreak" {
							result = append(result, &CT_LastRenderedPageBreak{Element{e: gc}})
						}
					}
				}
			}
		}
	}
	return result
}

// SetSectPr unconditionally replaces or adds sectPr as grandchild in correct sequence.
func (p *CT_P) SetSectPr(sectPr *CT_SectPr) {
	pPr := p.GetOrAddPPr()
	pPr.RemoveSectPr()
	pPr.insertSectPr(sectPr)
}

// Style returns the styleId of the paragraph, or nil if not set.
func (p *CT_P) Style() (*string, error) {
	pPr := p.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.StyleVal()
}

// SetStyle sets the paragraph style. Passing nil removes the pStyle element.
func (p *CT_P) SetStyle(styleID *string) error {
	pPr := p.GetOrAddPPr()
	if err := pPr.SetStyleVal(styleID); err != nil {
		return err
	}
	return nil
}

// ParagraphText returns the full text of the paragraph by concatenating text from
// all run and hyperlink children. Named ParagraphText to avoid conflict with
// embedded Element.Text().
func (p *CT_P) ParagraphText() string {
	var sb strings.Builder
	for _, child := range p.e.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			r := &CT_R{Element{e: child}}
			sb.WriteString(r.RunText())
		} else if child.Space == "w" && child.Tag == "hyperlink" {
			h := &CT_Hyperlink{Element{e: child}}
			sb.WriteString(h.HyperlinkText())
		}
	}
	return sb.String()
}

// ReplaceText replaces all non-overlapping occurrences of old with new in the
// text of this paragraph, correctly handling replacements that span across
// run and hyperlink boundaries.
//
// Formatting (<w:rPr>), non-textual elements (<w:drawing>, <w:commentReference>),
// and XML structure (hyperlinks, comment ranges, bookmarks) are preserved.
// Replacement text inherits the formatting of the first affected run.
//
// Returns the number of replacements performed.
// Returns 0 if old == new (no-op optimization — XML is not modified even
// though occurrences may exist in the text).
func (p *CT_P) ReplaceText(old, new string) int {
	if old == "" || old == new {
		return 0
	}
	atoms, fullText := collectTextAtoms(p.e)
	if len(atoms) == 0 {
		return 0
	}
	return applyReplacements(atoms, fullText, old, new)
}
