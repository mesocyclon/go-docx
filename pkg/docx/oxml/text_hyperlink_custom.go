package oxml

import "strings"

// --- CT_Hyperlink custom methods ---

// HyperlinkText returns the textual content of this hyperlink by concatenating
// text from all child w:r elements.
func (h *CT_Hyperlink) HyperlinkText() string {
	var sb strings.Builder
	for _, child := range h.e.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			r := &CT_R{Element{e: child}}
			sb.WriteString(r.RunText())
		}
	}
	return sb.String()
}

// HyperlinkLastRenderedPageBreaks returns all w:lastRenderedPageBreak descendants
// inside runs of this hyperlink.
func (h *CT_Hyperlink) HyperlinkLastRenderedPageBreaks() []*CT_LastRenderedPageBreak {
	var result []*CT_LastRenderedPageBreak
	for _, child := range h.e.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			for _, gc := range child.ChildElements() {
				if gc.Space == "w" && gc.Tag == "lastRenderedPageBreak" {
					result = append(result, &CT_LastRenderedPageBreak{Element{e: gc}})
				}
			}
		}
	}
	return result
}
