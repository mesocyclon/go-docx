package oxml

import (
	"strconv"
	"strings"

	"github.com/beevik/etree"
)

// --- CT_R custom methods ---

// AddTWithText adds a <w:t> element containing the given text.
// Sets xml:space="preserve" if the text has leading or trailing whitespace.
func (r *CT_R) AddTWithText(text string) *CT_Text {
	t := r.addT()
	t.SetText(text)
	if len(strings.TrimSpace(text)) < len(text) {
		t.e.CreateAttr("xml:space", "preserve")
	}
	return t
}

// AddDrawingWithInline adds a <w:drawing> element containing the given inline element.
func (r *CT_R) AddDrawingWithInline(inline *CT_Inline) *CT_Drawing {
	drawing := r.addDrawing()
	drawing.e.AddChild(inline.e)
	return drawing
}

// ClearContent removes all child elements except <w:rPr>.
func (r *CT_R) ClearContent() {
	var toRemove []*etree.Element
	for _, child := range r.e.ChildElements() {
		if !(child.Space == "w" && child.Tag == "rPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		r.e.RemoveChild(child)
	}
}

// Style returns the styleId of the run, or nil if not set.
func (r *CT_R) Style() (*string, error) {
	rPr := r.RPr()
	if rPr == nil {
		return nil, nil
	}
	return rPr.StyleVal()
}

// SetStyle sets the run style. Passing nil removes the rStyle element.
func (r *CT_R) SetStyle(styleID *string) error {
	rPr := r.GetOrAddRPr()
	if err := rPr.SetStyleVal(styleID); err != nil {
		return err
	}
	return nil
}

// RunText returns the textual content of this run by concatenating text equivalents
// of all inner-content elements (w:t, w:br, w:cr, w:tab, w:noBreakHyphen, w:ptab).
func (r *CT_R) RunText() string {
	var sb strings.Builder
	for _, child := range r.e.ChildElements() {
		if child.Space != "w" {
			continue
		}
		switch child.Tag {
		case "t":
			sb.WriteString(child.Text())
		case "br":
			br := &CT_Br{Element{e: child}}
			sb.WriteString(br.TextEquivalent())
		case "cr":
			sb.WriteString("\n")
		case "tab":
			sb.WriteString("\t")
		case "noBreakHyphen":
			sb.WriteString("-")
		case "ptab":
			sb.WriteString("\t")
		}
	}
	return sb.String()
}

// SetRunText replaces all run content with elements representing the given text.
// Tab characters become <w:tab/>, newlines/carriage-returns become <w:br/>,
// and regular characters are grouped into <w:t> elements.
func (r *CT_R) SetRunText(text string) {
	r.ClearContent()
	appendRunContentFromText(r, text)
}

// LastRenderedPageBreaks returns all <w:lastRenderedPageBreak> descendants of this run.
func (r *CT_R) LastRenderedPageBreaks() []*CT_LastRenderedPageBreak {
	var result []*CT_LastRenderedPageBreak
	for _, child := range r.e.ChildElements() {
		if child.Space == "w" && child.Tag == "lastRenderedPageBreak" {
			result = append(result, &CT_LastRenderedPageBreak{Element{e: child}})
		}
	}
	return result
}

// NewDetachedBr creates a <w:br> element that is NOT yet attached to this run's
// XML tree. Configure attributes (SetType, SetClear) on the returned element,
// then call AttachBr to insert it in the correct sequence position.
//
// This avoids orphan elements: if attribute configuration fails, the tree is
// left untouched.
func (r *CT_R) NewDetachedBr() *CT_Br {
	return r.newBr()
}

// AttachBr inserts a previously detached <w:br> into this run in correct
// sequence order. Typically called after NewDetachedBr + attribute setup.
func (r *CT_R) AttachBr(br *CT_Br) {
	r.insertBr(br)
}

// --- CT_Br custom methods ---

// TextEquivalent returns the text equivalent of this break element.
// Line breaks produce "\n"; column and page breaks produce "".
func (br *CT_Br) TextEquivalent() string {
	if br.Type() == "textWrapping" {
		return "\n"
	}
	return ""
}

// --- CT_Cr custom methods ---

// TextEquivalent returns the text equivalent of a carriage return element: "\n".
func (cr *CT_Cr) TextEquivalent() string {
	return "\n"
}

// --- CT_NoBreakHyphen custom methods ---

// TextEquivalent returns the text equivalent of a non-breaking hyphen: "-".
func (nbh *CT_NoBreakHyphen) TextEquivalent() string {
	return "-"
}

// --- CT_PTab custom methods ---

// TextEquivalent returns the text equivalent of an absolute-position tab: "\t".
func (pt *CT_PTab) TextEquivalent() string {
	return "\t"
}

// --- CT_Text custom methods ---

// ContentText returns the text content of this <w:t> element, or empty string if none.
func (t *CT_Text) ContentText() string {
	return t.e.Text()
}

// SetPreserveSpace sets xml:space="preserve" on this <w:t> element.
func (t *CT_Text) SetPreserveSpace() {
	t.e.CreateAttr("xml:space", "preserve")
}

// --- Run content appender utility ---

// RunInnerContentItem represents one item from a run's inner content:
// either a string (accumulated text), *CT_Drawing, or *CT_LastRenderedPageBreak.
type RunInnerContentItem = interface{}

// InnerContentItems returns the inner content items of this run in document order.
// Text-like elements (w:t, w:br, w:cr, w:tab, w:noBreakHyphen, w:ptab) are
// accumulated into contiguous strings. Drawing and LastRenderedPageBreak elements
// are yielded individually, interrupting any accumulated text.
//
// Mirrors Python CT_R.inner_content_items using a TextAccumulator pattern.
func (r *CT_R) InnerContentItems() []RunInnerContentItem {
	var result []RunInnerContentItem
	var textBuf strings.Builder

	flushText := func() {
		if textBuf.Len() > 0 {
			result = append(result, textBuf.String())
			textBuf.Reset()
		}
	}

	for _, child := range r.e.ChildElements() {
		if child.Space != "w" {
			continue
		}
		switch child.Tag {
		case "drawing":
			flushText()
			result = append(result, &CT_Drawing{Element{e: child}})
		case "lastRenderedPageBreak":
			flushText()
			result = append(result, &CT_LastRenderedPageBreak{Element{e: child}})
		case "t":
			textBuf.WriteString(child.Text())
		case "br":
			br := &CT_Br{Element{e: child}}
			textBuf.WriteString(br.TextEquivalent())
		case "cr":
			textBuf.WriteString("\n")
		case "tab":
			textBuf.WriteString("\t")
		case "noBreakHyphen":
			textBuf.WriteString("-")
		case "ptab":
			textBuf.WriteString("\t")
		}
	}
	flushText()
	return result
}

// appendRunContentFromText translates a string into run content elements.
// Tabs → <w:tab/>, newlines → <w:br/>, regular chars → <w:t>.
func appendRunContentFromText(r *CT_R, text string) {
	var buf strings.Builder
	flush := func() {
		if buf.Len() > 0 {
			r.AddTWithText(buf.String())
			buf.Reset()
		}
	}
	for _, ch := range text {
		switch ch {
		case '\t':
			flush()
			r.AddTab()
		case '\n', '\r':
			flush()
			r.AddBr()
		default:
			buf.WriteRune(ch)
		}
	}
	flush()
}

// InsertCommentRangeStartAbove inserts a <w:commentRangeStart> element
// with the given commentID immediately before this run in its parent.
//
// Mirrors Python CT_R.insert_comment_range_start_above.
func (r *CT_R) InsertCommentRangeStartAbove(commentID int) {
	parent := r.e.Parent()
	if parent == nil {
		return
	}
	crs := parent.CreateElement("w:commentRangeStart")
	crs.Space = "w"
	crs.Tag = "commentRangeStart"
	crs.CreateAttr("w:id", strconv.Itoa(commentID))
	// Move before this run: find run index, remove crs from end, insert at run index
	idx := childIndex(parent, r.e)
	parent.RemoveChild(crs)
	parent.InsertChildAt(idx, crs)
}

// InsertCommentRangeEndAndReferenceBelow inserts a <w:commentRangeEnd> and
// a <w:r> with <w:commentReference> immediately after this run in its parent.
//
// Produces:
//
//	<w:commentRangeEnd w:id="N"/>
//	<w:r>
//	  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
//	  <w:commentReference w:id="N"/>
//	</w:r>
//
// Mirrors Python CT_R.insert_comment_range_end_and_reference_below.
func (r *CT_R) InsertCommentRangeEndAndReferenceBelow(commentID int) {
	parent := r.e.Parent()
	if parent == nil {
		return
	}
	idStr := strconv.Itoa(commentID)
	idx := childIndex(parent, r.e)

	// Build commentRangeEnd
	cre := parent.CreateElement("w:commentRangeEnd")
	cre.Space = "w"
	cre.Tag = "commentRangeEnd"
	cre.CreateAttr("w:id", idStr)
	// Move to just after this run
	parent.RemoveChild(cre)
	parent.InsertChildAt(idx+1, cre)

	// Build reference run: <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="N"/></w:r>
	refRun := parent.CreateElement("w:r")
	refRun.Space = "w"
	refRun.Tag = "r"
	rPr := refRun.CreateElement("w:rPr")
	rPr.Space = "w"
	rPr.Tag = "rPr"
	rStyle := rPr.CreateElement("w:rStyle")
	rStyle.Space = "w"
	rStyle.Tag = "rStyle"
	rStyle.CreateAttr("w:val", "CommentReference")
	cr := refRun.CreateElement("w:commentReference")
	cr.Space = "w"
	cr.Tag = "commentReference"
	cr.CreateAttr("w:id", idStr)
	// Move to just after commentRangeEnd
	parent.RemoveChild(refRun)
	parent.InsertChildAt(idx+2, refRun)
}

// childIndex returns the index of child in parent's children, or -1.
func childIndex(parent, child *etree.Element) int {
	for i, c := range parent.Child {
		if el, ok := c.(*etree.Element); ok && el == child {
			return i
		}
	}
	return -1
}
