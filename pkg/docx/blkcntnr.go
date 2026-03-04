package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// InnerContentItem represents either a *Paragraph or a *Table found in a
// block-item container. Callers inspect the type via Paragraph() / Table().
type InnerContentItem struct {
	paragraph *Paragraph
	table     *Table
}

// IsParagraph returns true if this item is a paragraph.
func (it *InnerContentItem) IsParagraph() bool { return it.paragraph != nil }

// IsTable returns true if this item is a table.
func (it *InnerContentItem) IsTable() bool { return it.table != nil }

// Paragraph returns the paragraph, or nil if this item is a table.
func (it *InnerContentItem) Paragraph() *Paragraph { return it.paragraph }

// Table returns the table, or nil if this item is a paragraph.
func (it *InnerContentItem) Table() *Table { return it.table }

// BlockItemContainer is the base for proxy objects that can contain block items
// (paragraphs and tables). These include Body, Cell, Header, Footer, and Comment.
//
// Mirrors Python BlockItemContainer(StoryChild).
type BlockItemContainer struct {
	element *etree.Element // CT_Body | CT_Comment | CT_HdrFtr | CT_Tc
	part    *parts.StoryPart
}

// newBlockItemContainer creates a new BlockItemContainer.
func newBlockItemContainer(element *etree.Element, part *parts.StoryPart) BlockItemContainer {
	return BlockItemContainer{element: element, part: part}
}

// AddParagraph appends a new paragraph to the end of this container. If text is
// non-empty, it is placed in a single run. If style is non-nil (string name),
// the paragraph style is applied.
//
// Mirrors Python BlockItemContainer.add_paragraph.
func (c *BlockItemContainer) AddParagraph(text string, style ...StyleRef) (*Paragraph, error) {
	p := c.addP()
	para := newParagraph(p, c.part)
	if text != "" {
		if _, err := para.AddRun(text); err != nil {
			return nil, fmt.Errorf("docx: adding run to paragraph: %w", err)
		}
	}
	if raw := resolveStyleRef(style); raw != nil {
		if err := para.setStyleRaw(raw); err != nil {
			return nil, fmt.Errorf("docx: setting paragraph style: %w", err)
		}
	}
	return para, nil
}

// AddTable appends a new table with the given rows, columns, and width (twips).
// The table is inserted before any trailing w:sectPr to maintain schema order.
//
// Mirrors Python BlockItemContainer.add_table (_insert_tbl with successor w:sectPr).
func (c *BlockItemContainer) AddTable(rows, cols int, widthTwips int) (*Table, error) {
	tbl := oxml.NewTbl(rows, cols, widthTwips)
	c.insertBeforeSectPr(tbl.RawElement())
	return newTable(tbl, c.part), nil
}

// IterInnerContent returns a slice of InnerContentItems (Paragraph or Table)
// in document order.
//
// Mirrors Python BlockItemContainer.iter_inner_content.
func (c *BlockItemContainer) IterInnerContent() []*InnerContentItem {
	var result []*InnerContentItem
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			p := &oxml.CT_P{Element: oxml.WrapElement(child)}
			result = append(result, &InnerContentItem{paragraph: newParagraph(p, c.part)})
		} else if child.Space == "w" && child.Tag == "tbl" {
			tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(child)}
			result = append(result, &InnerContentItem{table: newTable(tbl, c.part)})
		}
	}
	return result
}

// Paragraphs returns all paragraphs in this container, in document order.
//
// Mirrors Python BlockItemContainer.paragraphs.
func (c *BlockItemContainer) Paragraphs() []*Paragraph {
	var result []*Paragraph
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			p := &oxml.CT_P{Element: oxml.WrapElement(child)}
			result = append(result, newParagraph(p, c.part))
		}
	}
	return result
}

// Tables returns all tables in this container, in document order.
//
// Mirrors Python BlockItemContainer.tables.
func (c *BlockItemContainer) Tables() []*Table {
	var result []*Table
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "tbl" {
			tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(child)}
			result = append(result, newTable(tbl, c.part))
		}
	}
	return result
}

// ReplaceText replaces all occurrences of old with new in all paragraphs
// and tables of this container, recursively. Returns the total number of
// replacements performed.
func (c *BlockItemContainer) ReplaceText(old, new string) int {
	count := 0
	for _, item := range c.IterInnerContent() {
		if item.IsParagraph() {
			count += item.Paragraph().ReplaceText(old, new)
		} else if item.IsTable() {
			count += item.Table().ReplaceText(old, new)
		}
	}
	return count
}

// Element returns the backing etree element.
func (c *BlockItemContainer) Element() *etree.Element { return c.element }

// Part returns the story part this container belongs to.
func (c *BlockItemContainer) Part() *parts.StoryPart { return c.part }

// addP creates and inserts a new <w:p> element before any trailing w:sectPr.
func (c *BlockItemContainer) addP() *oxml.CT_P {
	pE := etree.NewElement("p")
	pE.Space = "w"
	c.insertBeforeSectPr(pE)
	return &oxml.CT_P{Element: oxml.WrapElement(pE)}
}

// insertBeforeSectPr inserts child into this container, placing it just before
// the first direct child <w:sectPr> if one exists. If no w:sectPr child is
// present, the element is appended to the end. This matches the Python
// xmlchemy successor constraint: both w:p and w:tbl have successors=('w:sectPr',).
//
// Only CT_Body has a trailing w:sectPr; for CT_Tc, CT_HdrFtr, and CT_Comment
// there is no such child, so this degrades to a simple append.
func (c *BlockItemContainer) insertBeforeSectPr(child *etree.Element) {
	children := c.element.Child
	for i, tok := range children {
		if el, ok := tok.(*etree.Element); ok {
			if el.Tag == "sectPr" && el.Space == "w" {
				// Remove child from any existing parent
				if p := child.Parent(); p != nil {
					p.RemoveChild(child)
				}
				c.element.InsertChildAt(i, child)
				return
			}
		}
	}
	// No w:sectPr — append normally
	c.element.AddChild(child)
}

// ---------------------------------------------------------------------------
// Phase 4: replaceTagWithElements engine + replaceWithTable wrapper
//
// The engine finds text tags in paragraphs, splits them using
// oxml.SplitParagraphAtTags, replaces placeholders with block elements
// built by a caller-supplied callback, and recurses into table cells.
//
// The only point of divergence between replacement types (table, content,
// image) is WHAT replaces the placeholder — isolated in the elementBuilder
// callback. Everything else (iteration, recursion, splice, cell invariant,
// width recalculation) is written exactly once.
// ---------------------------------------------------------------------------

// elementBuilder creates block elements to insert at a placeholder position.
// widthTwips is the available content width of the current container.
// Must return fresh elements on each call (no reuse).
type elementBuilder func(widthTwips int) ([]*etree.Element, error)

// splitWork records one paragraph that needs splitting.
type splitWork struct {
	pEl       *etree.Element
	fragments []oxml.Fragment
}

// replaceTagWithElements is the general engine for replacing text tags with
// block elements. widthTwips is the available content width of the current
// container (for passing to buildFn). On recursion into table cells the
// width is recalculated from tcPr.
//
// Step 1 processes paragraphs (split + splice). Step 2 recurses into tables.
// Steps are separated because Step 1 only touches <w:p> elements and Step 2
// only touches <w:tbl> elements — no overlap.
func (c *BlockItemContainer) replaceTagWithElements(
	old string,
	buildFn elementBuilder,
	widthTwips int,
) (int, error) {
	// ------------------------------------------------------------------
	// Step 1: process paragraphs (splice).
	// ------------------------------------------------------------------

	// Collect all paragraphs that contain the tag.
	var work []splitWork
	// Collect pre-existing tables BEFORE Step 1: Step 1 may insert new
	// <w:tbl> elements (via buildFn), but those fresh tables cannot contain
	// the tag and don't need recursion. Skipping them avoids unnecessary
	// SplitParagraphAtTags calls on their cells.
	var existingTables []*etree.Element
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			frags := oxml.SplitParagraphAtTags(child, old)
			if frags != nil {
				work = append(work, splitWork{pEl: child, fragments: frags})
			}
		} else if child.Space == "w" && child.Tag == "tbl" {
			existingTables = append(existingTables, child)
		}
	}

	// Process in reverse order so that earlier insertions don't shift
	// the indices of later ones.
	count := 0
	for i := len(work) - 1; i >= 0; i-- {
		sw := work[i]
		seq, n, err := c.buildSpliceSequence(sw.fragments, buildFn, widthTwips)
		if err != nil {
			return count, err
		}
		c.spliceElement(sw.pEl, seq)
		count += n
	}

	// ------------------------------------------------------------------
	// Step 2: recurse into pre-existing tables.
	// ------------------------------------------------------------------
	// Only tables collected BEFORE Step 1 are visited. Tables inserted by
	// Step 1 (via buildFn) are fresh and cannot contain the tag.
	for _, child := range existingTables {
		tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(child)}
		for _, tc := range tbl.IterTcs() {
			cellWidth := c.cellWidthOrDefault(tc, widthTwips)
			bic := newBlockItemContainer(tc.RawElement(), c.part)
			n, err := bic.replaceTagWithElements(old, buildFn, cellWidth)
			if err != nil {
				return count, err
			}
			count += n
			c.ensureTcHasParagraph(tc)
		}
	}

	return count, nil
}

// buildSpliceSequence converts fragments into a flat element sequence.
// For each paragraph fragment — takes its Element().
// For each placeholder — calls buildFn(widthTwips) and appends the result.
// Returns the element sequence and the number of placeholders processed.
func (c *BlockItemContainer) buildSpliceSequence(
	frags []oxml.Fragment,
	buildFn elementBuilder,
	widthTwips int,
) ([]*etree.Element, int, error) {
	var result []*etree.Element
	placeholders := 0
	for _, frag := range frags {
		if frag.IsParagraph() {
			result = append(result, frag.Element())
		} else {
			els, err := buildFn(widthTwips)
			if err != nil {
				return nil, 0, err
			}
			result = append(result, els...)
			placeholders++
		}
	}
	return result, placeholders, nil
}

// spliceElement replaces origEl in the container with the given sequence of
// elements.
//
// IMPORTANT: Uses parent.Child (all tokens: *etree.Element, *etree.CharData,
// etc.), not ChildElements() (only *etree.Element). The index of origEl in
// Child ≠ index in ChildElements(). This matches the pattern used in
// insertBefore (element.go:245) and insertBeforeSectPr (blkcntnr.go:158).
func (c *BlockItemContainer) spliceElement(origEl *etree.Element, newEls []*etree.Element) {
	parent := c.element
	// Find index of origEl in parent.Child (includes CharData, Comments, etc.)
	origIdx := -1
	for i, tok := range parent.Child {
		if el, ok := tok.(*etree.Element); ok && el == origEl {
			origIdx = i
			break
		}
	}
	if origIdx < 0 {
		return // defensive: element not found
	}

	// Remove original
	parent.RemoveChild(origEl)

	// Insert new elements at the same position, in order.
	// After RemoveChild, indices shift by -1, but we insert at origIdx+i
	// which is correct because InsertChildAt shifts subsequent children.
	for i, el := range newEls {
		if p := el.Parent(); p != nil {
			p.RemoveChild(el)
		}
		parent.InsertChildAt(origIdx+i, el)
	}
}

// cellWidthOrDefault returns the cell width in twips, falling back to
// parentWidth if not specified. Mirrors Cell.AddTable logic (table.go:274-278).
func (c *BlockItemContainer) cellWidthOrDefault(tc *oxml.CT_Tc, parentWidth int) int {
	if w, err := tc.WidthTwips(); err == nil && w != nil {
		return *w
	}
	return parentWidth
}

// ensureTcHasParagraph adds an empty <w:p> to a cell if it has no paragraph
// children. Required by OOXML schema (ISO/IEC 29500-1:2016 §17.4.66).
// Mirrors Cell.AddTable trailing paragraph (table.go:285).
func (c *BlockItemContainer) ensureTcHasParagraph(tc *oxml.CT_Tc) {
	for _, child := range tc.RawElement().ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			return // has at least one
		}
	}
	pE := oxml.OxmlElement("w:p")
	tc.RawElement().AddChild(pE)
}

// replaceWithTable replaces all occurrences of old with a table built from td.
// widthTwips is the available width for the table (page content width for body,
// cell width for nested tables, etc.).
//
// td must already be a defensive copy (done once in Document.ReplaceWithTable).
func (c *BlockItemContainer) replaceWithTable(old string, td TableData, widthTwips int) (int, error) {
	return c.replaceTagWithElements(old, func(w int) ([]*etree.Element, error) {
		el, err := buildTableElement(td, w)
		if err != nil {
			return nil, err
		}
		return []*etree.Element{el}, nil
	}, widthTwips)
}

// replaceWithContent replaces all occurrences of old with body elements from
// a source document. widthTwips is passed through to the engine for correct
// width recalculation when recursing into table cells, but the elementBuilder
// itself does not use it — content preserves its own dimensions.
//
// prepared must already contain remapped elements (produced by
// prepareContentElements). Each call to the elementBuilder returns fresh
// deep copies so that multiple placeholders get independent element trees.
func (c *BlockItemContainer) replaceWithContent(old string, prepared *preparedContent, widthTwips int) (int, error) {
	return c.replaceTagWithElements(old, func(w int) ([]*etree.Element, error) {
		// Fresh deep copy on each call — multiple placeholders get
		// independent element trees.
		cloned := make([]*etree.Element, len(prepared.elements))
		for i, el := range prepared.elements {
			cloned[i] = el.Copy()
		}
		// Renumber bare `id` attributes (wp:docPr, pic:cNvPr, etc.) so
		// that drawing shape ids are unique within this story part.
		renumberDrawingIDs(cloned, c.part.NextID)
		return cloned, nil
	}, widthTwips)
}
