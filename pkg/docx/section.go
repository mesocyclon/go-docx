package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Section provides access to section and page setup settings.
//
// Mirrors Python Section.
type Section struct {
	sectPr  *oxml.CT_SectPr
	docPart *parts.DocumentPart
}

// newSection creates a new Section proxy.
func newSection(sectPr *oxml.CT_SectPr, docPart *parts.DocumentPart) *Section {
	return &Section{sectPr: sectPr, docPart: docPart}
}

// BottomMargin returns the bottom margin in twips, or nil if not set.
func (s *Section) BottomMargin() (*int, error) { return s.sectPr.BottomMargin() }

// SetBottomMargin sets the bottom margin in twips.
func (s *Section) SetBottomMargin(v *int) error { return s.sectPr.SetBottomMargin(v) }

// TopMargin returns the top margin in twips, or nil if not set.
func (s *Section) TopMargin() (*int, error) { return s.sectPr.TopMargin() }

// SetTopMargin sets the top margin in twips.
func (s *Section) SetTopMargin(v *int) error { return s.sectPr.SetTopMargin(v) }

// LeftMargin returns the left margin in twips, or nil if not set.
func (s *Section) LeftMargin() (*int, error) { return s.sectPr.LeftMargin() }

// SetLeftMargin sets the left margin in twips.
func (s *Section) SetLeftMargin(v *int) error { return s.sectPr.SetLeftMargin(v) }

// RightMargin returns the right margin in twips, or nil if not set.
func (s *Section) RightMargin() (*int, error) { return s.sectPr.RightMargin() }

// SetRightMargin sets the right margin in twips.
func (s *Section) SetRightMargin(v *int) error { return s.sectPr.SetRightMargin(v) }

// PageWidth returns the page width in twips, or nil if not set.
func (s *Section) PageWidth() (*int, error) { return s.sectPr.PageWidth() }

// SetPageWidth sets the page width in twips.
func (s *Section) SetPageWidth(v *int) error { return s.sectPr.SetPageWidth(v) }

// PageHeight returns the page height in twips, or nil if not set.
func (s *Section) PageHeight() (*int, error) { return s.sectPr.PageHeight() }

// SetPageHeight sets the page height in twips.
func (s *Section) SetPageHeight(v *int) error { return s.sectPr.SetPageHeight(v) }

// Orientation returns the page orientation.
func (s *Section) Orientation() (enum.WdOrientation, error) { return s.sectPr.Orientation() }

// SetOrientation sets the page orientation.
func (s *Section) SetOrientation(v enum.WdOrientation) error { return s.sectPr.SetOrientation(v) }

// StartType returns the section start type.
func (s *Section) StartType() (enum.WdSectionStart, error) { return s.sectPr.StartType() }

// SetStartType sets the section start type.
func (s *Section) SetStartType(v enum.WdSectionStart) error { return s.sectPr.SetStartType(v) }

// Gutter returns the gutter in twips, or nil if not set.
func (s *Section) Gutter() (*int, error) { return s.sectPr.GutterMargin() }

// SetGutter sets the gutter in twips.
func (s *Section) SetGutter(v *int) error { return s.sectPr.SetGutterMargin(v) }

// HeaderDistance returns the header distance in twips, or nil if not set.
func (s *Section) HeaderDistance() (*int, error) { return s.sectPr.HeaderMargin() }

// SetHeaderDistance sets the header distance.
func (s *Section) SetHeaderDistance(v *int) error { return s.sectPr.SetHeaderMargin(v) }

// FooterDistance returns the footer distance in twips, or nil if not set.
func (s *Section) FooterDistance() (*int, error) { return s.sectPr.FooterMargin() }

// SetFooterDistance sets the footer distance.
func (s *Section) SetFooterDistance(v *int) error { return s.sectPr.SetFooterMargin(v) }

// DifferentFirstPageHeaderFooter returns true if this section displays a distinct
// first-page header and footer.
func (s *Section) DifferentFirstPageHeaderFooter() bool { return s.sectPr.TitlePgVal() }

// SetDifferentFirstPageHeaderFooter sets the first-page header/footer flag.
func (s *Section) SetDifferentFirstPageHeaderFooter(v bool) error {
	return s.sectPr.SetTitlePgVal(v)
}

// Header returns the default (primary) page header.
func (s *Section) Header() *Header {
	return newHeader(s.sectPr, s.docPart, enum.WdHeaderFooterIndexPrimary)
}

// Footer returns the default (primary) page footer.
func (s *Section) Footer() *Footer {
	return newFooter(s.sectPr, s.docPart, enum.WdHeaderFooterIndexPrimary)
}

// EvenPageHeader returns the even-page header.
func (s *Section) EvenPageHeader() *Header {
	return newHeader(s.sectPr, s.docPart, enum.WdHeaderFooterIndexEvenPage)
}

// EvenPageFooter returns the even-page footer.
func (s *Section) EvenPageFooter() *Footer {
	return newFooter(s.sectPr, s.docPart, enum.WdHeaderFooterIndexEvenPage)
}

// FirstPageHeader returns the first-page header.
func (s *Section) FirstPageHeader() *Header {
	return newHeader(s.sectPr, s.docPart, enum.WdHeaderFooterIndexFirstPage)
}

// FirstPageFooter returns the first-page footer.
func (s *Section) FirstPageFooter() *Footer {
	return newFooter(s.sectPr, s.docPart, enum.WdHeaderFooterIndexFirstPage)
}

// IterInnerContent returns paragraphs and tables in this section body.
//
// Mirrors Python Section.iter_inner_content → CT_SectPr.iter_inner_content →
// _SectBlockElementIterator. Only returns block-items belonging to THIS section,
// not the entire document body.
func (s *Section) IterInnerContent() []*InnerContentItem {
	body := s.sectPr.BodyElement()
	if body == nil {
		return nil
	}

	// Determine section boundaries. Sections are delimited by sectPr elements:
	// - Paragraph-based: w:p/w:pPr/w:sectPr marks end of a section (the p itself belongs to that section)
	// - Body-based: w:body/w:sectPr marks the last section
	//
	// We walk body children once, collecting (start, end) ranges for each section.

	type sectionRange struct {
		startIdx int // inclusive index into body children
		endIdx   int // exclusive index into body children
		sectPrEl *etree.Element
	}

	children := body.ChildElements()
	var ranges []sectionRange
	rangeStart := 0

	for i, child := range children {
		if child.Space == "w" && child.Tag == "p" {
			// Check if this paragraph contains w:pPr/w:sectPr
			if pSectPr := findParagraphSectPr(child); pSectPr != nil {
				// This paragraph (inclusive) ends a section
				ranges = append(ranges, sectionRange{
					startIdx: rangeStart,
					endIdx:   i + 1, // include this p
					sectPrEl: pSectPr,
				})
				rangeStart = i + 1
			}
		} else if child.Space == "w" && child.Tag == "sectPr" {
			// Body-level sectPr: last section
			ranges = append(ranges, sectionRange{
				startIdx: rangeStart,
				endIdx:   i, // exclude the sectPr itself
				sectPrEl: child,
			})
		}
	}

	// Find which range matches our sectPr
	for _, sr := range ranges {
		if sr.sectPrEl == s.sectPr.RawElement() {
			return collectBlockItems(children[sr.startIdx:sr.endIdx], s.docPart)
		}
	}

	return nil
}

// findParagraphSectPr returns the w:sectPr element inside w:p/w:pPr, or nil.
func findParagraphSectPr(p *etree.Element) *etree.Element {
	for _, child := range p.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			for _, gc := range child.ChildElements() {
				if gc.Space == "w" && gc.Tag == "sectPr" {
					return gc
				}
			}
		}
	}
	return nil
}

// collectBlockItems filters elements for w:p and w:tbl, wrapping them as
// InnerContentItems with the appropriate proxy objects.
func collectBlockItems(elems []*etree.Element, docPart *parts.DocumentPart) []*InnerContentItem {
	var result []*InnerContentItem
	var sp *parts.StoryPart
	if docPart != nil {
		sp = &docPart.StoryPart
	}
	for _, child := range elems {
		switch {
		case child.Space == "w" && child.Tag == "p":
			p := &oxml.CT_P{Element: oxml.WrapElement(child)}
			result = append(result, &InnerContentItem{paragraph: newParagraph(p, sp)})
		case child.Space == "w" && child.Tag == "tbl":
			tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(child)}
			result = append(result, &InnerContentItem{table: newTable(tbl, sp)})
		}
	}
	return result
}

// --------------------------------------------------------------------------
// Sections
// --------------------------------------------------------------------------

// Sections is a sequence of Section objects corresponding to the sections in a document.
//
// Mirrors Python Sections(Sequence).
type Sections struct {
	docElm  *oxml.CT_Document
	docPart *parts.DocumentPart
}

// newSections creates a new Sections proxy.
func newSections(docElm *oxml.CT_Document, docPart *parts.DocumentPart) *Sections {
	return &Sections{docElm: docElm, docPart: docPart}
}

// Len returns the number of sections.
func (ss *Sections) Len() int {
	return len(ss.docElm.SectPrList())
}

// Get returns the section at the given index.
func (ss *Sections) Get(idx int) (*Section, error) {
	lst := ss.docElm.SectPrList()
	if idx < 0 || idx >= len(lst) {
		return nil, fmt.Errorf("docx: section index [%d] out of range", idx)
	}
	return newSection(lst[idx], ss.docPart), nil
}

// Iter returns all sections in document order.
func (ss *Sections) Iter() []*Section {
	lst := ss.docElm.SectPrList()
	result := make([]*Section, len(lst))
	for i, sp := range lst {
		result[i] = newSection(sp, ss.docPart)
	}
	return result
}

// --------------------------------------------------------------------------
// Header / Footer — _BaseHeaderFooter pattern
//
// Mirrors Python _BaseHeaderFooter(BlockItemContainer). The shared logic
// (IsLinkedToPrevious, AddParagraph, Part, getOrAddDefinition, etc.) lives
// in baseHeaderFooter. Type-specific operations (hasDefinition, definition,
// addDefinition, dropDefinition, prior) are provided via the hdrFtrOps
// interface, implemented separately by Header and Footer.
// --------------------------------------------------------------------------

// hdrFtrOps encapsulates the five type-specific hook methods that differ
// between Header and Footer. Everything else is shared in baseHeaderFooter.
type hdrFtrOps interface {
	// hasDefinition reports whether an explicit definition exists for this
	// header/footer in its sectPr (i.e. a headerReference / footerReference).
	hasDefinition() (bool, error)

	// definition returns the StoryPart that contains the content.
	definition() (*parts.StoryPart, error)

	// addDefinition creates a new part and wires the reference in sectPr.
	addDefinition() (*parts.StoryPart, error)

	// dropDefinition removes the reference and drops the related part.
	dropDefinition() error

	// prior returns the same-type header/footer ops for the preceding section,
	// or nil if this is the first section.
	prior() hdrFtrOps

	// kind returns "header" or "footer" for error messages.
	kind() string
}

// baseHeaderFooter holds all the shared logic for Header and Footer.
//
// Mirrors Python _BaseHeaderFooter: is_linked_to_previous, part,
// _get_or_add_definition, _element, and all BlockItemContainer delegators.
type baseHeaderFooter struct {
	ops hdrFtrOps
}

// IsLinkedToPrevious reports whether this header/footer uses the definition
// from the prior section. Returns false on error (conservative: assume own
// definition).
//
// Mirrors Python _BaseHeaderFooter.is_linked_to_previous.
func (b *baseHeaderFooter) IsLinkedToPrevious() bool {
	has, err := b.ops.hasDefinition()
	if err != nil {
		return false
	}
	return !has
}

// SetIsLinkedToPrevious sets the linked-to-previous state.
func (b *baseHeaderFooter) SetIsLinkedToPrevious(v bool) error {
	if v == b.IsLinkedToPrevious() {
		return nil
	}
	if v {
		return b.ops.dropDefinition()
	}
	_, err := b.ops.addDefinition()
	return err
}

// AddParagraph appends a new paragraph to this header/footer.
//
// Mirrors Python BlockItemContainer.add_paragraph (inherited by _BaseHeaderFooter).
func (b *baseHeaderFooter) AddParagraph(text string, style ...StyleRef) (*Paragraph, error) {
	bic, err := b.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: %s add paragraph: %w", b.ops.kind(), err)
	}
	return bic.AddParagraph(text, style...)
}

// AddTable appends a new table to this header/footer.
//
// Mirrors Python BlockItemContainer.add_table (inherited by _BaseHeaderFooter).
func (b *baseHeaderFooter) AddTable(rows, cols int, widthTwips int) (*Table, error) {
	bic, err := b.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: %s add table: %w", b.ops.kind(), err)
	}
	return bic.AddTable(rows, cols, widthTwips)
}

// Paragraphs returns the paragraphs in this header/footer.
//
// Mirrors Python BlockItemContainer.paragraphs (inherited by _BaseHeaderFooter).
func (b *baseHeaderFooter) Paragraphs() ([]*Paragraph, error) {
	bic, err := b.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: %s paragraphs: %w", b.ops.kind(), err)
	}
	return bic.Paragraphs(), nil
}

// Tables returns the tables in this header/footer.
//
// Mirrors Python BlockItemContainer.tables (inherited by _BaseHeaderFooter).
func (b *baseHeaderFooter) Tables() ([]*Table, error) {
	bic, err := b.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: %s tables: %w", b.ops.kind(), err)
	}
	return bic.Tables(), nil
}

// IterInnerContent returns paragraphs and tables in document order.
//
// Mirrors Python BlockItemContainer.iter_inner_content (inherited by _BaseHeaderFooter).
func (b *baseHeaderFooter) IterInnerContent() ([]*InnerContentItem, error) {
	bic, err := b.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: %s inner content: %w", b.ops.kind(), err)
	}
	return bic.IterInnerContent(), nil
}

// Part returns the underlying StoryPart for style resolution and image
// insertion.
//
// Mirrors Python _BaseHeaderFooter.part property.
func (b *baseHeaderFooter) Part() (*parts.StoryPart, error) {
	sp, err := b.getOrAddDefinition()
	if err != nil {
		return nil, fmt.Errorf("docx: resolving %s part: %w", b.ops.kind(), err)
	}
	return sp, nil
}

// replaceText performs text replacement in this header/footer's content.
// If the header/footer has no own definition (linked to previous), returns
// 0, nil — the content will be processed through the owning section.
func (b *baseHeaderFooter) replaceText(old, new string) (int, error) {
	has, err := b.ops.hasDefinition()
	if err != nil {
		return 0, err
	}
	if !has {
		return 0, nil
	}
	bic, err := b.blockItemContainer()
	if err != nil {
		return 0, err
	}
	return bic.ReplaceText(old, new), nil
}

// replaceTextDedup is a variant with StoryPart deduplication.
// Used by Document.ReplaceText to prevent double replacement when
// multiple sections share the same HeaderPart/FooterPart.
//
// Uses hasDefinition() + definition() (which returns an existing
// StoryPart without creating one) instead of Part() (which calls
// getOrAddDefinition() and may create an empty definition as a
// side effect).
func (b *baseHeaderFooter) replaceTextDedup(old, new string, seen map[*parts.StoryPart]bool) (int, error) {
	return b.applyToContainerDedup(func(bic *BlockItemContainer) (int, error) {
		return bic.ReplaceText(old, new), nil
	}, seen)
}

// applyToContainerDedup executes fn on this header/footer's BlockItemContainer
// with StoryPart deduplication. Skips linked-to-previous headers/footers and
// already-processed parts. Used by Document.ReplaceText and
// Document.ReplaceWithTable.
//
// Deduplication logic: hasDefinition → definition → dedup check →
// blockItemContainer → fn. Identical to the original replaceTextDedup before
// it was refactored into a delegate.
func (b *baseHeaderFooter) applyToContainerDedup(
	fn func(*BlockItemContainer) (int, error),
	seen map[*parts.StoryPart]bool,
) (int, error) {
	has, err := b.ops.hasDefinition()
	if err != nil {
		return 0, err
	}
	if !has {
		return 0, nil
	}
	sp, err := b.ops.definition()
	if err != nil {
		return 0, err
	}
	if sp == nil {
		return 0, nil
	}
	if seen[sp] {
		return 0, nil
	}
	seen[sp] = true
	bic, err := b.blockItemContainer()
	if err != nil {
		return 0, err
	}
	return fn(bic)
}

// replaceWithTableDedup replaces text tags with tables in this header/footer
// with StoryPart deduplication. Used by Document.ReplaceWithTable.
func (b *baseHeaderFooter) replaceWithTableDedup(
	old string, td TableData, widthTwips int,
	seen map[*parts.StoryPart]bool,
) (int, error) {
	return b.applyToContainerDedup(func(bic *BlockItemContainer) (int, error) {
		return bic.replaceWithTable(old, td, widthTwips)
	}, seen)
}

// replaceWithContentDedup replaces text tags with source document content in
// this header/footer with StoryPart deduplication. Used by
// Document.ReplaceWithContent.
//
// Unlike replaceWithTableDedup, this method calls prepareContentElements per
// unique StoryPart because each header/footer part has its own relationships.
// The importedParts map is shared across all containers so that part blobs
// are copied into the target package only once.
func (b *baseHeaderFooter) replaceWithContentDedup(
	old string,
	sourceDoc *Document,
	ri *ResourceImporter,
	widthTwips int,
	seen map[*parts.StoryPart]bool,
) (int, error) {
	return b.applyToContainerDedup(func(bic *BlockItemContainer) (int, error) {
		// Prepare content with relationships mapped to THIS header's StoryPart.
		prep, err := prepareContentElements(sourceDoc, bic.part, ri)
		if err != nil {
			return 0, err
		}
		return bic.replaceWithContent(old, prep, widthTwips)
	}, seen)
}

// blockItemContainer creates a BlockItemContainer backed by the header/footer
// part's element and StoryPart. Created fresh each call to match Python's
// property behavior (no stale cache if definition changes).
func (b *baseHeaderFooter) blockItemContainer() (*BlockItemContainer, error) {
	sp, err := b.getOrAddDefinition()
	if err != nil {
		return nil, fmt.Errorf("docx: resolving %s definition: %w", b.ops.kind(), err)
	}
	el := sp.Element()
	if el == nil {
		return nil, fmt.Errorf("docx: %s part has nil element", b.ops.kind())
	}
	bic := newBlockItemContainer(el, sp)
	return &bic, nil
}

// getOrAddDefinition mirrors Python _BaseHeaderFooter._get_or_add_definition.
// Walks backward through preceding sections looking for an existing definition.
// If no section in the chain has one, adds a new definition on the earliest
// (first) section — matching the original recursive semantics without unbounded
// call depth.
func (b *baseHeaderFooter) getOrAddDefinition() (*parts.StoryPart, error) {
	cur := b.ops
	for {
		has, err := cur.hasDefinition()
		if err != nil {
			return nil, err
		}
		if has {
			return cur.definition()
		}
		p := cur.prior()
		if p == nil {
			// First section reached — create a new definition here.
			return cur.addDefinition()
		}
		cur = p
	}
}

// --------------------------------------------------------------------------
// Header
// --------------------------------------------------------------------------

// Header is a proxy for a page header.
//
// Mirrors Python _Header(_BaseHeaderFooter(BlockItemContainer)).
// Provides BlockItemContainer methods (AddParagraph, AddTable, Paragraphs,
// Tables, IterInnerContent) via the embedded baseHeaderFooter; type-specific
// operations are implemented directly on Header to satisfy hdrFtrOps.
type Header struct {
	baseHeaderFooter
	sectPr  *oxml.CT_SectPr
	docPart *parts.DocumentPart
	index   enum.WdHeaderFooterIndex
}

// newHeader creates a new Header proxy.
func newHeader(sectPr *oxml.CT_SectPr, docPart *parts.DocumentPart, index enum.WdHeaderFooterIndex) *Header {
	h := &Header{sectPr: sectPr, docPart: docPart, index: index}
	h.ops = h
	return h
}

func (h *Header) kind() string { return "header" }

// ReplaceText replaces all occurrences of old with new in this header's
// content. If the header is linked to previous (no own definition),
// returns 0, nil.
func (h *Header) ReplaceText(old, new string) (int, error) {
	return h.replaceText(old, new)
}

func (h *Header) hasDefinition() (bool, error) {
	ref, err := h.sectPr.GetHeaderRef(h.index)
	if err != nil {
		return false, fmt.Errorf("docx: checking header ref for index %d: %w", h.index, err)
	}
	return ref != nil, nil
}

func (h *Header) definition() (*parts.StoryPart, error) {
	ref, err := h.sectPr.GetHeaderRef(h.index)
	if err != nil {
		return nil, fmt.Errorf("docx: getting header ref for index %d: %w", h.index, err)
	}
	if ref == nil {
		return nil, nil
	}
	rId, err := ref.RId()
	if err != nil {
		return nil, fmt.Errorf("docx: getting header rId: %w", err)
	}
	hp, err := h.docPart.HeaderPartByRID(rId)
	if err != nil {
		return nil, fmt.Errorf("docx: resolving header part for rId %q: %w", rId, err)
	}
	return &hp.StoryPart, nil
}

func (h *Header) addDefinition() (*parts.StoryPart, error) {
	hp, rId, err := h.docPart.AddHeaderPart()
	if err != nil {
		return nil, err
	}
	if _, err := h.sectPr.AddHeaderRef(h.index, rId); err != nil {
		return nil, err
	}
	return &hp.StoryPart, nil
}

func (h *Header) dropDefinition() error {
	rId, err := h.sectPr.RemoveHeaderRef(h.index)
	if err != nil {
		return fmt.Errorf("docx: removing header ref: %w", err)
	}
	if rId != "" {
		h.docPart.DropHeaderPart(rId)
	}
	return nil
}

func (h *Header) prior() hdrFtrOps {
	prev := h.sectPr.PrecedingSectPr()
	if prev == nil {
		return nil
	}
	return newHeader(prev, h.docPart, h.index)
}

// --------------------------------------------------------------------------
// Footer
// --------------------------------------------------------------------------

// Footer is a proxy for a page footer.
//
// Mirrors Python _Footer(_BaseHeaderFooter(BlockItemContainer)).
// Provides BlockItemContainer methods (AddParagraph, AddTable, Paragraphs,
// Tables, IterInnerContent) via the embedded baseHeaderFooter; type-specific
// operations are implemented directly on Footer to satisfy hdrFtrOps.
type Footer struct {
	baseHeaderFooter
	sectPr  *oxml.CT_SectPr
	docPart *parts.DocumentPart
	index   enum.WdHeaderFooterIndex
}

// newFooter creates a new Footer proxy.
func newFooter(sectPr *oxml.CT_SectPr, docPart *parts.DocumentPart, index enum.WdHeaderFooterIndex) *Footer {
	f := &Footer{sectPr: sectPr, docPart: docPart, index: index}
	f.ops = f
	return f
}

func (f *Footer) kind() string { return "footer" }

// ReplaceText replaces all occurrences of old with new in this footer's
// content. If the footer is linked to previous (no own definition),
// returns 0, nil.
func (f *Footer) ReplaceText(old, new string) (int, error) {
	return f.replaceText(old, new)
}

func (f *Footer) hasDefinition() (bool, error) {
	ref, err := f.sectPr.GetFooterRef(f.index)
	if err != nil {
		return false, fmt.Errorf("docx: checking footer ref for index %d: %w", f.index, err)
	}
	return ref != nil, nil
}

func (f *Footer) definition() (*parts.StoryPart, error) {
	ref, err := f.sectPr.GetFooterRef(f.index)
	if err != nil {
		return nil, fmt.Errorf("docx: getting footer ref for index %d: %w", f.index, err)
	}
	if ref == nil {
		return nil, nil
	}
	rId, err := ref.RId()
	if err != nil {
		return nil, fmt.Errorf("docx: getting footer rId: %w", err)
	}
	fp, err := f.docPart.FooterPartByRID(rId)
	if err != nil {
		return nil, fmt.Errorf("docx: resolving footer part for rId %q: %w", rId, err)
	}
	return &fp.StoryPart, nil
}

func (f *Footer) addDefinition() (*parts.StoryPart, error) {
	fp, rId, err := f.docPart.AddFooterPart()
	if err != nil {
		return nil, err
	}
	if _, err := f.sectPr.AddFooterRef(f.index, rId); err != nil {
		return nil, err
	}
	return &fp.StoryPart, nil
}

func (f *Footer) dropDefinition() error {
	rId, err := f.sectPr.RemoveFooterRef(f.index)
	if err != nil {
		return fmt.Errorf("docx: removing footer ref: %w", err)
	}
	if rId != "" {
		f.docPart.DropRel(rId)
	}
	return nil
}

func (f *Footer) prior() hdrFtrOps {
	prev := f.sectPr.PrecedingSectPr()
	if prev == nil {
		return nil
	}
	return newFooter(prev, f.docPart, f.index)
}
