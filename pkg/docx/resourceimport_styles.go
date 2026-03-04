package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// resourceimport_styles.go — Style import for ResourceImporter
//
// Scans the source document body for pStyle, rStyle, tblStyle references,
// computes the transitive closure (basedOn, next, link chains), and merges
// each style into the target document using UseDestinationStyles strategy:
//
//   - Style exists in target → use target definition (no copy).
//   - Style missing from target → deep-copy from source.
//
// The styleMap is populated for later use by remapAll.
// --------------------------------------------------------------------------

// importStyles imports all styles referenced by the source document body
// into the target document. Computes the transitive closure of style
// dependencies (basedOn, next, link) and merges each style.
//
// Safe to call multiple times — only runs once (idempotent via styleDone flag).
func (ri *ResourceImporter) importStyles() error {
	if ri.styleDone {
		return nil
	}
	ri.styleDone = true

	// 0. Detect default paragraph style mismatch.
	// If source and target have different defaults, paragraphs without
	// explicit pStyle will silently change appearance. We record the
	// source default so materializeImplicitStyles can fix this later.
	if err := ri.detectDefaultStyleMismatch(); err != nil {
		return fmt.Errorf("docx: detecting default style mismatch: %w", err)
	}

	// 1. Collect all styleId values from source body.
	seedIds := collectStyleIdsFromBody(ri.sourceDoc)

	// If materialization is needed, the source default style must also
	// be in the seed set (it will be referenced by materialized pStyle attrs).
	if ri.srcDefaultParaStyleId != "" {
		seedIds = appendUnique(seedIds, ri.srcDefaultParaStyleId)
	}

	if len(seedIds) == 0 {
		return nil
	}

	// 2. Compute transitive closure over style dependencies.
	closure := ri.collectStyleClosure(seedIds)

	// 3. Merge each style into target.
	for _, srcStyle := range closure {
		if err := ri.mergeOneStyle(srcStyle); err != nil {
			return err
		}
	}
	return nil
}

// importStylesForElements imports styles referenced by arbitrary elements.
// Used by importFootnotes (Phase 4) to import styles from footnote bodies
// (e.g. FootnoteText) that are not present in the document body.
//
// Idempotent via styleMap — styles already imported are skipped.
func (ri *ResourceImporter) importStylesForElements(elements []*etree.Element) error {
	seedIds := collectStyleIdsFromElements(elements)
	if len(seedIds) == 0 {
		return nil
	}
	closure := ri.collectStyleClosure(seedIds)
	for _, srcStyle := range closure {
		if err := ri.mergeOneStyle(srcStyle); err != nil {
			return err
		}
	}
	return nil
}

// --------------------------------------------------------------------------
// Style access helpers
// --------------------------------------------------------------------------

// sourceStyles returns CT_Styles from the source document.
// Returns nil, error if source has no styles part.
func (ri *ResourceImporter) sourceStyles() (*oxml.CT_Styles, error) {
	return ri.sourceDoc.part.Styles()
}

// targetStyles returns CT_Styles from the target document.
// StylesPart auto-creates a default if not present.
func (ri *ResourceImporter) targetStyles() (*oxml.CT_Styles, error) {
	return ri.targetDoc.part.Styles()
}

// --------------------------------------------------------------------------
// Closure computation
// --------------------------------------------------------------------------

// collectStyleClosure performs BFS over style dependencies starting from
// seedIds. Returns styles in BFS order (children before parents).
//
// Dependencies traversed: basedOn, next, link. Missing styles in source
// are silently skipped (orphaned references).
func (ri *ResourceImporter) collectStyleClosure(seedIds []string) []*oxml.CT_Style {
	srcStyles, err := ri.sourceStyles()
	if err != nil {
		// No styles part in source — nothing to import.
		return nil
	}

	queue := make([]string, len(seedIds))
	copy(queue, seedIds)
	visited := map[string]bool{}
	var result []*oxml.CT_Style

	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		if visited[id] {
			continue
		}
		visited[id] = true

		s := srcStyles.GetByID(id)
		if s == nil {
			// Orphaned reference — style not defined in source.
			continue
		}
		result = append(result, s)

		// Dependencies: basedOn, next, link.
		if v, _ := s.BasedOnVal(); v != "" {
			queue = append(queue, v)
		}
		if v, _ := s.NextVal(); v != "" {
			queue = append(queue, v)
		}
		// link: no typed accessor in generated code — use raw etree.
		if link := s.RawElement().FindElement("w:link"); link != nil {
			if v := link.SelectAttrValue("w:val", ""); v != "" {
				queue = append(queue, v)
			}
		}
	}
	return result
}

// --------------------------------------------------------------------------
// Merge logic
// --------------------------------------------------------------------------

// mergeOneStyle merges a single source style into the target document
// using UseDestinationStyles strategy:
//
//   - If the style already exists in the target (by styleId), use the target
//     definition. The styleMap records the identity mapping.
//   - If the style is missing from the target, deep-copy it from source,
//     remap numId references inside the copy, and insert into target.
//
// Idempotent via styleMap check.
func (ri *ResourceImporter) mergeOneStyle(srcStyle *oxml.CT_Style) error {
	id := srcStyle.StyleId()
	if id == "" {
		return nil
	}
	if _, done := ri.styleMap[id]; done {
		return nil
	}

	tgtStyles, err := ri.targetStyles()
	if err != nil {
		return fmt.Errorf("docx: accessing target styles: %w", err)
	}

	if tgtStyles.GetByID(id) != nil {
		// UseDestinationStyles: style exists in template — use it.
		ri.styleMap[id] = id
		return nil
	}

	// Style not in target — deep-copy from source.
	clone := srcStyle.RawElement().Copy()

	// Remap numId inside the copied style definition (if present).
	// This is done HERE, not in remapAll, because styles inserted into
	// styles.xml are not part of the body element copies that remapAll
	// processes.
	ri.remapNumIdsInElement(clone)

	// Remap basedOn/next/link references through styleMap.
	// Under UseDestinationStyles this is a no-op (styleMap[id] == id),
	// but correct to implement for future KeepSourceFormatting support.
	ri.remapStyleRefsInElement(clone)

	tgtStyles.RawElement().AddChild(clone)
	ri.styleMap[id] = id
	return nil
}

// --------------------------------------------------------------------------
// Body scanning
// --------------------------------------------------------------------------

// collectStyleIdsFromBody scans the source document body for all unique
// style references (pStyle, rStyle, tblStyle).
// Returns a deduplicated slice in document order.
func collectStyleIdsFromBody(sourceDoc *Document) []string {
	srcBody := sourceDoc.element.Body()
	if srcBody == nil {
		return nil
	}
	return collectStyleIdsFromElements(srcBody.RawElement().ChildElements())
}

// collectStyleIdsFromElements scans element trees for style references.
// Looks for w:pStyle, w:rStyle, w:tblStyle elements and extracts w:val.
func collectStyleIdsFromElements(elements []*etree.Element) []string {
	seen := map[string]bool{}
	var result []string
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			if el.Space == "w" {
				switch el.Tag {
				case "pStyle", "rStyle", "tblStyle":
					if v := el.SelectAttrValue("w:val", ""); v != "" && !seen[v] {
						seen[v] = true
						result = append(result, v)
					}
				}
			}
			// Reverse push so LIFO stack yields document order.
			children := el.ChildElements()
			for i := len(children) - 1; i >= 0; i-- {
				stack = append(stack, children[i])
			}
		}
	}
	return result
}

// --------------------------------------------------------------------------
// In-element remapping helpers
// --------------------------------------------------------------------------

// remapNumIdsInElement rewrites w:numId values inside a cloned style
// element using the numIdMap populated during importNumbering.
// Called when deep-copying a style definition into the target.
func (ri *ResourceImporter) remapNumIdsInElement(el *etree.Element) {
	if len(ri.numIdMap) == 0 {
		return
	}
	stack := []*etree.Element{el}
	for len(stack) > 0 {
		node := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if node.Space == "w" && node.Tag == "numId" {
			ri.remapAttrValInt(node, ri.numIdMap)
		}
		stack = append(stack, node.ChildElements()...)
	}
}

// remapStyleRefsInElement rewrites basedOn/next/link values in a cloned
// style element using the styleMap.
//
// Under UseDestinationStyles this is a no-op because styleMap[id] == id
// (no renaming). Implemented for correctness and future extensibility.
func (ri *ResourceImporter) remapStyleRefsInElement(el *etree.Element) {
	if len(ri.styleMap) == 0 {
		return
	}
	for _, child := range el.ChildElements() {
		if child.Space != "w" {
			continue
		}
		switch child.Tag {
		case "basedOn", "next", "link":
			if v := child.SelectAttrValue("w:val", ""); v != "" {
				if newVal, ok := ri.styleMap[v]; ok {
					child.CreateAttr("w:val", newVal)
				}
			}
		}
	}
}

// --------------------------------------------------------------------------
// Default style mismatch detection and materialization
// --------------------------------------------------------------------------

// detectDefaultStyleMismatch compares the default paragraph styles of source
// and target documents. If they differ (by styleId), sets srcDefaultParaStyleId
// so that materializeImplicitStyles can add explicit pStyle to paragraphs
// that rely on the implicit default.
//
// This mirrors Aspose.Words behavior: when source default ≠ target default,
// paragraphs without explicit pStyle get the source default materialized
// to preserve their appearance.
func (ri *ResourceImporter) detectDefaultStyleMismatch() error {
	srcStyles, err := ri.sourceStyles()
	if err != nil {
		return err
	}
	tgtStyles, err := ri.targetStyles()
	if err != nil {
		return err
	}

	srcDefault, _ := srcStyles.DefaultFor(enum.WdStyleTypeParagraph)
	tgtDefault, _ := tgtStyles.DefaultFor(enum.WdStyleTypeParagraph)

	srcId := ""
	tgtId := ""
	if srcDefault != nil {
		srcId = srcDefault.StyleId()
	}
	if tgtDefault != nil {
		tgtId = tgtDefault.StyleId()
	}

	// If both have the same styleId (or both are empty), no materialization
	// needed — the target default will produce the same style resolution.
	if srcId == tgtId {
		return nil
	}

	ri.srcDefaultParaStyleId = srcId
	return nil
}

// materializeImplicitStyles adds explicit pStyle to paragraphs that rely
// on the implicit default paragraph style. Called from prepareContentElements
// after deep-copy/sanitize but before remapAll.
//
// Without this, paragraphs without pStyle silently change appearance when
// the source and target documents have different default paragraph styles
// (e.g. source Normal = Times New Roman 24pt, target default = Calibri 11pt).
//
// Only runs when detectDefaultStyleMismatch found a mismatch.
func (ri *ResourceImporter) materializeImplicitStyles(elements []*etree.Element) {
	if ri.srcDefaultParaStyleId == "" {
		return
	}
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			if el.Space == "w" && el.Tag == "p" {
				if !hasPStyle(el) {
					materializePStyle(el, ri.srcDefaultParaStyleId)
				}
			}
			stack = append(stack, el.ChildElements()...)
		}
	}
}

// hasPStyle returns true if the paragraph has an explicit pStyle in its pPr.
func hasPStyle(p *etree.Element) bool {
	for _, child := range p.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			for _, grandchild := range child.ChildElements() {
				if grandchild.Space == "w" && grandchild.Tag == "pStyle" {
					return true
				}
			}
			return false
		}
	}
	return false
}

// materializePStyle adds <w:pPr><w:pStyle w:val="styleId"/> to a paragraph
// that has no pPr, or adds <w:pStyle> into existing pPr.
func materializePStyle(p *etree.Element, styleId string) {
	// Find or create pPr.
	var pPr *etree.Element
	for _, child := range p.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			pPr = child
			break
		}
	}
	if pPr == nil {
		pPr = etree.NewElement("w:pPr")
		// pPr must be the first child of w:p per OOXML schema.
		p.InsertChildAt(0, pPr)
	}

	// Add pStyle as first child of pPr.
	pStyle := etree.NewElement("w:pStyle")
	pStyle.CreateAttr("w:val", styleId)
	pPr.InsertChildAt(0, pStyle)
}

// appendUnique appends val to slice if not already present.
func appendUnique(slice []string, val string) []string {
	for _, s := range slice {
		if s == val {
			return slice
		}
	}
	return append(slice, val)
}
