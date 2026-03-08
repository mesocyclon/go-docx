package docx

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// resourceimport_styles.go — Style import for ResourceImporter
//
// Scans the source document body for pStyle, rStyle, tblStyle references,
// computes the transitive closure (basedOn, next, link chains), and merges
// each style into the target document using the configured ImportFormatMode:
//
//   - UseDestinationStyles: conflict → use target definition (no copy).
//   - KeepSourceFormatting: conflict → expand source formatting into direct
//     attributes (or copy with suffix if ForceCopyStyles).
//   - KeepDifferentStyles: conflict → compare formatting; identical = use
//     target, different = behave like KeepSourceFormatting.
//
// Missing styles are always deep-copied from source (all 3 modes agree).
//
// The styleMap is populated for later use by remapAll. The expandStyles
// map is populated for later use by expandDirectFormatting (Step 4).
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

// mergeOneStyle merges a single source style into the target document.
// Behavior depends on ri.importFormatMode:
//
// All modes — style NOT in target:
//
//	Deep-copy from source. All 3 modes agree on this.
//
// UseDestinationStyles — style EXISTS in target:
//
//	Use target definition. styleMap[id] = id.
//
// KeepSourceFormatting — style EXISTS in target:
//
//	ForceCopyStyles: copy with unique suffix (Heading1 → Heading1_0).
//	Default: mark for expansion to direct attributes. styleMap[id] = target default.
//
// KeepDifferentStyles — style EXISTS in target:
//
//	Identical formatting: use target (like UseDestinationStyles).
//	Different formatting: behave like KeepSourceFormatting.
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

	existing := tgtStyles.GetByID(id)

	// --- Style NOT in target: always copy (all 3 modes agree) ---
	if existing == nil {
		return ri.copyStyleToTarget(srcStyle, id)
	}

	// --- Style EXISTS in target: behavior depends on mode ---
	switch ri.importFormatMode {

	case UseDestinationStyles:
		// Conflict → use target definition. Original behavior.
		ri.styleMap[id] = id

	case KeepSourceFormatting:
		if ri.opts.ForceCopyStyles {
			// Copy with unique suffix: Heading1 → Heading1_0
			return ri.copyStyleToTarget(srcStyle, ri.uniqueStyleId(id))
		}
		// Default: mark for expansion to direct attributes (Step 4).
		ri.expandStyles[id] = srcStyle
		ri.styleMap[id] = ri.targetDefaultParaStyleId()

	case KeepDifferentStyles:
		if stylesContentEqual(srcStyle, existing) {
			// Same formatting → use target (like UseDestinationStyles).
			ri.styleMap[id] = id
		} else {
			// Different formatting → behave like KeepSourceFormatting.
			if ri.opts.ForceCopyStyles {
				return ri.copyStyleToTarget(srcStyle, ri.uniqueStyleId(id))
			}
			ri.expandStyles[id] = srcStyle
			ri.styleMap[id] = ri.targetDefaultParaStyleId()
		}
	}
	return nil
}

// copyStyleToTarget deep-copies srcStyle into the target styles.xml under
// targetId. If targetId differs from the source styleId, the clone's
// w:styleId attribute and w:name value are updated to prevent confusion
// in Word's style gallery.
//
// Handles numId and basedOn/next/link remapping on the clone.
func (ri *ResourceImporter) copyStyleToTarget(srcStyle *oxml.CT_Style, targetId string) error {
	tgtStyles, err := ri.targetStyles()
	if err != nil {
		return fmt.Errorf("docx: accessing target styles: %w", err)
	}

	clone := srcStyle.RawElement().Copy()

	// Rename styleId and display name when copying under a new ID.
	if targetId != srcStyle.StyleId() {
		clone.CreateAttr("w:styleId", targetId)
		if nameEl := findChild(clone, "w", "name"); nameEl != nil {
			if v := nameEl.SelectAttrValue("w:val", ""); v != "" {
				nameEl.CreateAttr("w:val", v+" (imported)")
			}
		}

		// Mark renamed copy as semi-hidden to avoid polluting Word's
		// Style Gallery with duplicate-looking entries. The style
		// remains fully functional and auto-appears in the gallery
		// when referenced by content (unhideWhenUsed).
		if findChild(clone, "w", "semiHidden") == nil {
			clone.AddChild(etree.NewElement("w:semiHidden"))
		}
		if findChild(clone, "w", "unhideWhenUsed") == nil {
			clone.AddChild(etree.NewElement("w:unhideWhenUsed"))
		}
	}

	// Remap numId inside the copied style definition (if present).
	// Done here because styles inserted into styles.xml are not part
	// of the body element copies that remapAll processes.
	ri.remapNumIdsInElement(clone)

	// Remap basedOn/next/link references through styleMap.
	ri.remapStyleRefsInElement(clone)

	tgtStyles.RawElement().AddChild(clone)
	ri.styleMap[srcStyle.StyleId()] = targetId
	return nil
}

// targetDefaultParaStyleId returns the styleId of the target document's
// default paragraph style. Used as the style reference replacement when
// expanding source formatting into direct attributes.
//
// Falls back to "Normal" if the target has no explicit default paragraph
// style (which is the OOXML-implied default).
func (ri *ResourceImporter) targetDefaultParaStyleId() string {
	tgtStyles, err := ri.targetStyles()
	if err != nil {
		return "Normal"
	}
	def, err := tgtStyles.DefaultFor(enum.WdStyleTypeParagraph)
	if err != nil || def == nil {
		return "Normal"
	}
	return def.StyleId()
}

// uniqueStyleId generates a unique styleId by appending _0, _1, etc. to
// the base ID. Checks both the target styles.xml and the current styleMap
// to avoid collisions with styles already imported in this session.
//
// Mirrors the Aspose.Words naming convention for ForceCopyStyles.
func (ri *ResourceImporter) uniqueStyleId(base string) string {
	tgtStyles, _ := ri.targetStyles()
	for i := 0; ; i++ {
		candidate := fmt.Sprintf("%s_%d", base, i)
		// Must not exist in target AND must not be already mapped.
		if tgtStyles.GetByID(candidate) == nil {
			if _, used := ri.styleMap[candidate]; !used {
				return candidate
			}
		}
	}
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

// --------------------------------------------------------------------------
// Style comparison
// --------------------------------------------------------------------------

// stylesContentEqual compares two styles by their formatting-relevant content,
// ignoring w:name (display name) and w:rsid* attributes (revision session
// IDs), which don't affect visual appearance.
//
// Used by KeepDifferentStyles to decide whether to use the target style
// (identical formatting) or expand to direct attributes (different formatting).
//
// The comparison serializes cloned, stripped style elements to canonical XML
// and compares byte-for-byte. This is robust against attribute ordering
// differences while being exact on content.
func stylesContentEqual(a, b *oxml.CT_Style) bool {
	ac := a.RawElement().Copy()
	bc := b.RawElement().Copy()
	stripNonFormattingAttrs(ac)
	stripNonFormattingAttrs(bc)

	var bufA, bufB bytes.Buffer

	docA := etree.NewDocument()
	docA.SetRoot(ac)
	docA.WriteSettings = etree.WriteSettings{CanonicalText: true}
	docA.WriteTo(&bufA)

	docB := etree.NewDocument()
	docB.SetRoot(bc)
	docB.WriteSettings = etree.WriteSettings{CanonicalText: true}
	docB.WriteTo(&bufB)

	return bytes.Equal(bufA.Bytes(), bufB.Bytes())
}

// stripNonFormattingAttrs removes w:name child elements and w:rsid*
// attributes from a cloned style element before comparison.
//
//   - w:name is a display label ("Heading 1") — two styles can have different
//     names but identical formatting.
//   - w:rsid* attributes are revision session IDs injected by Word on every
//     edit. They change constantly and carry no formatting information.
func stripNonFormattingAttrs(el *etree.Element) {
	// Remove w:name child element.
	var toRemove []*etree.Element
	for _, child := range el.ChildElements() {
		if child.Space == "w" && child.Tag == "name" {
			toRemove = append(toRemove, child)
		}
	}
	for _, rm := range toRemove {
		el.RemoveChild(rm)
	}

	// Remove rsid* attributes (revision session IDs).
	filtered := el.Attr[:0]
	for _, a := range el.Attr {
		if !strings.HasPrefix(a.Key, "rsid") {
			filtered = append(filtered, a)
		}
	}
	el.Attr = filtered
}

// --------------------------------------------------------------------------
// Expand to direct attributes (KeepSourceFormatting / KeepDifferentStyles)
// --------------------------------------------------------------------------

// expandDirectFormatting walks prepared content elements and for each
// paragraph/run whose style is in expandStyles, merges the resolved source
// formatting into direct attributes on that element.
//
// This is the core logic of KeepSourceFormatting mode: when a source style
// ID conflicts with a target style ID and ForceCopyStyles is false, the
// source style's formatting is "inlined" into every element that references
// it, preserving the visual appearance without modifying the target's style
// definitions.
//
// Handles BOTH paragraph styles (pStyle → pPr + rPr) and character styles
// (rStyle → rPr). This mirrors Aspose.Words which expands both style types
// to direct attributes.
//
// Pipeline position: after materializeImplicitStyles, before remapAll.
// expandDirectFormatting reads original (unmapped) styleIds; remapAll then
// replaces them with the target default.
func (ri *ResourceImporter) expandDirectFormatting(elements []*etree.Element) {
	if len(ri.expandStyles) == 0 {
		return
	}

	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			if el.Space == "w" {
				switch el.Tag {
				case "p":
					ri.expandParagraphStyle(el)
					// Process runs inline — avoids revisiting them from the
					// DFS stack (runs cannot contain nested paragraphs/tables,
					// so skipping their subtrees loses nothing).
					ri.expandRunStylesInParagraph(el)
				case "r":
					// Top-level runs outside paragraphs — rare but possible
					// in malformed OOXML. Process only if not already handled
					// by expandRunStylesInParagraph above.
					ri.expandRunStyle(el)
				}
			}

			// Push children but skip <w:r> inside <w:p> — already processed
			// by expandRunStylesInParagraph. Runs have no block-level
			// descendants, so nothing is lost.
			isParagraph := el.Space == "w" && el.Tag == "p"
			for _, child := range el.ChildElements() {
				if isParagraph && child.Space == "w" && child.Tag == "r" {
					continue
				}
				stack = append(stack, child)
			}
		}
	}
}

// expandParagraphStyle checks if the paragraph's pStyle is in expandStyles.
// If so, resolves the full formatting chain from the source style hierarchy
// and merges the result into the paragraph's pPr and its default rPr.
//
// Existing direct attributes take precedence over style-derived values
// (user formatting > style formatting), matching OOXML precedence rules.
func (ri *ResourceImporter) expandParagraphStyle(pEl *etree.Element) {
	pPr := findChild(pEl, "w", "pPr")
	if pPr == nil {
		return
	}
	pStyleEl := findChild(pPr, "w", "pStyle")
	if pStyleEl == nil {
		return
	}
	styleId := pStyleEl.SelectAttrValue("w:val", "")
	srcStyle, needsExpand := ri.expandStyles[styleId]
	if !needsExpand {
		return
	}

	// Resolve full formatting from source style chain.
	resolvedPPr, resolvedRPr := ri.resolveStyleChain(srcStyle)

	// Merge resolved pPr into existing pPr (direct attrs take precedence).
	if resolvedPPr != nil {
		mergePropertiesDeep(pPr, resolvedPPr)
	}

	// Merge resolved rPr into the paragraph-level default rPr.
	// This is the <w:rPr> inside <w:pPr> — defines default run formatting
	// for runs in this paragraph that don't have their own rPr.
	if resolvedRPr != nil {
		existingRPr := findChild(pPr, "w", "rPr")
		if existingRPr == nil {
			existingRPr = etree.NewElement("w:rPr")
			pPr.AddChild(existingRPr)
		}
		mergePropertiesDeep(existingRPr, resolvedRPr)
	}
}

// expandRunStylesInParagraph processes all <w:r> children of a paragraph,
// expanding character styles (rStyle) that are in expandStyles.
func (ri *ResourceImporter) expandRunStylesInParagraph(pEl *etree.Element) {
	for _, child := range pEl.ChildElements() {
		if child.Space == "w" && child.Tag == "r" {
			ri.expandRunStyle(child)
		}
	}
}

// expandRunStyle checks if the run's rStyle is in expandStyles. If so,
// resolves the source character style chain and merges the rPr into the
// run's existing rPr (direct attrs take precedence).
func (ri *ResourceImporter) expandRunStyle(rEl *etree.Element) {
	rPr := findChild(rEl, "w", "rPr")
	if rPr == nil {
		return
	}
	rStyleEl := findChild(rPr, "w", "rStyle")
	if rStyleEl == nil {
		return
	}
	styleId := rStyleEl.SelectAttrValue("w:val", "")
	srcStyle, needsExpand := ri.expandStyles[styleId]
	if !needsExpand {
		return
	}

	// For character styles, only rPr is relevant.
	_, resolvedRPr := ri.resolveStyleChain(srcStyle)
	if resolvedRPr != nil {
		mergePropertiesDeep(rPr, resolvedRPr)
	}
}

// resolveStyleChain walks the basedOn chain in the source styles part
// and merges pPr/rPr properties from base to derived (child overrides
// parent). Returns deep copies safe to modify.
//
// The chain is built root-first [style, parent, grandparent, ...] then
// merged in reverse order so that derived style properties override
// inherited ones.
//
// Cycle protection via visited set prevents infinite loops on malformed
// style definitions.
func (ri *ResourceImporter) resolveStyleChain(style *oxml.CT_Style) (pPr, rPr *etree.Element) {
	srcStyles, err := ri.sourceStyles()
	if err != nil {
		return nil, nil
	}

	// Build chain: [style, parent, grandparent, ...]
	var chain []*oxml.CT_Style
	visited := map[string]bool{}
	current := style
	for current != nil {
		id := current.StyleId()
		if visited[id] {
			break // cycle protection
		}
		visited[id] = true
		chain = append(chain, current)

		basedOn, _ := current.BasedOnVal()
		if basedOn == "" {
			break
		}
		current = srcStyles.GetByID(basedOn)
	}

	// Merge from base to derived (so derived overrides base).
	for i := len(chain) - 1; i >= 0; i-- {
		raw := chain[i].RawElement()
		if p := findChild(raw, "w", "pPr"); p != nil {
			if pPr == nil {
				pPr = p.Copy()
			} else {
				overridePropertiesDeep(pPr, p)
			}
		}
		if r := findChild(raw, "w", "rPr"); r != nil {
			if rPr == nil {
				rPr = r.Copy()
			} else {
				overridePropertiesDeep(rPr, r)
			}
		}
	}

	// Strip pStyle/rStyle from resolved properties — they are style-internal
	// references (basedOn inheritance), not meaningful as direct attributes.
	if pPr != nil {
		removeChild(pPr, "w", "pStyle")
	}
	if rPr != nil {
		removeChild(rPr, "w", "rStyle")
	}

	return
}

// --------------------------------------------------------------------------
// Property merging helpers
// --------------------------------------------------------------------------

// mergePropertiesDeep merges children of src into dst with attribute-level
// granularity, where dst (existing direct formatting) takes precedence.
//
// For each child in src:
//   - If dst has no child with the same space:tag → copy entire child from src
//   - If dst has child with the same space:tag → merge attributes from src
//     child into dst child (dst attributes take precedence)
//
// This produces correct results for complex properties like <w:rFonts>
// where src might have w:ascii and dst might have w:hAnsi — the result
// contains both attributes.
func mergePropertiesDeep(dst, src *etree.Element) {
	for _, srcChild := range src.ChildElements() {
		dstChild := findChild(dst, srcChild.Space, srcChild.Tag)
		if dstChild == nil {
			// Not present in dst — copy entire element from src.
			dst.AddChild(srcChild.Copy())
		} else {
			// Both have this property — merge attributes.
			// dst attributes take precedence (direct formatting > style).
			mergeAttrs(dstChild, srcChild)
		}
	}
}

// overridePropertiesDeep merges children of src into dst where src (derived
// style) takes precedence over dst (base style). Used during style chain
// resolution where child style overrides parent.
//
// For each child in src:
//   - If dst has no child with the same space:tag → copy from src
//   - If dst has child with the same space:tag → src attrs override dst attrs
func overridePropertiesDeep(dst, src *etree.Element) {
	for _, srcChild := range src.ChildElements() {
		dstChild := findChild(dst, srcChild.Space, srcChild.Tag)
		if dstChild == nil {
			dst.AddChild(srcChild.Copy())
		} else {
			// src (derived) overrides dst (base).
			overrideAttrs(dstChild, srcChild)
		}
	}
}

// mergeAttrs copies attributes from src to dst that don't already exist
// in dst. dst attributes take precedence — existing values are never
// overwritten. Used when merging style properties into direct formatting.
func mergeAttrs(dst, src *etree.Element) {
	dstKeys := make(map[string]bool, len(dst.Attr))
	for _, a := range dst.Attr {
		dstKeys[a.FullKey()] = true
	}
	for _, a := range src.Attr {
		if !dstKeys[a.FullKey()] {
			dst.Attr = append(dst.Attr, a)
		}
	}
}

// overrideAttrs copies attributes from src to dst, overwriting any existing
// attribute with the same key. Used during style chain resolution where
// derived style properties override inherited ones.
func overrideAttrs(dst, src *etree.Element) {
	for _, srcAttr := range src.Attr {
		dst.CreateAttr(srcAttr.FullKey(), srcAttr.Value)
	}
}

// removeChild removes the first child with given space:tag from el.
func removeChild(el *etree.Element, space, tag string) {
	if child := findChild(el, space, tag); child != nil {
		el.RemoveChild(child)
	}
}

// --------------------------------------------------------------------------
// Low-level etree helpers
// --------------------------------------------------------------------------

// findChild returns the first child element of el with the given namespace
// prefix and local tag name, or nil if not found.
//
// This is a package-level utility for working with raw *etree.Element trees
// where the oxml.Element.FindChild method is not available.
func findChild(el *etree.Element, space, tag string) *etree.Element {
	for _, child := range el.ChildElements() {
		if child.Space == space && child.Tag == tag {
			return child
		}
	}
	return nil
}

// --------------------------------------------------------------------------
// Misc helpers
// --------------------------------------------------------------------------

// appendUnique appends val to slice if not already present.
func appendUnique(slice []string, val string) []string {
	for _, s := range slice {
		if s == val {
			return slice
		}
	}
	return append(slice, val)
}
