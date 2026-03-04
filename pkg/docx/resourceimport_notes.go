package docx

import (
	"fmt"
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// --------------------------------------------------------------------------
// resourceimport_notes.go — Footnote and endnote import for ResourceImporter
//
// Scans the source document body for footnoteReference / endnoteReference
// elements, copies the corresponding footnote/endnote bodies into the
// target document with fresh IDs, and imports styles and relationships
// within those bodies.
//
// The footnoteIdMap / endnoteIdMap are populated for later use by remapAll.
// --------------------------------------------------------------------------

// importFootnotes imports all footnotes referenced by the source document
// body into the target document. For each footnote body, styles, numbering
// references, and relationships (images, hyperlinks) are also imported.
//
// Safe to call multiple times — only runs once (idempotent via footnotesDone).
func (ri *ResourceImporter) importFootnotes() error {
	if ri.footnotesDone {
		return nil
	}
	ri.footnotesDone = true

	// 1. Collect all footnoteReference ids from source body.
	refIds := collectNoteRefsFromBody(ri.sourceDoc, "footnoteReference")
	if len(refIds) == 0 {
		return nil
	}

	// 2. Get source footnotes part.
	srcFP, err := ri.sourceDoc.part.FootnotesPart()
	if err != nil {
		// No footnotes part — references are orphaned, skip.
		return nil
	}

	// 3. Ensure target footnotes part exists.
	tgtFP, err := ri.targetDoc.part.GetOrAddFootnotesPart()
	if err != nil {
		return fmt.Errorf("docx: ensuring target footnotes part: %w", err)
	}

	// 4. Import each referenced footnote.
	return ri.importNoteEntries(
		refIds, ri.footnoteIdMap,
		srcFP.Element(), tgtFP.Element(), "footnote",
		&srcFP.StoryPart, &tgtFP.StoryPart,
	)
}

// importEndnotes imports all endnotes referenced by the source document
// body into the target document. Mirrors importFootnotes.
//
// Safe to call multiple times — only runs once (idempotent via endnotesDone).
func (ri *ResourceImporter) importEndnotes() error {
	if ri.endnotesDone {
		return nil
	}
	ri.endnotesDone = true

	// 1. Collect all endnoteReference ids from source body.
	refIds := collectNoteRefsFromBody(ri.sourceDoc, "endnoteReference")
	if len(refIds) == 0 {
		return nil
	}

	// 2. Get source endnotes part.
	srcEP, err := ri.sourceDoc.part.EndnotesPart()
	if err != nil {
		return nil
	}

	// 3. Ensure target endnotes part exists.
	tgtEP, err := ri.targetDoc.part.GetOrAddEndnotesPart()
	if err != nil {
		return fmt.Errorf("docx: ensuring target endnotes part: %w", err)
	}

	// 4. Import each referenced endnote.
	return ri.importNoteEntries(
		refIds, ri.endnoteIdMap,
		srcEP.Element(), tgtEP.Element(), "endnote",
		&srcEP.StoryPart, &tgtEP.StoryPart,
	)
}

// --------------------------------------------------------------------------
// Shared import engine
// --------------------------------------------------------------------------

// importNoteEntries is the shared engine for importing footnote or endnote
// entries. Called by importFootnotes and importEndnotes with resolved parts.
//
// The per-note pipeline mirrors prepareContentElements exactly:
// deep-copy → sanitize → import styles → materialize → remap → import rIds → renumber IDs → insert.
func (ri *ResourceImporter) importNoteEntries(
	refIds []int,
	idMap map[int]int,
	srcNotesEl *etree.Element,
	tgtNotesEl *etree.Element,
	noteTag string,
	srcStory *parts.StoryPart,
	tgtStory *parts.StoryPart,
) error {
	for _, srcId := range refIds {
		if _, done := idMap[srcId]; done {
			continue
		}

		// a. Find source note element.
		srcNote := findNoteById(srcNotesEl, noteTag, srcId)
		if srcNote == nil {
			continue
		}

		// b. Deep-copy.
		clone := srcNote.Copy()

		// c. Assign fresh ID in target.
		tgtId := nextNoteId(tgtNotesEl, noteTag)
		setNoteId(clone, tgtId)

		bodyEls := noteBodyElements(clone)

		// d. Sanitize: strip annotation markers (bookmarks, comment refs,
		// move tracking) that carry source-scoped w:id values.
		sanitizeForInsertion(bodyEls)

		// e. Import styles (delta — picks up FootnoteText etc. not in body).
		if err := ri.importStylesForElements(bodyEls); err != nil {
			return fmt.Errorf("docx: importing styles in %s %d: %w", noteTag, srcId, err)
		}

		// f. Materialize implicit default styles.
		ri.materializeImplicitStyles(bodyEls)

		// g. Remap styles and numId.
		ri.remapAll(bodyEls)

		// h. Import and remap relationships (images, hyperlinks).
		if err := ri.importRIdsForPart(bodyEls, srcStory, tgtStory); err != nil {
			return fmt.Errorf("docx: importing rIds in %s %d: %w", noteTag, srcId, err)
		}

		// i. Renumber drawing IDs (wp:docPr, pic:cNvPr) to avoid
		// collisions with existing drawings in target.
		renumberDrawingIDs(bodyEls, tgtStory.NextID)

		// j. Insert into target.
		tgtNotesEl.AddChild(clone)

		idMap[srcId] = tgtId
	}
	return nil
}

// --------------------------------------------------------------------------
// Relationship import for note bodies
// --------------------------------------------------------------------------

// importRIdsForPart imports relationships referenced by elements from
// srcPart into tgtPart. Reuses collectReferencedRIds, importRelationship,
// and remapRIds from contentdata.go.
func (ri *ResourceImporter) importRIdsForPart(
	elements []*etree.Element,
	srcPart *parts.StoryPart,
	tgtPart *parts.StoryPart,
) error {
	referencedRIds := collectReferencedRIds(elements)
	if len(referencedRIds) == 0 {
		return nil
	}

	srcRels := srcPart.Rels()
	ridMap := make(map[string]string, len(referencedRIds))

	for _, srcRId := range referencedRIds {
		srcRel := srcRels.GetByRID(srcRId)
		if srcRel == nil {
			continue
		}
		tgtRId, err := importRelationship(srcRel, tgtPart, ri.targetPkg, ri.importedParts)
		if err != nil {
			return fmt.Errorf("importing relationship %s: %w", srcRId, err)
		}
		ridMap[srcRId] = tgtRId
	}

	if len(ridMap) > 0 {
		remapRIds(elements, ridMap)
	}
	return nil
}

// --------------------------------------------------------------------------
// Body scanning
// --------------------------------------------------------------------------

// collectNoteRefsFromBody scans the source document body for all unique
// w:footnoteReference or w:endnoteReference w:id values.
// Returns a deduplicated slice. Skips id<=0 (separator/continuationSeparator).
func collectNoteRefsFromBody(sourceDoc *Document, refTag string) []int {
	srcBody := sourceDoc.element.Body()
	if srcBody == nil {
		return nil
	}
	return collectNoteRefsFromElements(srcBody.RawElement().ChildElements(), refTag)
}

// collectNoteRefsFromElements scans element trees for note reference IDs.
func collectNoteRefsFromElements(elements []*etree.Element, refTag string) []int {
	seen := map[int]bool{}
	var result []int
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			if el.Space == "w" && el.Tag == refTag {
				if v := el.SelectAttrValue("w:id", ""); v != "" {
					if id, err := strconv.Atoi(v); err == nil && id > 0 && !seen[id] {
						seen[id] = true
						result = append(result, id)
					}
				}
			}
			// Reverse push for document-order traversal.
			children := el.ChildElements()
			for i := len(children) - 1; i >= 0; i-- {
				stack = append(stack, children[i])
			}
		}
	}
	return result
}

// --------------------------------------------------------------------------
// Note element helpers
// --------------------------------------------------------------------------

// findNoteById finds a <w:footnote> or <w:endnote> child with the given
// w:id attribute value.
func findNoteById(notesEl *etree.Element, noteTag string, id int) *etree.Element {
	target := strconv.Itoa(id)
	for _, child := range notesEl.ChildElements() {
		if child.Space == "w" && child.Tag == noteTag {
			if child.SelectAttrValue("w:id", "") == target {
				return child
			}
		}
	}
	return nil
}

// nextNoteId returns the next available positive integer w:id for a
// footnote or endnote. Scans existing children, finds max id > 0,
// returns max+1. Starts at 1 (ids -1 and 0 are reserved for separators).
func nextNoteId(notesEl *etree.Element, noteTag string) int {
	maxId := 0
	for _, child := range notesEl.ChildElements() {
		if child.Space == "w" && child.Tag == noteTag {
			if v := child.SelectAttrValue("w:id", ""); v != "" {
				if id, err := strconv.Atoi(v); err == nil && id > maxId {
					maxId = id
				}
			}
		}
	}
	return maxId + 1
}

// setNoteId sets the w:id attribute on a footnote/endnote element.
func setNoteId(noteEl *etree.Element, id int) {
	noteEl.CreateAttr("w:id", strconv.Itoa(id))
}

// noteBodyElements returns the child elements of a footnote/endnote
// (paragraphs, tables) that may contain style, numId, and rId references.
func noteBodyElements(noteEl *etree.Element) []*etree.Element {
	return noteEl.ChildElements()
}
