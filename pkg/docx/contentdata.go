package docx

import (
	"fmt"
	"path"
	"strings"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/internal/xmlutil"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// --------------------------------------------------------------------------
// contentdata.go — ContentData type, relationship import, and rId helpers
//
// Contains the public ContentData structure for describing document content
// to insert in place of a text placeholder, private helper functions for
// scanning and rewriting relationship references (r:id, r:embed, r:link)
// in etree element trees, and the relationship import layer that copies
// parts (images, generic blobs) and external links from a source document
// into a target package.
//
// Phases 1–3 of the ReplaceWithContent feature.
// --------------------------------------------------------------------------

// ImportFormatMode controls how style conflicts between source and target
// documents are resolved during content import.
//
// This mirrors Aspose.Words ImportFormatMode — the industry-standard
// document processing library — providing three well-defined strategies
// for handling style ID collisions.
//
// Zero value (UseDestinationStyles) preserves backward compatibility.
type ImportFormatMode int

const (
	// UseDestinationStyles uses the destination document's style definition
	// when a style with the same ID exists in both documents. Styles present
	// only in the source are deep-copied (including basedOn/next/link chains).
	//
	// This is the default and preserves the current behavior.
	UseDestinationStyles ImportFormatMode = iota

	// KeepSourceFormatting preserves the source document's visual formatting.
	// When a style ID conflict occurs:
	//   - Default behavior: source style properties are expanded into direct
	//     paragraph/run attributes and the style reference is changed to the
	//     target's default paragraph style.
	//   - With ForceCopyStyles: the source style is copied with a unique
	//     suffix (_0, _1, ...) to avoid collision.
	KeepSourceFormatting

	// KeepDifferentStyles is a hybrid strategy. For each conflicting style:
	//   - If source and target definitions have identical formatting → uses
	//     the destination style (like UseDestinationStyles).
	//   - If formatting differs → behaves like KeepSourceFormatting (expands
	//     to direct attributes, or copies with suffix if ForceCopyStyles).
	KeepDifferentStyles
)

// ImportFormatOptions provides fine-grained control over content import
// behavior beyond what ImportFormatMode alone offers.
//
// All fields default to false (zero value), preserving backward compatibility.
// This mirrors Aspose.Words ImportFormatOptions.
type ImportFormatOptions struct {
	// ForceCopyStyles forces conflicting styles to be copied into the target
	// with a unique suffix (_0, _1, ...) instead of expanding formatting
	// into direct attributes. Only effective with KeepSourceFormatting and
	// KeepDifferentStyles modes.
	//
	// Mirrors Aspose.Words ImportFormatOptions.ForceCopyStyles.
	ForceCopyStyles bool

	// KeepSourceNumbering preserves source list numbering as a separate
	// list definition in the target. When false (default), source lists
	// merge into matching target lists and continue their numbering.
	//
	// Current project behavior is equivalent to KeepSourceNumbering=true
	// (always creates separate list definitions).
	//
	// Mirrors Aspose.Words ImportFormatOptions.KeepSourceNumbering.
	KeepSourceNumbering bool
}

// ContentData describes the content to insert in place of a text placeholder.
// It wraps a source Document whose body elements (paragraphs, tables) will
// replace each occurrence of the placeholder.
//
// Headers, footers, and section properties of the source document are NOT
// included — only the block-level children of <w:body>.
//
// The source Document must remain open (not garbage-collected) until
// ReplaceWithContent returns. After the call, the source may be closed.
//
// Backward compatibility: ContentData{Source: doc} uses zero-value Format
// (UseDestinationStyles) and zero-value Options — identical to the behavior
// before ImportFormatMode was introduced.
type ContentData struct {
	// Source is the opened document whose body content will be inserted.
	Source *Document

	// Format controls how style conflicts between source and target are
	// resolved. Default (zero value): UseDestinationStyles.
	Format ImportFormatMode

	// Options provides fine-grained import control.
	// Default (zero value): all options disabled.
	Options ImportFormatOptions
}

// preparedContent holds the pre-processed elements from a source document,
// ready to be cloned and inserted into the target.
type preparedContent struct {
	elements []*etree.Element // body elements with remapped rIds
}

// --------------------------------------------------------------------------
// Relationship-reference scanning
// --------------------------------------------------------------------------

// isRelAttr reports whether attr is a relationship-reference attribute.
// Covers the three attribute names used in OOXML to reference relationships:
//   - r:id    — on <w:hyperlink>, <w:headerReference>, <w:footerReference>, etc.
//   - r:embed — on <a:blip> (embedded images), <a:videoFile>
//   - r:link  — on <a:blip> (linked images)
//
// Both the short prefix form ("r") and the full namespace URI form are
// recognised.
func isRelAttr(attr etree.Attr) bool {
	switch attr.Key {
	case "id", "embed", "link":
		// Short prefix — most common case.
		if attr.Space == "r" {
			return true
		}
		// Full namespace URI form (rare but possible in hand-crafted XML).
		if strings.Contains(attr.Space, "officeDocument/2006/relationships") {
			return true
		}
	}
	return false
}

// collectReferencedRIds recursively scans elements and returns all unique
// rId values referenced via r:id, r:embed, or r:link attributes.
//
// The returned slice preserves first-encounter order and contains no
// duplicates. Elements are scanned depth-first.
func collectReferencedRIds(elements []*etree.Element) []string {
	seen := map[string]bool{}
	var result []string
	for _, root := range elements {
		// Iterative DFS via explicit stack.
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			for _, attr := range el.Attr {
				if isRelAttr(attr) && attr.Value != "" && !seen[attr.Value] {
					seen[attr.Value] = true
					result = append(result, attr.Value)
				}
			}
			// Push children (order within a single parent is not important
			// for collecting a unique set, but first-encounter order across
			// roots is preserved because we process roots sequentially).
			stack = append(stack, el.ChildElements()...)
		}
	}
	return result
}

// --------------------------------------------------------------------------
// Relationship-reference rewriting
// --------------------------------------------------------------------------

// remapRIds rewrites all relationship-reference attributes (r:id, r:embed,
// r:link) in the element trees according to ridMap (source rId → target rId).
// Attributes whose current value is not in ridMap are left unchanged.
func remapRIds(elements []*etree.Element, ridMap map[string]string) {
	for _, root := range elements {
		// Iterative DFS via explicit stack.
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			for i, attr := range el.Attr {
				if isRelAttr(attr) {
					if newRId, ok := ridMap[attr.Value]; ok {
						el.Attr[i].Value = newRId
					}
				}
			}
			stack = append(stack, el.ChildElements()...)
		}
	}
}

// --------------------------------------------------------------------------
// Relationship import (Phase 2)
// --------------------------------------------------------------------------

// importRelationship imports a single relationship from the source document
// into the target part. Returns the new rId in the target.
//
// importedParts tracks already-imported generic parts (source PartName →
// target Part) to avoid duplicating blobs when the same source part is
// referenced by multiple rIds or across multiple calls.
//
// Handles three cases:
//  1. External relationship (hyperlink URL) → GetOrAddExtRel
//  2. Image part → WmlPackage.GetOrAddImagePart (SHA-256 dedup) + relate
//  3. Other internal part (chart, drawing, VML, etc.) → copy blob, add part, relate
func importRelationship(
	srcRel *opc.Relationship,
	targetPart *parts.StoryPart,
	targetPkg *parts.WmlPackage,
	importedParts map[opc.PackURI]opc.Part,
) (string, error) {
	// Case 1: External relationship (e.g. hyperlink URL).
	if srcRel.IsExternal {
		rId := targetPart.Rels().GetOrAddExtRel(srcRel.RelType, srcRel.TargetRef)
		return rId, nil
	}

	// Internal relationship — need to import the target part.
	if srcRel.TargetPart == nil {
		return "", fmt.Errorf("internal relationship %s has no target part", srcRel.RID)
	}

	// Case 2: Image — use WmlPackage dedup.
	if srcRel.RelType == opc.RTImage {
		srcIP, ok := srcRel.TargetPart.(*parts.ImagePart)
		if !ok {
			return "", fmt.Errorf("image relationship %s target is %T, want *ImagePart",
				srcRel.RID, srcRel.TargetPart)
		}
		// Clone the ImagePart (blob + metadata) for the target package.
		blob, err := srcIP.Blob()
		if err != nil {
			return "", fmt.Errorf("reading image blob: %w", err)
		}
		cloneIP := parts.NewImagePart(srcIP.PartName(), srcIP.ContentType(), blob, nil)
		if w, err := srcIP.PxWidth(); err == nil {
			if h, errH := srcIP.PxHeight(); errH == nil {
				if hDpi, errD := srcIP.HorzDpi(); errD == nil {
					if vDpi, errV := srcIP.VertDpi(); errV == nil {
						cloneIP.SetImageMeta(w, h, hDpi, vDpi)
					}
				}
			}
		}
		cloneIP.SetFilename(srcIP.Filename())

		// GetOrAddImagePart handles SHA-256 dedup and partname allocation.
		dedupIP, err := targetPkg.GetOrAddImagePart(cloneIP)
		if err != nil {
			return "", fmt.Errorf("dedup image: %w", err)
		}
		// Wire relationship from targetPart to the (deduped) image part.
		rel := targetPart.Rels().GetOrAdd(opc.RTImage, dedupIP)
		return rel.RID, nil
	}

	// Case 3: Other internal part (chart, VML drawing, embedded package, etc.).
	// Generic import: copy blob into target package, relate.
	return importGenericPart(srcRel, targetPart, targetPkg, importedParts)
}

// importGenericPart copies a non-image internal part from source to target.
// Allocates a new partname in the target package and creates the relationship.
//
// importedParts provides dedup: if the same source part (by PartName) was
// already imported, the existing target part is reused — only a new
// relationship from targetPart is created. This prevents blob duplication
// when multiple rIds in the source point to the same part, or when the same
// source is inserted into multiple containers (body + header).
func importGenericPart(
	srcRel *opc.Relationship,
	targetPart *parts.StoryPart,
	targetPkg *parts.WmlPackage,
	importedParts map[opc.PackURI]opc.Part,
) (string, error) {
	srcPart := srcRel.TargetPart
	srcPN := srcPart.PartName()

	// Check dedup map first.
	if existing, ok := importedParts[srcPN]; ok {
		rel := targetPart.Rels().GetOrAdd(srcRel.RelType, existing)
		return rel.RID, nil
	}

	blob, err := srcPart.Blob()
	if err != nil {
		return "", fmt.Errorf("reading part blob: %w", err)
	}

	// Compute new partname. Use the source partname's directory + extension
	// pattern for NextPartname.
	template := partNameTemplate(srcPN)
	newPN := targetPkg.OpcPackage.NextPartname(template)

	// Create a BasePart copy in the target package.
	// Initialize empty Rels so IterParts (called during Save) does not
	// panic on part.Rels().All(). Sub-relationships of the source part
	// are NOT imported (shallow copy — see §13 of the plan).
	newPart := opc.NewBasePart(newPN, srcPart.ContentType(), blob, targetPkg.OpcPackage)
	newPart.SetRels(opc.NewRelationships(newPN.BaseURI()))
	targetPkg.OpcPackage.AddPart(newPart)

	// Record in dedup map.
	importedParts[srcPN] = newPart

	// Create relationship.
	targetRef := newPN.RelativeRef(targetPart.Rels().BaseURI())
	rel := targetPart.Rels().Add(srcRel.RelType, targetRef, newPart, false)
	return rel.RID, nil
}

// --------------------------------------------------------------------------
// Content preparation (Phase 3)
// --------------------------------------------------------------------------

// prepareContentElements extracts body elements from sourceDoc, deep-copies
// them, and remaps all relationships (images, hyperlinks, etc.) from source
// to targetPart.
//
// The returned elements are "template" copies — the caller must el.Copy()
// each one before inserting into the target tree (to support multiple
// placeholder replacements).
//
// ri is the shared ResourceImporter for this ReplaceWithContent call.
// It provides the target package, imported-parts dedup map, and (in later
// phases) style/numbering/footnote mappings.
func prepareContentElements(
	sourceDoc *Document,
	targetPart *parts.StoryPart,
	ri *ResourceImporter,
) (*preparedContent, error) {
	// Step 1: Extract body elements from source, skipping <w:sectPr>.
	srcBody := sourceDoc.element.Body()
	if srcBody == nil {
		return &preparedContent{}, nil
	}
	var srcElements []*etree.Element
	for _, child := range srcBody.RawElement().ChildElements() {
		if child.Space == "w" && child.Tag == "sectPr" {
			continue
		}
		srcElements = append(srcElements, child)
	}
	if len(srcElements) == 0 {
		return &preparedContent{}, nil
	}

	// Step 2: Deep-copy each element (do not modify source).
	elements := make([]*etree.Element, len(srcElements))
	for i, el := range srcElements {
		elements[i] = el.Copy()
	}

	// Step 2b: Sanitize copies for insertion into target.
	// Removes source-specific structural and annotation markup that
	// would be invalid or create ID collisions in the target document.
	// Must run before rId collection (Step 3) so that structural
	// references (e.g. headerReference in sectPr) are not imported.
	sanitizeForInsertion(elements)

	// Step 2c: Materialize implicit default styles.
	// If source and target have different default paragraph styles,
	// paragraphs without explicit pStyle get the source default added
	// so they keep their original appearance after insertion.
	ri.materializeImplicitStyles(elements)

	// Step 2d: Remap resource references (styles, numbering, footnotes).
	// Uses mappings populated during the import phase (Phase 1 of
	// ReplaceWithContent). In Phase 1 this is a no-op; branches are
	// added in Phases 2-4.
	ri.remapAll(elements)

	// Step 3: Collect all rId references from the copies.
	referencedRIds := collectReferencedRIds(elements)

	// Step 4: Import each referenced relationship into the target.
	if len(referencedRIds) > 0 {
		srcRels := sourceDoc.Part().Rels()
		ridMap := make(map[string]string, len(referencedRIds))
		for _, srcRId := range referencedRIds {
			srcRel := srcRels.GetByRID(srcRId)
			if srcRel == nil {
				continue // orphaned reference — skip silently
			}
			targetRId, err := importRelationship(srcRel, targetPart, ri.targetPkg, ri.importedParts)
			if err != nil {
				return nil, fmt.Errorf("docx: importing relationship %s: %w", srcRId, err)
			}
			ridMap[srcRId] = targetRId
		}

		// Step 5: Rewrite rId attributes in the copies.
		if len(ridMap) > 0 {
			remapRIds(elements, ridMap)
		}
	}

	return &preparedContent{elements: elements}, nil
}

// --------------------------------------------------------------------------
// Content sanitization for cross-document insertion
// --------------------------------------------------------------------------

// annotationMarkers is the set of OOXML annotation/range marker elements
// that carry document-scoped w:id values. These must be stripped when
// copying body content from one document to another because:
//
//   - Comment markers reference entries in the source's comments.xml,
//     which do not exist in the target.
//   - Bookmark, move-tracking, custom-XML-change, and permission markers
//     all use w:id that must be unique per document. Duplicating them
//     when the same source is inserted multiple times (or into a document
//     that already has overlapping IDs) produces invalid OOXML that Word
//     flags as corrupt.
//
// Renumbering these paired start/end markers correctly is complex (need
// to match pairs, update cross-references). Since annotations from the
// source have no semantic meaning in the target, stripping is the correct
// behavior for a content insertion API.
var annotationMarkers = map[string]bool{
	// Comment markers (reference comments.xml)
	"commentRangeStart": true,
	"commentRangeEnd":   true,
	"commentReference":  true,

	// Bookmark markers
	"bookmarkStart": true,
	"bookmarkEnd":   true,

	// Move tracking (revision markup)
	"moveFromRangeStart": true,
	"moveFromRangeEnd":   true,
	"moveToRangeStart":   true,
	"moveToRangeEnd":     true,

	// Custom XML change tracking
	"customXmlInsRangeStart":      true,
	"customXmlInsRangeEnd":        true,
	"customXmlDelRangeStart":      true,
	"customXmlDelRangeEnd":        true,
	"customXmlMoveFromRangeStart": true,
	"customXmlMoveFromRangeEnd":   true,
	"customXmlMoveToRangeStart":   true,
	"customXmlMoveToRangeEnd":     true,

	// Permission ranges
	"permStart": true,
	"permEnd":   true,
}

// sanitizeForInsertion removes source-specific markup from deep-copied
// elements in a single DFS pass. This is called after deep-copy (Step 2)
// and before rId collection (Step 3) so that structural references inside
// stripped elements are not imported into the target.
//
// Removes two categories of source-specific markup:
//
//  1. Paragraph-level sectPr — <w:sectPr> inside <w:pPr>. These carry
//     header/footer references (r:id) that are source-specific. Top-level
//     <w:sectPr> is already excluded in Step 1. Paragraph-level sectPr
//     would inject false section breaks and import structural parts as
//     generic BaseParts, corrupting the target.
//
//  2. Annotation range markers — 17 element types (comments, bookmarks,
//     move tracking, custom XML changes, permissions) that carry
//     document-scoped w:id values. These reference external state
//     (comments.xml) or must be unique per document.
func sanitizeForInsertion(elements []*etree.Element) {
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			var toRemove []*etree.Element
			for _, child := range el.ChildElements() {
				if child.Space != "w" {
					stack = append(stack, child)
					continue
				}

				// Category 1: sectPr inside pPr of a paragraph.
				if child.Tag == "sectPr" && el.Tag == "pPr" {
					toRemove = append(toRemove, child)
					continue
				}

				// Category 2: annotation range markers.
				if annotationMarkers[child.Tag] {
					toRemove = append(toRemove, child)
					continue
				}

				stack = append(stack, child)
			}

			for _, rm := range toRemove {
				el.RemoveChild(rm)
			}
		}
	}
}

// --------------------------------------------------------------------------
// Drawing ID renumbering
// --------------------------------------------------------------------------

// renumberDrawingIDs replaces all bare numeric `id` attributes in the element
// trees with fresh values obtained from nextID. This prevents duplicate
// docPr/cNvPr ids when the same source content is inserted multiple times
// or into a document that already has drawings.
//
// OOXML uses bare (non-namespaced) `id` attributes on elements like
// <wp:docPr> and <pic:cNvPr> as shape identifiers that must be unique
// within a story part. Without renumbering, deep-copied drawing elements
// carry their source ids and create duplicates that Word flags as corrupt.
//
// Only attributes matching (Key=="id", Space=="", value is all-digits) are
// replaced — the same criteria used by StoryPart.NextID / collectMaxID.
func renumberDrawingIDs(elements []*etree.Element, nextID func() int) {
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			for i, attr := range el.Attr {
				if attr.Key == "id" && attr.Space == "" && xmlutil.IsDigits(attr.Value) {
					el.Attr[i].Value = fmt.Sprintf("%d", nextID())
				}
			}
			stack = append(stack, el.ChildElements()...)
		}
	}
}

// --------------------------------------------------------------------------
// Partname helpers (Phase 2)
// --------------------------------------------------------------------------

// partNameTemplate converts a PackURI like "/word/charts/chart1.xml" into
// a printf template "/word/charts/chart%d.xml" for NextPartname.
//
// Uses path.Ext (not filepath.Ext) because OPC URIs always use forward
// slashes regardless of OS.
func partNameTemplate(pn opc.PackURI) string {
	s := string(pn)
	ext := path.Ext(s)
	base := s[:len(s)-len(ext)]
	// Strip trailing digits.
	// E.g. "/word/media/image3" → "/word/media/image"
	i := len(base)
	for i > 0 && base[i-1] >= '0' && base[i-1] <= '9' {
		i--
	}
	if i == len(base) {
		// No trailing digits — append %d before extension.
		return base + "%d" + ext
	}
	return base[:i] + "%d" + ext
}
