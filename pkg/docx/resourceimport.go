package docx

import (
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// --------------------------------------------------------------------------
// resourceimport.go — ResourceImporter: coordinates resource transfer from
// a source document into a target document during ReplaceWithContent.
//
// Created once per ReplaceWithContent call. Not safe for concurrent use.
// --------------------------------------------------------------------------

// ResourceImporter coordinates the transfer of resources (styles, numbering,
// footnotes, endnotes) from sourceDoc into targetDoc. It is created once per
// ReplaceWithContent call and passed through the entire pipeline so that all
// containers (body, headers, footers, comments) share the same mappings and
// dedup state.
type ResourceImporter struct {
	sourceDoc *Document
	targetDoc *Document
	targetPkg *parts.WmlPackage

	// Import configuration — controls style conflict resolution strategy
	// and fine-grained import options. Set once at construction time.
	importFormatMode ImportFormatMode
	opts             ImportFormatOptions

	// Dedup for generic parts (charts, VML) — existing mechanism.
	importedParts map[opc.PackURI]opc.Part

	// Numbering: source ID → target ID.
	numIdMap    map[int]int
	absNumIdMap map[int]int
	numDone     bool

	// Styles: source styleId → target styleId.
	// With UseDestinationStyles the value always equals the key.
	// With KeepSourceFormatting + ForceCopyStyles, the value may differ
	// (e.g. "Heading1" → "Heading1_0").
	styleMap  map[string]string
	styleDone bool

	// Styles marked for expansion to direct attributes.
	// Populated by mergeOneStyle when KeepSourceFormatting or
	// KeepDifferentStyles encounters a conflict (and ForceCopyStyles
	// is false). Key = source styleId, Value = source CT_Style.
	//
	// After import, expandDirectFormatting uses this map to walk the
	// prepared content elements and inline the source style properties
	// before remapAll changes the style references.
	expandStyles map[string]*oxml.CT_Style

	// When source and target have different default paragraph styles,
	// this holds the source default styleId so that paragraphs without
	// explicit pStyle can be materialized before remapAll.
	// Empty string means defaults match (no materialization needed).
	srcDefaultParaStyleId string

	// Footnotes: source id → target id.
	footnoteIdMap map[int]int
	endnoteIdMap  map[int]int
	footnotesDone bool
	endnotesDone  bool
}

// newResourceImporter creates a ResourceImporter for a single
// ReplaceWithContent call. All maps are initialized empty and ready to use.
//
// mode and opts control style conflict resolution — see ImportFormatMode
// and ImportFormatOptions. Zero values produce the original behavior
// (UseDestinationStyles, all options disabled).
func newResourceImporter(
	sourceDoc, targetDoc *Document,
	targetPkg *parts.WmlPackage,
	mode ImportFormatMode,
	opts ImportFormatOptions,
) *ResourceImporter {
	return &ResourceImporter{
		sourceDoc:        sourceDoc,
		targetDoc:        targetDoc,
		targetPkg:        targetPkg,
		importFormatMode: mode,
		opts:             opts,
		importedParts:    make(map[opc.PackURI]opc.Part),
		numIdMap:         make(map[int]int),
		absNumIdMap:      make(map[int]int),
		styleMap:         make(map[string]string),
		expandStyles:     make(map[string]*oxml.CT_Style),
		footnoteIdMap:    make(map[int]int),
		endnoteIdMap:     make(map[int]int),
	}
}

// remapAll rewrites resource references (styles, numbering, footnotes,
// endnotes) in the already-copied element trees using the mappings
// populated during the import phase.
//
// This is a single DFS pass over the elements — O(n) where n is the total
// number of etree nodes. Extending to a new resource type requires adding
// one case branch.
func (ri *ResourceImporter) remapAll(elements []*etree.Element) {
	// Skip traversal only when ALL maps are empty.
	if len(ri.numIdMap) == 0 && len(ri.styleMap) == 0 &&
		len(ri.footnoteIdMap) == 0 && len(ri.endnoteIdMap) == 0 {
		return
	}

	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			if el.Space == "w" {
				switch el.Tag {
				// Phase 2: numbering
				case "numId":
					ri.remapAttrValInt(el, ri.numIdMap)
				// Phase 3: styles
				case "pStyle", "rStyle", "tblStyle":
					ri.remapAttrVal(el, ri.styleMap)
				// Phase 4: footnotes/endnotes
				case "footnoteReference":
					ri.remapAttrId(el, ri.footnoteIdMap)
				case "endnoteReference":
					ri.remapAttrId(el, ri.endnoteIdMap)
				}
			}
			stack = append(stack, el.ChildElements()...)
		}
	}
}

// remapAttrValInt rewrites the w:val attribute of el using the given int map.
// If the current value is not in the map, the attribute is left unchanged.
func (ri *ResourceImporter) remapAttrValInt(el *etree.Element, m map[int]int) {
	v := el.SelectAttrValue("w:val", "")
	if v == "" {
		return
	}
	oldId, err := strconv.Atoi(v)
	if err != nil {
		return
	}
	if newId, ok := m[oldId]; ok {
		el.CreateAttr("w:val", strconv.Itoa(newId))
	}
}

// remapAttrVal rewrites the w:val attribute of el using the given string map.
// If the current value is not in the map, the attribute is left unchanged.
func (ri *ResourceImporter) remapAttrVal(el *etree.Element, m map[string]string) {
	v := el.SelectAttrValue("w:val", "")
	if v == "" {
		return
	}
	if newVal, ok := m[v]; ok {
		el.CreateAttr("w:val", newVal)
	}
}

// remapAttrId rewrites the w:id attribute of el using the given int map.
// Used for footnoteReference/endnoteReference elements (Phase 4).
func (ri *ResourceImporter) remapAttrId(el *etree.Element, m map[int]int) {
	v := el.SelectAttrValue("w:id", "")
	if v == "" {
		return
	}
	oldId, err := strconv.Atoi(v)
	if err != nil {
		return
	}
	if newId, ok := m[oldId]; ok {
		el.CreateAttr("w:id", strconv.Itoa(newId))
	}
}
