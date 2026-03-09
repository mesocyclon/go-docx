package docx

import (
	"fmt"
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// renumberBookmarks rewrites w:id on bookmarkStart/bookmarkEnd pairs
// with fresh values from docPart.NextBookmarkID(), and appends a numeric
// suffix to w:name on bookmarkStart to avoid name collisions in the
// target document.
//
// Bookmark w:id must be unique per document (paired start/end share
// the same id). Bookmark w:name must also be unique — Word silently
// discards duplicates.
//
// Two-pass algorithm:
//  1. DFS pass 1: collect all bookmarkStart w:id values → build
//     oldId → newId map using docPart.NextBookmarkID(). Simultaneously
//     append suffix to w:name (e.g. "_imp1", "_imp2") to guarantee uniqueness.
//  2. DFS pass 2: rewrite w:id on both bookmarkStart and bookmarkEnd
//     using the map.
func renumberBookmarks(elements []*etree.Element, docPart *parts.DocumentPart) {
	// Pass 1: collect bookmarkStart ids and build mapping.
	idMap := map[int]int{}
	walkBookmarks(elements, func(el *etree.Element) {
		if el.Tag != "bookmarkStart" {
			return
		}
		oldID := parseBookmarkID(el)
		if oldID < 0 {
			return
		}
		if _, ok := idMap[oldID]; !ok {
			idMap[oldID] = docPart.NextBookmarkID()
		}
		// Deduplicate name by appending suffix.
		name := el.SelectAttrValue("w:name", "")
		if name != "" {
			suffix := docPart.NextBookmarkNameSuffix()
			el.CreateAttr("w:name", fmt.Sprintf("%s_imp%d", name, suffix))
		}
	})

	if len(idMap) == 0 {
		return
	}

	// Pass 2: rewrite w:id on bookmarkStart and bookmarkEnd.
	walkBookmarks(elements, func(el *etree.Element) {
		oldID := parseBookmarkID(el)
		if oldID < 0 {
			return
		}
		if newID, ok := idMap[oldID]; ok {
			el.CreateAttr("w:id", strconv.Itoa(newID))
		}
	})
}

// walkBookmarks performs a DFS over all elements and calls fn for each
// bookmarkStart or bookmarkEnd element encountered.
func walkBookmarks(elements []*etree.Element, fn func(el *etree.Element)) {
	for _, root := range elements {
		stack := []*etree.Element{root}
		for len(stack) > 0 {
			el := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			if el.Space == "w" && (el.Tag == "bookmarkStart" || el.Tag == "bookmarkEnd") {
				fn(el)
			}
			stack = append(stack, el.ChildElements()...)
		}
	}
}

// parseBookmarkID extracts the w:id integer from a bookmark element.
// Returns -1 if the attribute is missing or not a valid integer.
func parseBookmarkID(el *etree.Element) int {
	v := el.SelectAttrValue("w:id", "")
	if v == "" {
		return -1
	}
	id, err := strconv.Atoi(v)
	if err != nil {
		return -1
	}
	return id
}
