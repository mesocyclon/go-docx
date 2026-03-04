package oxml

import (
	"strings"

	"github.com/beevik/etree"
)

// --------------------------------------------------------------------------
// replacetext.go — cross-run text replacement engine
//
// Provides the core algorithm for replacing text across run boundaries
// in a paragraph. The approach uses "text atoms" — individual XML elements
// (<w:t>, <w:br>, <w:tab>, etc.) mapped to character positions in the
// concatenated paragraph text. This preserves formatting (<w:rPr>),
// non-textual content (<w:drawing>, <w:commentReference>), and XML
// structure (hyperlinks, comment ranges, bookmarks).
//
// See also: CT_P.ReplaceText in text_paragraph_custom.go (public entry point).
// --------------------------------------------------------------------------

// textAtom is the minimal unit of the text ↔ XML mapping.
// Each atom represents one child element of a <w:r> that contributes
// characters to the concatenated paragraph text.
//
//   - editable atoms (<w:t>): text can be changed arbitrarily via SetText.
//   - fixed atoms (<w:br>, <w:cr>, <w:tab>, <w:noBreakHyphen>, <w:ptab>):
//     produce exactly 1 character; can only be removed entirely.
type textAtom struct {
	elem     *etree.Element // the concrete XML element
	run      *etree.Element // parent <w:r> element, captured at collection time
	text     string         // text equivalent of this element
	startPos int            // byte offset in the concatenated string (byte-based, not rune-based — intentional)
	editable bool           // true for <w:t>, false for fixed elements
}

// collectTextAtoms walks the children of a <w:p> element, building a slice
// of text atoms and the concatenated full text of the paragraph.
//
// Traversal order: direct child <w:r> elements and <w:r> elements inside
// <w:hyperlink> children, in document order.
//
// Skipped at <w:p> level: <w:pPr>, <w:bookmarkStart>, <w:bookmarkEnd>,
// <w:commentRangeStart>, <w:commentRangeEnd>, <w:proofErr>, <w:ins>,
// <w:del>, <w:sdt>, and any other non-run/non-hyperlink children.
func collectTextAtoms(pElem *etree.Element) ([]textAtom, string) {
	var atoms []textAtom
	pos := 0

	for _, child := range pElem.ChildElements() {
		if child.Space != "w" {
			continue
		}
		switch child.Tag {
		case "r":
			collectRunAtoms(child, &atoms, &pos)
		case "hyperlink":
			for _, grandchild := range child.ChildElements() {
				if grandchild.Space == "w" && grandchild.Tag == "r" {
					collectRunAtoms(grandchild, &atoms, &pos)
				}
			}
		}
	}

	// Build the concatenated text from atoms.
	var sb strings.Builder
	for i := range atoms {
		sb.WriteString(atoms[i].text)
	}
	return atoms, sb.String()
}

// collectRunAtoms appends text atoms from a single <w:r> element.
//
// Collected elements:
//   - <w:t>             → editable, text = element.Text()
//   - <w:br type="">    → fixed, "\n"  (textWrapping or absent type)
//   - <w:cr>            → fixed, "\n"
//   - <w:tab>           → fixed, "\t"
//   - <w:noBreakHyphen> → fixed, "-"
//   - <w:ptab>          → fixed, "\t"
//
// Skipped: <w:rPr>, <w:drawing>, <w:lastRenderedPageBreak>,
// <w:commentReference>, <w:footnoteReference>, <w:endnoteReference>,
// <w:br type="page">, <w:br type="column">, and any other non-text children.
func collectRunAtoms(rElem *etree.Element, atoms *[]textAtom, pos *int) {
	for _, child := range rElem.ChildElements() {
		if child.Space != "w" {
			continue
		}
		switch child.Tag {
		case "t":
			text := child.Text()
			*atoms = append(*atoms, textAtom{
				elem:     child,
				run:      rElem,
				text:     text,
				startPos: *pos,
				editable: true,
			})
			*pos += len(text)

		case "br":
			// Only textWrapping breaks produce text; page/column do not.
			brType := etreeAttrVal(child, "w", "type")
			if brType == "" || brType == "textWrapping" {
				*atoms = append(*atoms, textAtom{
					elem:     child,
					run:      rElem,
					text:     "\n",
					startPos: *pos,
					editable: false,
				})
				*pos++
			}

		case "cr":
			*atoms = append(*atoms, textAtom{
				elem:     child,
				run:      rElem,
				text:     "\n",
				startPos: *pos,
				editable: false,
			})
			*pos++

		case "tab":
			*atoms = append(*atoms, textAtom{
				elem:     child,
				run:      rElem,
				text:     "\t",
				startPos: *pos,
				editable: false,
			})
			*pos++

		case "noBreakHyphen":
			*atoms = append(*atoms, textAtom{
				elem:     child,
				run:      rElem,
				text:     "-",
				startPos: *pos,
				editable: false,
			})
			*pos++

		case "ptab":
			*atoms = append(*atoms, textAtom{
				elem:     child,
				run:      rElem,
				text:     "\t",
				startPos: *pos,
				editable: false,
			})
			*pos++
		}
	}
}

// findOccurrences returns the byte-offset starting positions of all
// non-overlapping occurrences of old in fullText, left to right.
func findOccurrences(fullText, old string) []int {
	var positions []int
	start := 0
	oldLen := len(old)
	for {
		idx := strings.Index(fullText[start:], old)
		if idx < 0 {
			break
		}
		positions = append(positions, start+idx)
		start += idx + oldLen
	}
	return positions
}

// applyReplacements modifies the XML elements referenced by atoms for every
// occurrence of old in fullText, replacing it with new. Occurrences are
// processed right-to-left so that byte positions of earlier matches remain
// valid. Returns the number of replacements performed.
func applyReplacements(atoms []textAtom, fullText, old, new string) int {
	matches := findOccurrences(fullText, old)
	if len(matches) == 0 {
		return 0
	}

	oldLen := len(old)

	// Process right-to-left to keep positions stable.
	for i := len(matches) - 1; i >= 0; i-- {
		matchStart := matches[i]
		matchEnd := matchStart + oldLen

		replacementPlaced := false

		for j := range atoms {
			atom := &atoms[j]
			atomEnd := atom.startPos + len(atom.text)

			// No intersection with this atom?
			if atom.startPos >= matchEnd || atomEnd <= matchStart {
				continue
			}

			// Byte range within this atom covered by the match.
			cutStart := matchStart - atom.startPos
			if cutStart < 0 {
				cutStart = 0
			}
			cutEnd := matchEnd - atom.startPos
			if cutEnd > len(atom.text) {
				cutEnd = len(atom.text)
			}

			if atom.editable {
				insert := ""
				if !replacementPlaced {
					insert = new
					replacementPlaced = true
				}
				newText := atom.text[:cutStart] + insert + atom.text[cutEnd:]
				atom.elem.SetText(newText)
				ensurePreserveSpace(atom.elem)
				atom.text = newText

			} else {
				// Fixed atom (1 char, fully covered): remove the element.
				if parent := atom.elem.Parent(); parent != nil {
					parent.RemoveChild(atom.elem)
				}
			}
		}

		// Edge case: match consisted entirely of fixed atoms (e.g. "\t\t")
		// and replacement is non-empty. We need to create a <w:t> to hold it.
		if !replacementPlaced && new != "" {
			insertReplacementText(atoms, matches[i], matchEnd, new)
		}
	}

	return len(matches)
}

// insertReplacementText handles the rare edge case where a match covers
// only fixed atoms (no editable <w:t> to write the replacement into).
// It creates a new <w:t> element in the parent <w:r> of the first matched
// atom (known from atom.run, captured at collection time), inserted right
// after <w:rPr> to preserve correct element ordering.
func insertReplacementText(atoms []textAtom, matchStart, matchEnd int, replacement string) {
	// Find the first atom that intersects the match — its .run is the
	// correct parent, even if the atom's elem was already removed.
	for j := range atoms {
		atom := &atoms[j]
		atomEnd := atom.startPos + len(atom.text)
		if atom.startPos >= matchEnd || atomEnd <= matchStart {
			continue
		}

		tEl := OxmlElement("w:t")
		tEl.SetText(replacement)
		ensurePreserveSpace(tEl)
		insertAfterRPr(atom.run, tEl)
		return
	}
}

// insertAfterRPr inserts child into run immediately after the <w:rPr> element
// (if present), or at position 0 otherwise. This preserves the canonical
// element ordering within a <w:r>.
func insertAfterRPr(run, child *etree.Element) {
	for i, c := range run.Child {
		if elem, ok := c.(*etree.Element); ok && elem.Space == "w" && elem.Tag == "rPr" {
			run.InsertChildAt(i+1, child)
			return
		}
	}
	// No rPr — insert at the beginning of the run.
	run.InsertChildAt(0, child)
}

// ensurePreserveSpace sets or removes xml:space="preserve" on a <w:t>
// element based on its text content.
func ensurePreserveSpace(elem *etree.Element) {
	text := elem.Text()
	if text == "" || len(strings.TrimSpace(text)) < len(text) {
		elem.CreateAttr("xml:space", "preserve")
	} else {
		elem.RemoveAttr("xml:space")
	}
}

// etreeAttrVal returns the value of a namespace-prefixed attribute,
// or "" if not present.
func etreeAttrVal(elem *etree.Element, space, key string) string {
	for _, attr := range elem.Attr {
		if attr.Key == key && attr.Space == space {
			return attr.Value
		}
	}
	return ""
}
