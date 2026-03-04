package oxml

import (
	"fmt"
	"sort"
)

// ===========================================================================
// CT_Comments — custom methods
// ===========================================================================

// AddCommentFull adds a new <w:comment> child with the given id, author, and
// a skeleton paragraph containing a CommentText style and CommentReference run style.
// The returned element is the minimum valid comment ready for content addition.
func (cs *CT_Comments) AddCommentFull() (*CT_Comment, error) {
	nextID := cs.nextAvailableCommentID()
	xml := fmt.Sprintf(
		`<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `+
			`w:id="%d" w:author="">`+
			`<w:p>`+
			`<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>`+
			`<w:r>`+
			`<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>`+
			`<w:annotationRef/>`+
			`</w:r>`+
			`</w:p>`+
			`</w:comment>`, nextID,
	)
	el, err := ParseXml([]byte(xml))
	if err != nil {
		return nil, fmt.Errorf("oxml: failed to parse comment XML: %w", err)
	}
	comment := &CT_Comment{Element{e: el}}
	cs.e.AddChild(comment.e)
	return comment, nil
}

// GetCommentByID returns the <w:comment> element with the specified id, or nil if not found.
func (cs *CT_Comments) GetCommentByID(commentID int) *CT_Comment {
	for _, c := range cs.CommentList() {
		id, err := c.Id()
		if err == nil && id == commentID {
			return c
		}
	}
	return nil
}

// nextAvailableCommentID returns the next available comment id.
// Uses max(existing_ids) + 1, falling back to sequential gap-filling
// if that would exceed 32-bit signed integer range.
func (cs *CT_Comments) nextAvailableCommentID() int {
	var usedIDs []int
	for _, c := range cs.CommentList() {
		id, err := c.Id()
		if err == nil {
			usedIDs = append(usedIDs, id)
		}
	}

	if len(usedIDs) == 0 {
		return 0
	}

	maxID := usedIDs[0]
	for _, id := range usedIDs[1:] {
		if id > maxID {
			maxID = id
		}
	}

	nextID := maxID + 1
	if nextID <= (1<<31 - 1) {
		return nextID
	}

	// Fallback: find first unused integer starting from 0
	sort.Ints(usedIDs)
	for expected, actual := range usedIDs {
		if expected != actual {
			return expected
		}
	}
	return len(usedIDs)
}

// ===========================================================================
// CT_Comment — custom methods
// ===========================================================================

// InnerContentElements returns all <w:p> and <w:tbl> direct children in document order.
func (c *CT_Comment) InnerContentElements() []BlockItem {
	var result []BlockItem
	for _, child := range c.e.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			result = append(result, &CT_P{Element{e: child}})
		} else if child.Space == "w" && child.Tag == "tbl" {
			result = append(result, &CT_Tbl{Element{e: child}})
		}
	}
	return result
}
