package docx

import (
	"fmt"
	"time"

	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Comments provides access to comments added to this document.
//
// Mirrors Python Comments.
type Comments struct {
	commentsElm  *oxml.CT_Comments
	commentsPart *parts.CommentsPart
}

// newComments creates a new Comments proxy.
func newComments(elm *oxml.CT_Comments, part *parts.CommentsPart) *Comments {
	return &Comments{commentsElm: elm, commentsPart: part}
}

// Len returns the number of comments in this collection.
func (cs *Comments) Len() int {
	return len(cs.commentsElm.CommentList())
}

// Iter returns all comments in this collection.
func (cs *Comments) Iter() []*Comment {
	list := cs.commentsElm.CommentList()
	result := make([]*Comment, len(list))
	for i, c := range list {
		result[i] = newComment(c, cs.commentsPart)
	}
	return result
}

// ReplaceText replaces all occurrences of old with new in every comment's
// content. Returns the total number of replacements performed.
func (cs *Comments) ReplaceText(old, new string) int {
	count := 0
	for _, c := range cs.Iter() {
		count += c.ReplaceText(old, new)
	}
	return count
}

// Get returns the comment identified by commentID, or nil if not found.
//
// Mirrors Python Comments.get.
func (cs *Comments) Get(commentID int) *Comment {
	c := cs.commentsElm.GetCommentByID(commentID)
	if c == nil {
		return nil
	}
	return newComment(c, cs.commentsPart)
}

// AddComment adds a new comment and returns it.
//
// If text is non-empty it is added to the comment; newlines create new paragraphs.
// Author is required but may be empty. Initials is optional; nil omits it.
//
// Mirrors Python Comments.add_comment.
func (cs *Comments) AddComment(text, author string, initials *string) (*Comment, error) {
	commentElm, err := cs.commentsElm.AddCommentFull()
	if err != nil {
		return nil, fmt.Errorf("docx: adding comment: %w", err)
	}

	if err := commentElm.SetAuthor(author); err != nil {
		return nil, err
	}
	if initials != nil {
		if err := commentElm.SetInitials(*initials); err != nil {
			return nil, err
		}
	}
	// Set date to now in ISO 8601 format
	now := time.Now().UTC().Format(time.RFC3339)
	if err := commentElm.SetDate(now); err != nil {
		return nil, err
	}

	comment := newComment(commentElm, cs.commentsPart)

	if text == "" {
		return comment, nil
	}

	// Split text on newlines; first paragraph uses existing empty paragraph
	paragraphs := splitNewlines(text)
	paras := comment.Paragraphs()
	if len(paras) > 0 && len(paragraphs) > 0 {
		if _, err := paras[0].AddRun(paragraphs[0]); err != nil {
			return nil, err
		}
		for _, s := range paragraphs[1:] {
			if _, err := comment.AddParagraph(s); err != nil {
				return nil, err
			}
		}
	}
	return comment, nil
}

// splitNewlines splits a string on newline sequences.
// Handles Unix (\n), Windows (\r\n), and classic Mac (\r) line endings.
func splitNewlines(s string) []string {
	// Normalize: \r\n → \n, then remaining \r → \n, then split.
	// Single pass: scan for \r and \n.
	var result []string
	start := 0
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '\r':
			result = append(result, s[start:i])
			// consume \n in a \r\n pair
			if i+1 < len(s) && s[i+1] == '\n' {
				i++
			}
			start = i + 1
		case '\n':
			result = append(result, s[start:i])
			start = i + 1
		}
	}
	result = append(result, s[start:])
	return result
}

// Comment is a single comment in the document.
//
// A comment is also a block-item container (can contain paragraphs and tables).
//
// Mirrors Python Comment(BlockItemContainer).
type Comment struct {
	BlockItemContainer
	commentElm *oxml.CT_Comment
}

// newComment creates a new Comment proxy.
func newComment(elm *oxml.CT_Comment, part *parts.CommentsPart) *Comment {
	sp := &part.StoryPart
	return &Comment{
		BlockItemContainer: newBlockItemContainer(elm.RawElement(), sp),
		commentElm:         elm,
	}
}

// Author returns the recorded author of this comment.
//
// Mirrors Python Comment.author (getter).
func (c *Comment) Author() (string, error) {
	v, err := c.commentElm.Author()
	if err != nil {
		return "", fmt.Errorf("docx: reading comment author: %w", err)
	}
	return v, nil
}

// SetAuthor sets the author of this comment.
//
// Mirrors Python Comment.author (setter).
func (c *Comment) SetAuthor(v string) error {
	return c.commentElm.SetAuthor(v)
}

// CommentID returns the unique identifier of this comment.
//
// Mirrors Python Comment.comment_id.
func (c *Comment) CommentID() (int, error) {
	return c.commentElm.Id()
}

// Initials returns the recorded initials of the comment author.
// Returns "" if not set.
//
// Mirrors Python Comment.initials (getter).
func (c *Comment) Initials() string {
	return c.commentElm.Initials()
}

// SetInitials sets the initials. Passing "" removes the attribute.
//
// Mirrors Python Comment.initials (setter).
func (c *Comment) SetInitials(v string) error {
	return c.commentElm.SetInitials(v)
}

// AddParagraph adds a paragraph to this comment. When style is nil, the
// "CommentText" paragraph style is applied (matching Word UI behavior).
//
// Mirrors Python Comment.add_paragraph override.
func (c *Comment) AddParagraph(text string, style ...StyleRef) (*Paragraph, error) {
	para, err := c.BlockItemContainer.AddParagraph(text, style...)
	if err != nil {
		return nil, err
	}
	// When no explicit style provided, apply CommentText directly to element
	// (same as Python: paragraph._p.style = "CommentText")
	if len(style) == 0 || style[0] == nil {
		commentText := "CommentText"
		if err := para.CT_P().SetStyle(&commentText); err != nil {
			return nil, fmt.Errorf("docx: setting comment paragraph style: %w", err)
		}
	}
	return para, nil
}

// Text returns the text content of this comment, paragraph boundaries
// separated by newlines.
//
// Mirrors Python Comment.text.
func (c *Comment) Text() string {
	paras := c.Paragraphs()
	if len(paras) == 0 {
		return ""
	}
	result := paras[0].Text()
	for _, p := range paras[1:] {
		result += "\n" + p.Text()
	}
	return result
}

// Timestamp returns the date/time this comment was authored, or nil
// if not set. The value is parsed from ISO 8601 format.
//
// Mirrors Python Comment.timestamp.
func (c *Comment) Timestamp() (*time.Time, error) {
	s := c.commentElm.Date()
	if s == "" {
		return nil, nil
	}
	t, err := time.Parse(time.RFC3339, s)
	if err != nil {
		return nil, fmt.Errorf("docx: parsing comment timestamp %q: %w", s, err)
	}
	return &t, nil
}
