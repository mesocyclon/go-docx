package docx

import (
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Body is a proxy for the <w:body> element in the document.
// Its primary role is as a container for document content (paragraphs,
// tables, sections).
//
// Mirrors Python _Body(BlockItemContainer).
type Body struct {
	BlockItemContainer
	ctBody *oxml.CT_Body
}

// newBody creates a Body wrapping the given CT_Body.
func newBody(ctBody *oxml.CT_Body, part *parts.StoryPart) *Body {
	return &Body{
		BlockItemContainer: newBlockItemContainer(ctBody.RawElement(), part),
		ctBody:             ctBody,
	}
}

// ClearContent removes all content from this body, preserving the
// trailing section properties.
//
// Mirrors Python _Body.clear_content.
func (b *Body) ClearContent() {
	b.ctBody.ClearContent()
}
