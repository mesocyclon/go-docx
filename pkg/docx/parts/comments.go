package parts

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// CommentsPart is the container part for comments added to the document.
//
// Mirrors Python CommentsPart(StoryPart).
type CommentsPart struct {
	StoryPart
}

// NewCommentsPart wraps an XmlPart as a CommentsPart.
func NewCommentsPart(xp *opc.XmlPart) *CommentsPart {
	return &CommentsPart{
		StoryPart: StoryPart{XmlPart: xp},
	}
}

// CommentsElement returns the CT_Comments wrapper for this part's root element.
//
// Mirrors Python CommentsPart.comments (element access portion — the domain
// Comments proxy is added in MR-11).
func (cp *CommentsPart) CommentsElement() (*oxml.CT_Comments, error) {
	el := cp.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: comments part element is nil")
	}
	return &oxml.CT_Comments{Element: oxml.WrapElement(el)}, nil
}

// DefaultCommentsPart creates a new CommentsPart from the default template.
//
// Mirrors Python CommentsPart.default classmethod.
func DefaultCommentsPart(pkg *opc.OpcPackage) (*CommentsPart, error) {
	xmlBytes, err := templates.FS.ReadFile("default-comments.xml")
	if err != nil {
		return nil, fmt.Errorf("parts: reading default-comments.xml: %w", err)
	}
	el, err := oxml.ParseXml(xmlBytes)
	if err != nil {
		return nil, fmt.Errorf("parts: parsing default-comments.xml: %w", err)
	}
	pn := opc.PackURI("/word/comments.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTWmlComments, el, pkg)
	return NewCommentsPart(xp), nil
}

// LoadCommentsPart is a PartConstructor for loading CommentsPart from a package.
func LoadCommentsPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading comments part %q: %w", partName, err)
	}
	return NewCommentsPart(xp), nil
}
