package parts

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// NumberingPart is the proxy for the numbering.xml part containing numbering
// definitions for a document or glossary.
//
// Mirrors Python NumberingPart(XmlPart).
type NumberingPart struct {
	*opc.XmlPart
}

// NewNumberingPart wraps an XmlPart as a NumberingPart.
func NewNumberingPart(xp *opc.XmlPart) *NumberingPart {
	return &NumberingPart{XmlPart: xp}
}

// Numbering returns the CT_Numbering wrapper for this part's root element.
func (np *NumberingPart) Numbering() (*oxml.CT_Numbering, error) {
	el := np.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: numbering part element is nil")
	}
	return &oxml.CT_Numbering{Element: oxml.WrapElement(el)}, nil
}

// DefaultNumberingPart creates a new NumberingPart from the default template.
//
// The template contains an empty <w:numbering> element with all required
// namespace declarations. Used by GetOrAddNumberingPart when the target
// document has no numbering part.
func DefaultNumberingPart(pkg *opc.OpcPackage) (*NumberingPart, error) {
	xmlBytes, err := templates.FS.ReadFile("default-numbering.xml")
	if err != nil {
		return nil, fmt.Errorf("parts: reading default-numbering.xml: %w", err)
	}
	el, err := oxml.ParseXml(xmlBytes)
	if err != nil {
		return nil, fmt.Errorf("parts: parsing default-numbering.xml: %w", err)
	}
	pn := opc.PackURI("/word/numbering.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTWmlNumbering, el, pkg)
	return NewNumberingPart(xp), nil
}

// LoadNumberingPart is a PartConstructor for loading NumberingPart from a package.
func LoadNumberingPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading numbering part %q: %w", partName, err)
	}
	return NewNumberingPart(xp), nil
}
