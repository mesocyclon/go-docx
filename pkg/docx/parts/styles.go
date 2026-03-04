package parts

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// StylesPart is the proxy for the styles.xml part containing style definitions.
//
// Mirrors Python StylesPart(XmlPart).
type StylesPart struct {
	*opc.XmlPart
}

// NewStylesPart wraps an XmlPart as a StylesPart.
func NewStylesPart(xp *opc.XmlPart) *StylesPart {
	return &StylesPart{XmlPart: xp}
}

// Styles returns the CT_Styles wrapper for this part's root element.
//
// Mirrors Python StylesPart.styles property.
func (sp *StylesPart) Styles() (*oxml.CT_Styles, error) {
	el := sp.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: styles part element is nil")
	}
	return &oxml.CT_Styles{Element: oxml.WrapElement(el)}, nil
}

// DefaultStylesPart creates a new StylesPart with the default styles template.
//
// Mirrors Python StylesPart.default classmethod.
func DefaultStylesPart(pkg *opc.OpcPackage) (*StylesPart, error) {
	xmlBytes, err := templates.FS.ReadFile("default-styles.xml")
	if err != nil {
		return nil, fmt.Errorf("parts: reading default-styles.xml: %w", err)
	}
	el, err := oxml.ParseXml(xmlBytes)
	if err != nil {
		return nil, fmt.Errorf("parts: parsing default-styles.xml: %w", err)
	}
	pn := opc.PackURI("/word/styles.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTWmlStyles, el, pkg)
	return NewStylesPart(xp), nil
}

// LoadStylesPart is a PartConstructor for loading StylesPart from a package.
func LoadStylesPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading styles part %q: %w", partName, err)
	}
	return NewStylesPart(xp), nil
}
