package parts

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// --------------------------------------------------------------------------
// HeaderPart
// --------------------------------------------------------------------------

// HeaderPart represents a section header definition.
//
// Mirrors Python HeaderPart(StoryPart).
type HeaderPart struct {
	StoryPart
}

// NewHeaderPart creates a new HeaderPart from the default header template,
// assigning the next available partname from the package.
//
// Mirrors Python HeaderPart.new(package).
func NewHeaderPart(pkg *opc.OpcPackage) (*HeaderPart, error) {
	xmlBytes, err := templates.FS.ReadFile("default-header.xml")
	if err != nil {
		return nil, fmt.Errorf("parts: reading default-header.xml: %w", err)
	}
	el, err := oxml.ParseXml(xmlBytes)
	if err != nil {
		return nil, fmt.Errorf("parts: parsing default-header.xml: %w", err)
	}
	pn := pkg.NextPartname("/word/header%d.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTWmlHeader, el, pkg)
	return &HeaderPart{StoryPart: StoryPart{XmlPart: xp}}, nil
}

// LoadHeaderPart is a PartConstructor for loading HeaderPart from a package.
func LoadHeaderPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading header part %q: %w", partName, err)
	}
	return &HeaderPart{StoryPart: StoryPart{XmlPart: xp}}, nil
}

// --------------------------------------------------------------------------
// FooterPart
// --------------------------------------------------------------------------

// FooterPart represents a section footer definition.
//
// Mirrors Python FooterPart(StoryPart).
type FooterPart struct {
	StoryPart
}

// NewFooterPart creates a new FooterPart from the default footer template,
// assigning the next available partname from the package.
//
// Mirrors Python FooterPart.new(package).
func NewFooterPart(pkg *opc.OpcPackage) (*FooterPart, error) {
	xmlBytes, err := templates.FS.ReadFile("default-footer.xml")
	if err != nil {
		return nil, fmt.Errorf("parts: reading default-footer.xml: %w", err)
	}
	el, err := oxml.ParseXml(xmlBytes)
	if err != nil {
		return nil, fmt.Errorf("parts: parsing default-footer.xml: %w", err)
	}
	pn := pkg.NextPartname("/word/footer%d.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTWmlFooter, el, pkg)
	return &FooterPart{StoryPart: StoryPart{XmlPart: xp}}, nil
}

// LoadFooterPart is a PartConstructor for loading FooterPart from a package.
func LoadFooterPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading footer part %q: %w", partName, err)
	}
	return &FooterPart{StoryPart: StoryPart{XmlPart: xp}}, nil
}
