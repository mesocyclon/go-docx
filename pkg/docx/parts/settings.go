package parts

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// SettingsPart is the document-level settings part of a WML package.
//
// Mirrors Python SettingsPart(XmlPart).
type SettingsPart struct {
	*opc.XmlPart
}

// NewSettingsPart wraps an XmlPart as a SettingsPart.
func NewSettingsPart(xp *opc.XmlPart) *SettingsPart {
	return &SettingsPart{XmlPart: xp}
}

// SettingsElement returns the CT_Settings wrapper for this part's root element.
//
// Mirrors Python SettingsPart.settings (element access portion — the domain
// Settings proxy is added in MR-11).
func (sp *SettingsPart) SettingsElement() (*oxml.CT_Settings, error) {
	el := sp.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: settings part element is nil")
	}
	return &oxml.CT_Settings{Element: oxml.WrapElement(el)}, nil
}

// DefaultSettingsPart creates a new SettingsPart from the default template.
//
// Mirrors Python SettingsPart.default classmethod.
func DefaultSettingsPart(pkg *opc.OpcPackage) (*SettingsPart, error) {
	xmlBytes, err := templates.FS.ReadFile("default-settings.xml")
	if err != nil {
		return nil, fmt.Errorf("parts: reading default-settings.xml: %w", err)
	}
	el, err := oxml.ParseXml(xmlBytes)
	if err != nil {
		return nil, fmt.Errorf("parts: parsing default-settings.xml: %w", err)
	}
	pn := opc.PackURI("/word/settings.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTWmlSettings, el, pkg)
	return NewSettingsPart(xp), nil
}

// LoadSettingsPart is a PartConstructor for loading SettingsPart from a package.
func LoadSettingsPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading settings part %q: %w", partName, err)
	}
	return NewSettingsPart(xp), nil
}
