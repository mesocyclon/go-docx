package parts

import (
	"fmt"
	"time"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// CorePropertiesPart is the proxy for the /docProps/core.xml part containing
// Dublin Core document metadata (author, title, created, etc.).
//
// Mirrors Python CorePropertiesPart(XmlPart).
type CorePropertiesPart struct {
	*opc.XmlPart
}

// NewCorePropertiesPart wraps an XmlPart as a CorePropertiesPart.
func NewCorePropertiesPart(xp *opc.XmlPart) *CorePropertiesPart {
	return &CorePropertiesPart{XmlPart: xp}
}

// CT returns the CT_CoreProperties wrapper for this part's root element.
//
// Mirrors Python CorePropertiesPart.core_properties → CoreProperties(self.element).
// (The domain-level CoreProperties proxy in pkg/docx/coreprops.go wraps this.)
func (cp *CorePropertiesPart) CT() (*oxml.CT_CoreProperties, error) {
	el := cp.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: core properties part element is nil")
	}
	return &oxml.CT_CoreProperties{Element: oxml.WrapElement(el)}, nil
}

// DefaultCorePropertiesPart creates a new CorePropertiesPart initialized with
// default values matching Python CorePropertiesPart.default():
//   - title        = "Word Document"
//   - last_modified_by = "go-docx"
//   - revision     = 1
//   - modified     = now (UTC)
//
// Mirrors Python CorePropertiesPart.default classmethod.
func DefaultCorePropertiesPart(pkg *opc.OpcPackage) (*CorePropertiesPart, error) {
	ct, err := oxml.NewCoreProperties()
	if err != nil {
		return nil, fmt.Errorf("parts: creating default core properties: %w", err)
	}

	pn := opc.PackURI("/docProps/core.xml")
	xp := opc.NewXmlPartFromElement(pn, opc.CTOpcCoreProperties, ct.RawElement(), pkg)
	part := NewCorePropertiesPart(xp)

	// Set default values — mirrors Python CorePropertiesPart.default
	ct.SetTitleText("Word Document")    //nolint:errcheck
	ct.SetLastModifiedByText("go-docx") //nolint:errcheck
	ct.SetRevisionNumber(1)             //nolint:errcheck
	ct.SetModifiedDatetime(time.Now().UTC())

	return part, nil
}

// LoadCorePropertiesPart is a PartConstructor for loading CorePropertiesPart
// from a package.
//
// Mirrors Python PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart.
func LoadCorePropertiesPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading core properties %q: %w", partName, err)
	}
	return NewCorePropertiesPart(xp), nil
}
