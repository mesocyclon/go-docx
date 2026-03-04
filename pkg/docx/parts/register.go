package parts

import (
	"strings"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// NewDocxPartFactory creates a PartFactory pre-configured with WML part
// constructors. This factory is used when opening a .docx package so that
// the generic OPC unmarshaller produces the correct part types.
//
// Mirrors Python PartFactory configuration in docx/__init__.py.
func NewDocxPartFactory() *opc.PartFactory {
	f := opc.NewPartFactory()

	// Register content-type → constructor mappings
	// Mirrors Python: PartFactory.part_type_for[CT.*] = *Part
	f.Register(opc.CTOpcCoreProperties, LoadCorePropertiesPart)
	f.Register(opc.CTWmlDocumentMain, LoadDocumentPart)
	f.Register(opc.CTWmlStyles, LoadStylesPart)
	f.Register(opc.CTWmlSettings, LoadSettingsPart)
	f.Register(opc.CTWmlComments, LoadCommentsPart)
	f.Register(opc.CTWmlHeader, LoadHeaderPart)
	f.Register(opc.CTWmlFooter, LoadFooterPart)
	f.Register(opc.CTWmlNumbering, LoadNumberingPart)
	f.Register(opc.CTWmlFootnotes, LoadFootnotesPart)
	f.Register(opc.CTWmlEndnotes, LoadEndnotesPart)

	// Selector: image/* content types with RTImage reltype → ImagePart
	f.SetSelector(func(contentType, relType string) opc.PartConstructor {
		if relType == opc.RTImage && strings.HasPrefix(contentType, "image/") {
			return LoadImagePart
		}
		return nil
	})

	return f
}

// LoadDocumentPart is a PartConstructor for loading DocumentPart from a package.
func LoadDocumentPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, err
	}
	return NewDocumentPart(xp), nil
}
