package docx

import (
	"fmt"
	"io"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/parts"
	"github.com/vortex/go-docx/pkg/docx/templates"
)

// New creates a new empty Document from the built-in default template.
//
// Mirrors Python: Document(None) → loads default.docx from templates.
func New() (*Document, error) {
	defaultDocx, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		return nil, fmt.Errorf("docx: reading default template: %w", err)
	}
	return OpenBytes(defaultDocx)
}

// Open creates a Document from an io.ReaderAt.
//
// Mirrors Python: Document(stream).
func Open(r io.ReaderAt, size int64) (*Document, error) {
	factory := parts.NewDocxPartFactory()
	pkg, err := opc.Open(r, size, factory)
	if err != nil {
		return nil, fmt.Errorf("docx: opening package: %w", err)
	}
	return documentFromPackage(pkg)
}

// OpenFile creates a Document from a file path.
//
// Mirrors Python: Document("/path/to/file.docx").
func OpenFile(path string) (*Document, error) {
	factory := parts.NewDocxPartFactory()
	pkg, err := opc.OpenFile(path, factory)
	if err != nil {
		return nil, fmt.Errorf("docx: opening file %q: %w", path, err)
	}
	return documentFromPackage(pkg)
}

// OpenBytes creates a Document from a byte slice.
func OpenBytes(data []byte) (*Document, error) {
	factory := parts.NewDocxPartFactory()
	pkg, err := opc.OpenBytes(data, factory)
	if err != nil {
		return nil, fmt.Errorf("docx: opening bytes: %w", err)
	}
	return documentFromPackage(pkg)
}

// documentFromPackage wires up a Document from a loaded OpcPackage.
//
// Mirrors Python api.py logic:
//  1. Get main document part
//  2. Validate content type
//  3. Create WmlPackage wrapper, run AfterUnmarshal
//  4. Create Document
func documentFromPackage(pkg *opc.OpcPackage) (*Document, error) {
	mainPart, err := pkg.MainDocumentPart()
	if err != nil {
		return nil, fmt.Errorf("docx: no main document part: %w", err)
	}
	docPart, ok := mainPart.(*parts.DocumentPart)
	if !ok {
		return nil, fmt.Errorf("docx: main part is %T, expected *DocumentPart", mainPart)
	}
	// Validate content type (mirrors Python check: CT.WML_DOCUMENT_MAIN).
	ct := docPart.ContentType()
	if ct != opc.CTWmlDocumentMain && ct != opc.CTWmlDocument {
		return nil, fmt.Errorf("docx: not a Word file, content type is %q", ct)
	}
	// Create WmlPackage wrapper, run AfterUnmarshal to gather image parts.
	wmlPkg := parts.NewWmlPackage(pkg)
	wmlPkg.AfterUnmarshal()

	// Store WmlPackage on OpcPackage so any Part can reach it via
	// Package().AppPackage(). This mirrors Python where Package subclasses
	// OpcPackage — every Part._package IS the WML-level Package.
	pkg.SetAppPackage(wmlPkg)

	return newDocument(docPart, wmlPkg)
}
