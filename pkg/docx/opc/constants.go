// Package opc implements the Open Packaging Conventions (OPC) layer for reading and
// writing ZIP-based OOXML packages. It is independent of WML (Word Markup Language)
// and could be reused for xlsx/pptx.
package opc

import "strings"

// --------------------------------------------------------------------------
// Content Types
// --------------------------------------------------------------------------

const (
	CTBmp                       = "image/bmp"
	CTDmlChart                  = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	CTDmlChartshapes            = "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"
	CTDmlDiagramColors          = "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"
	CTDmlDiagramData            = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
	CTDmlDiagramLayout          = "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
	CTDmlDiagramStyle           = "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
	CTGif                       = "image/gif"
	CTJpeg                      = "image/jpeg"
	CTMsPhoto                   = "image/vnd.ms-photo"
	CTOfcCustomProperties       = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
	CTOfcCustomXmlProperties    = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
	CTOfcDrawing                = "application/vnd.openxmlformats-officedocument.drawing+xml"
	CTOfcExtendedProperties     = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	CTOfcOleObject              = "application/vnd.openxmlformats-officedocument.oleObject"
	CTOfcPackage                = "application/vnd.openxmlformats-officedocument.package"
	CTOfcTheme                  = "application/vnd.openxmlformats-officedocument.theme+xml"
	CTOfcThemeOverride          = "application/vnd.openxmlformats-officedocument.themeOverride+xml"
	CTOfcVmlDrawing             = "application/vnd.openxmlformats-officedocument.vmlDrawing"
	CTOpcCoreProperties         = "application/vnd.openxmlformats-package.core-properties+xml"
	CTOpcDigitalSignatureCert   = "application/vnd.openxmlformats-package.digital-signature-certificate"
	CTOpcDigitalSignatureOrigin = "application/vnd.openxmlformats-package.digital-signature-origin"
	CTOpcDigitalSignatureXmlsig = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
	CTOpcRelationships          = "application/vnd.openxmlformats-package.relationships+xml"
	CTPng                       = "image/png"
	CTTiff                      = "image/tiff"
	CTWmlComments               = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
	CTWmlDocument               = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	CTWmlDocumentGlossary       = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
	CTWmlDocumentMain           = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	CTWmlEndnotes               = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
	CTWmlFontTable              = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
	CTWmlFooter                 = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	CTWmlFootnotes              = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
	CTWmlHeader                 = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	CTWmlNumbering              = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
	CTWmlPrinterSettings        = "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings"
	CTWmlSettings               = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	CTWmlStyles                 = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	CTWmlWebSettings            = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
	CTXml                       = "application/xml"
	CTXEmf                      = "image/x-emf"
	CTXFontdata                 = "application/x-fontdata"
	CTXFontTtf                  = "application/x-font-ttf"
	CTXWmf                      = "image/x-wmf"
	CTSmlSheet                  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	CTSmlSheetMain              = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	CTSmlStyles                 = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	CTSmlWorksheet              = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	CTPmlPresentationMain       = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	CTPmlSlide                  = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	CTPmlSlideLayout            = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	CTPmlSlideMaster            = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	CTPmlPrinterSettings        = "application/vnd.openxmlformats-officedocument.presentationml.printerSettings"
	CTSmlPrinterSettings        = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
)

// --------------------------------------------------------------------------
// OPC Namespaces
// --------------------------------------------------------------------------

const (
	NsOpcContentTypes  = "http://schemas.openxmlformats.org/package/2006/content-types"
	NsOpcRelationships = "http://schemas.openxmlformats.org/package/2006/relationships"
	NsOfcRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// --------------------------------------------------------------------------
// OOXML Strict namespace prefixes
// --------------------------------------------------------------------------
//
// ISO 29500 Strict uses different namespace URIs for relationship types.
// We normalize strict → transitional at read time so that all downstream
// code can compare against the RT* constants without branching.

const (
	// Strict relationship namespace prefix (replaces NsOfcRelationships).
	nsStrictOfcRel = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
	// Transitional relationship namespace prefix.
	nsTransitionalOfcRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
)

// NormalizeRelType converts an OOXML Strict relationship type URI to its
// Transitional equivalent.  Transitional URIs pass through unchanged.
//
// Example:
//
//	"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
//	→ "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
func NormalizeRelType(relType string) string {
	if strings.HasPrefix(relType, nsStrictOfcRel) {
		return nsTransitionalOfcRel + relType[len(nsStrictOfcRel):]
	}
	return relType
}

// --------------------------------------------------------------------------
// Relationship Target Mode
// --------------------------------------------------------------------------

const (
	TargetModeInternal = "Internal"
	TargetModeExternal = "External"
)

// --------------------------------------------------------------------------
// Relationship Types
// --------------------------------------------------------------------------

const (
	RTOfficeDocument     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	RTStyles             = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	RTNumbering          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
	RTSettings           = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	RTComments           = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	RTHeader             = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	RTFooter             = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	RTImage              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RTCoreProperties     = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	RTHyperlink          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	RTFontTable          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	RTTheme              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	RTWebSettings        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	RTEndnotes           = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
	RTFootnotes          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
	RTExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	RTCustomProperties   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	RTGlossaryDocument   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument"
	RTThumbnail          = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
	RTDrawing            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
	RTChart              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	RTCustomXml          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
	RTCustomXmlProps     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
	RTSlide              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	RTSlideLayout        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	RTSlideMaster        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	RTPresProps          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
	RTViewProps          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
	RTTableStyles        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
	RTPrinterSettings    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
	RTVmlDrawing         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
	RTPackage            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
)

// --------------------------------------------------------------------------
// Default content types (extension → content-type mapping)
// --------------------------------------------------------------------------

// defaultContentTypePair represents an (extension, contentType) default mapping.
// Multiple pairs with the same extension are allowed (e.g. "bin" maps to
// printer settings for WML, PML, and SML), matching Python's tuple-of-tuples.
type defaultContentTypePair struct {
	Ext         string
	ContentType string
}

// DefaultContentTypes is the set of default (extension → content-type) mappings
// used when building [Content_Types].xml. Matches Python opc/spec.py.
var DefaultContentTypes = []defaultContentTypePair{
	{"bin", CTPmlPrinterSettings},
	{"bin", CTSmlPrinterSettings},
	{"bin", CTWmlPrinterSettings},
	{"bmp", CTBmp},
	{"emf", CTXEmf},
	{"fntdata", CTXFontdata},
	{"gif", CTGif},
	{"jpe", CTJpeg},
	{"jpeg", CTJpeg},
	{"jpg", CTJpeg},
	{"png", CTPng},
	{"rels", CTOpcRelationships},
	{"tif", CTTiff},
	{"tiff", CTTiff},
	{"wdp", CTMsPhoto},
	{"wmf", CTXWmf},
	{"xlsx", CTSmlSheet},
	{"xml", CTXml},
}

// IsDefaultContentType returns true if (ext, contentType) is in the default set.
// Extension comparison is case-insensitive.
func IsDefaultContentType(ext, contentType string) bool {
	lower := strings.ToLower(ext)
	for _, pair := range DefaultContentTypes {
		if pair.Ext == lower && pair.ContentType == contentType {
			return true
		}
	}
	return false
}
