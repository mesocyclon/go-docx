package docx

import (
	"time"

	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// CoreProperties provides typed access to the Dublin Core and OPC core
// properties of a document (author, title, created date, etc.).
//
// Mirrors Python docx.opc.coreprops.CoreProperties.
type CoreProperties struct {
	ct *oxml.CT_CoreProperties
}

// newCoreProperties creates a CoreProperties proxy wrapping the given
// CT_CoreProperties element. The element must be the shared element owned
// by a CorePropertiesPart — mutations are reflected on save.
//
// Mirrors Python CoreProperties(element).
func newCoreProperties(ct *oxml.CT_CoreProperties) *CoreProperties {
	return &CoreProperties{ct: ct}
}

// --------------------------------------------------------------------------
// Text properties — getters return "" when absent, matching Python behavior
// --------------------------------------------------------------------------

// Author returns the document author (dc:creator), or "".
func (cp *CoreProperties) Author() string { return cp.ct.AuthorText() }

// SetAuthor sets the document author (dc:creator).
func (cp *CoreProperties) SetAuthor(v string) error { return cp.ct.SetAuthorText(v) }

// Title returns the document title (dc:title), or "".
func (cp *CoreProperties) Title() string { return cp.ct.TitleText() }

// SetTitle sets the document title (dc:title).
func (cp *CoreProperties) SetTitle(v string) error { return cp.ct.SetTitleText(v) }

// Subject returns the document subject (dc:subject), or "".
func (cp *CoreProperties) Subject() string { return cp.ct.SubjectText() }

// SetSubject sets the document subject (dc:subject).
func (cp *CoreProperties) SetSubject(v string) error { return cp.ct.SetSubjectText(v) }

// Category returns the document category (cp:category), or "".
func (cp *CoreProperties) Category() string { return cp.ct.CategoryText() }

// SetCategory sets the document category (cp:category).
func (cp *CoreProperties) SetCategory(v string) error { return cp.ct.SetCategoryText(v) }

// Keywords returns the document keywords (cp:keywords), or "".
func (cp *CoreProperties) Keywords() string { return cp.ct.KeywordsText() }

// SetKeywords sets the document keywords (cp:keywords).
func (cp *CoreProperties) SetKeywords(v string) error { return cp.ct.SetKeywordsText(v) }

// Comments returns the document comments/description (dc:description), or "".
func (cp *CoreProperties) Comments() string { return cp.ct.CommentsText() }

// SetComments sets the document comments/description (dc:description).
func (cp *CoreProperties) SetComments(v string) error { return cp.ct.SetCommentsText(v) }

// LastModifiedBy returns the last modifier (cp:lastModifiedBy), or "".
func (cp *CoreProperties) LastModifiedBy() string { return cp.ct.LastModifiedByText() }

// SetLastModifiedBy sets the last modifier (cp:lastModifiedBy).
func (cp *CoreProperties) SetLastModifiedBy(v string) error { return cp.ct.SetLastModifiedByText(v) }

// ContentStatus returns the content status (cp:contentStatus), or "".
func (cp *CoreProperties) ContentStatus() string { return cp.ct.ContentStatusText() }

// SetContentStatus sets the content status (cp:contentStatus).
func (cp *CoreProperties) SetContentStatus(v string) error { return cp.ct.SetContentStatusText(v) }

// Identifier returns the identifier (dc:identifier), or "".
func (cp *CoreProperties) Identifier() string { return cp.ct.IdentifierText() }

// SetIdentifier sets the identifier (dc:identifier).
func (cp *CoreProperties) SetIdentifier(v string) error { return cp.ct.SetIdentifierText(v) }

// Language returns the language (dc:language), or "".
func (cp *CoreProperties) Language() string { return cp.ct.LanguageText() }

// SetLanguage sets the language (dc:language).
func (cp *CoreProperties) SetLanguage(v string) error { return cp.ct.SetLanguageText(v) }

// Version returns the version (cp:version), or "".
func (cp *CoreProperties) Version() string { return cp.ct.VersionText() }

// SetVersion sets the version (cp:version).
func (cp *CoreProperties) SetVersion(v string) error { return cp.ct.SetVersionText(v) }

// --------------------------------------------------------------------------
// Datetime properties — getters return nil when absent
// --------------------------------------------------------------------------

// Created returns the creation time (dcterms:created), or nil.
func (cp *CoreProperties) Created() (*time.Time, error) { return cp.ct.CreatedDatetime() }

// SetCreated sets the creation time (dcterms:created).
func (cp *CoreProperties) SetCreated(t time.Time) { cp.ct.SetCreatedDatetime(t) }

// Modified returns the last modification time (dcterms:modified), or nil.
func (cp *CoreProperties) Modified() (*time.Time, error) { return cp.ct.ModifiedDatetime() }

// SetModified sets the last modification time (dcterms:modified).
func (cp *CoreProperties) SetModified(t time.Time) { cp.ct.SetModifiedDatetime(t) }

// LastPrinted returns the last printed time (cp:lastPrinted), or nil.
func (cp *CoreProperties) LastPrinted() (*time.Time, error) { return cp.ct.LastPrintedDatetime() }

// SetLastPrinted sets the last printed time (cp:lastPrinted).
func (cp *CoreProperties) SetLastPrinted(t time.Time) { cp.ct.SetLastPrintedDatetime(t) }

// --------------------------------------------------------------------------
// Revision
// --------------------------------------------------------------------------

// Revision returns the revision number (cp:revision), or 0 if absent.
func (cp *CoreProperties) Revision() int { return cp.ct.RevisionNumber() }

// SetRevision sets the revision number (cp:revision). Must be >= 1.
func (cp *CoreProperties) SetRevision(v int) error { return cp.ct.SetRevisionNumber(v) }
