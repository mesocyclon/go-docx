package docx

import (
	"fmt"
	"io"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Document is the top-level object for a .docx file.
//
// Mirrors Python Document(ElementProxy).
type Document struct {
	element *oxml.CT_Document
	part    *parts.DocumentPart
	wmlPkg  *parts.WmlPackage
	body    *Body // lazy, mirrors Python _body
}

// newDocument creates a Document from its constituent pieces.
func newDocument(docPart *parts.DocumentPart, wmlPkg *parts.WmlPackage) (*Document, error) {
	el := docPart.Element()
	if el == nil {
		return nil, fmt.Errorf("docx: document part element is nil")
	}
	ctDoc := &oxml.CT_Document{Element: oxml.WrapElement(el)}
	return &Document{
		element: ctDoc,
		part:    docPart,
		wmlPkg:  wmlPkg,
	}, nil
}

// --------------------------------------------------------------------------
// Content mutation
// --------------------------------------------------------------------------

// AddHeading appends a heading paragraph to the end of the document.
// Level 0 produces a "Title" style; 1-9 produce "Heading N".
//
// Mirrors Python Document.add_heading.
func (d *Document) AddHeading(text string, level int) (*Paragraph, error) {
	if level < 0 || level > 9 {
		return nil, fmt.Errorf("docx: level must be in range 0-9, got %d", level)
	}
	style := "Title"
	if level > 0 {
		style = fmt.Sprintf("Heading %d", level)
	}
	return d.AddParagraph(text, StyleName(style))
}

// AddPageBreak appends a new paragraph containing only a page break.
//
// Mirrors Python Document.add_page_break.
func (d *Document) AddPageBreak() (*Paragraph, error) {
	para, err := d.AddParagraph("")
	if err != nil {
		return nil, err
	}
	run, err := para.AddRun("")
	if err != nil {
		return nil, err
	}
	if err := run.AddBreak(enum.WdBreakTypePage); err != nil {
		return nil, err
	}
	return para, nil
}

// AddParagraph appends a new paragraph to the end of the document body.
// text may contain tab (\t) and newline (\n, \r) characters. style may
// be a StyleName or a *BaseStyle. Omit to apply no explicit style.
//
// Mirrors Python Document.add_paragraph → self._body.add_paragraph(text, style).
func (d *Document) AddParagraph(text string, style ...StyleRef) (*Paragraph, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, err
	}
	return b.AddParagraph(text, style...)
}

// AddPicture adds an inline picture in its own paragraph at the end of the
// document. The image is read from r. width and height are optional EMU
// values; pass nil for native/proportional sizing.
//
// Mirrors Python Document.add_picture → add_paragraph().add_run().add_picture().
func (d *Document) AddPicture(r io.ReadSeeker, width, height *int64) (*InlineShape, error) {
	para, err := d.AddParagraph("")
	if err != nil {
		return nil, fmt.Errorf("docx: add picture paragraph: %w", err)
	}
	run, err := para.AddRun("")
	if err != nil {
		return nil, fmt.Errorf("docx: add picture run: %w", err)
	}
	return run.AddPicture(r, width, height)
}

// AddSection adds a new section break at the end of the document and returns
// the new Section. startType defaults to WdSectionStartNewPage.
//
// Mirrors Python Document.add_section.
func (d *Document) AddSection(startType enum.WdSectionStart) (*Section, error) {
	body := d.element.Body()
	if body == nil {
		return nil, fmt.Errorf("docx: document has no body")
	}
	newSectPr := body.AddSectionBreak()
	if err := newSectPr.SetStartType(startType); err != nil {
		return nil, fmt.Errorf("docx: setting section start type: %w", err)
	}
	return newSection(newSectPr, d.part), nil
}

// AddTable appends a new table with the given row and column counts.
// style may be a StyleName or *BaseStyle. Omit to apply no explicit style.
//
// Mirrors Python Document.add_table.
func (d *Document) AddTable(rows, cols int, style ...StyleRef) (*Table, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, err
	}
	bw, err := d.blockWidth()
	if err != nil {
		return nil, err
	}
	table, err := b.AddTable(rows, cols, bw)
	if err != nil {
		return nil, err
	}
	if raw := resolveStyleRef(style); raw != nil {
		if err := table.setStyleRaw(raw); err != nil {
			return nil, fmt.Errorf("docx: setting table style: %w", err)
		}
	}
	return table, nil
}

// AddComment adds a comment anchored to the specified runs.
// runs must contain at least one Run; the first and last are used to
// delimit the comment range. text, author, and initials populate the
// comment metadata.
//
// Mirrors Python Document.add_comment.
func (d *Document) AddComment(runs []*Run, text, author string, initials *string) (*Comment, error) {
	if len(runs) == 0 {
		return nil, fmt.Errorf("docx: at least one run required for comment")
	}
	firstRun := runs[0]
	lastRun := runs[len(runs)-1]

	comments, err := d.Comments()
	if err != nil {
		return nil, err
	}
	comment, err := comments.AddComment(text, author, initials)
	if err != nil {
		return nil, err
	}
	commentID, err := comment.CommentID()
	if err != nil {
		return nil, fmt.Errorf("docx: getting comment ID: %w", err)
	}
	if err := firstRun.MarkCommentRange(lastRun, commentID); err != nil {
		return nil, fmt.Errorf("docx: marking comment range: %w", err)
	}
	return comment, nil
}

// ReplaceText replaces all occurrences of old with new throughout the entire
// document: body, headers, footers, and comments of all sections.
//
// Headers/footers without their own definition (linked to previous) are
// skipped. Additionally, already-processed StoryParts are tracked by pointer
// to avoid double replacement when multiple sections share the same
// HeaderPart/FooterPart.
//
// Returns the total number of replacements performed.
func (d *Document) ReplaceText(old, new string) (int, error) {
	if old == "" {
		return 0, nil
	}

	// 1. Document body.
	b, err := d.getBody()
	if err != nil {
		return 0, err
	}
	count := b.ReplaceText(old, new)

	// 2. Headers/footers of all sections, with deduplication.
	seen := map[*parts.StoryPart]bool{}
	for _, sect := range d.Sections().Iter() {
		hfs := []*baseHeaderFooter{
			&sect.Header().baseHeaderFooter,
			&sect.Footer().baseHeaderFooter,
			&sect.EvenPageHeader().baseHeaderFooter,
			&sect.EvenPageFooter().baseHeaderFooter,
			&sect.FirstPageHeader().baseHeaderFooter,
			&sect.FirstPageFooter().baseHeaderFooter,
		}
		for _, hf := range hfs {
			n, err := hf.replaceTextDedup(old, new, seen)
			if err != nil {
				return count, fmt.Errorf("docx: replacing text in %s: %w", hf.ops.kind(), err)
			}
			count += n
		}
	}

	// 3. Comments.
	n, err := d.replaceTextInComments(old, new)
	if err != nil {
		return count, err
	}
	count += n

	return count, nil
}

// replaceTextInComments replaces text in all comments. Returns 0 if
// no comments part exists (avoids creating one as a side effect).
func (d *Document) replaceTextInComments(old, new string) (int, error) {
	if !d.part.HasCommentsPart() {
		return 0, nil
	}
	comments, err := d.Comments()
	if err != nil {
		return 0, fmt.Errorf("docx: replacing text in comments: %w", err)
	}
	return comments.ReplaceText(old, new), nil
}

// --------------------------------------------------------------------------
// Properties
// --------------------------------------------------------------------------

// Comments returns the Comments collection for this document.
//
// Mirrors Python Document.comments → self._part.comments.
func (d *Document) Comments() (*Comments, error) {
	cp, err := d.part.CommentsPart()
	if err != nil {
		return nil, fmt.Errorf("docx: getting comments part: %w", err)
	}
	elm, err := cp.CommentsElement()
	if err != nil {
		return nil, fmt.Errorf("docx: getting comments element: %w", err)
	}
	return newComments(elm, cp), nil
}

// CoreProperties returns the CoreProperties for this document.
//
// Mirrors Python Document.core_properties → self._part.core_properties.
func (d *Document) CoreProperties() (*CoreProperties, error) {
	cpp, err := d.part.CoreProperties()
	if err != nil {
		return nil, fmt.Errorf("docx: getting core properties: %w", err)
	}
	elm, err := cpp.CT()
	if err != nil {
		return nil, fmt.Errorf("docx: getting core properties element: %w", err)
	}
	return newCoreProperties(elm), nil
}

// InlineShapes returns the InlineShapes collection for this document.
//
// Mirrors Python Document.inline_shapes → self._part.inline_shapes.
func (d *Document) InlineShapes() (*InlineShapes, error) {
	body := d.element.Body()
	if body == nil || body.RawElement() == nil {
		return nil, fmt.Errorf("docx: document has no body element")
	}
	return newInlineShapes(body.RawElement(), &d.part.StoryPart), nil
}

// IterInnerContent returns all paragraphs and tables in document order.
//
// Mirrors Python Document.iter_inner_content → self._body.iter_inner_content().
func (d *Document) IterInnerContent() ([]*InnerContentItem, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, fmt.Errorf("docx: getting body: %w", err)
	}
	return b.IterInnerContent(), nil
}

// Paragraphs returns all top-level paragraphs in document order.
//
// Mirrors Python Document.paragraphs → self._body.paragraphs.
func (d *Document) Paragraphs() ([]*Paragraph, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, fmt.Errorf("docx: getting body: %w", err)
	}
	return b.Paragraphs(), nil
}

// Part returns the DocumentPart for this document.
//
// Mirrors Python Document.part.
func (d *Document) Part() *parts.DocumentPart {
	return d.part
}

// Sections returns the Sections collection for this document.
//
// Mirrors Python Document.sections → Sections(self._element, self._part).
func (d *Document) Sections() *Sections {
	return newSections(d.element, d.part)
}

// Settings returns the Settings proxy for this document.
//
// Mirrors Python Document.settings → self._part.settings.
func (d *Document) Settings() (*Settings, error) {
	elm, err := d.part.Settings()
	if err != nil {
		return nil, fmt.Errorf("docx: getting settings: %w", err)
	}
	return newSettings(elm), nil
}

// Styles returns the Styles proxy for this document.
//
// Mirrors Python Document.styles → self._part.styles.
func (d *Document) Styles() (*Styles, error) {
	elm, err := d.part.Styles()
	if err != nil {
		return nil, fmt.Errorf("docx: getting styles: %w", err)
	}
	return newStyles(elm), nil
}

// Tables returns all top-level tables in document order.
//
// Mirrors Python Document.tables → self._body.tables.
func (d *Document) Tables() ([]*Table, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, fmt.Errorf("docx: getting body: %w", err)
	}
	return b.Tables(), nil
}

// --------------------------------------------------------------------------
// Save
// --------------------------------------------------------------------------

// Save writes this document to w.
//
// Mirrors Python Document.save(stream).
func (d *Document) Save(w io.Writer) error {
	return d.wmlPkg.Save(w)
}

// SaveFile writes this document to a file.
//
// Mirrors Python Document.save(path).
func (d *Document) SaveFile(path string) error {
	return d.wmlPkg.SaveToFile(path)
}

// --------------------------------------------------------------------------
// Internal
// --------------------------------------------------------------------------

// blockWidth returns the available width between margins of the last section,
// in twips. Used for table column width calculation.
//
// Delegates to sectionBlockWidth to avoid duplicating the width computation
// algorithm. See sectionBlockWidth for details and defaults.
//
// Mirrors Python Document._block_width (but in twips, not EMU, since Go
// Section methods return twips).
func (d *Document) blockWidth() (int, error) {
	sections := d.Sections()
	if sections.Len() == 0 {
		return Inches(6.5).Twips(), nil
	}
	last, err := sections.Get(sections.Len() - 1)
	if err != nil {
		return 0, fmt.Errorf("docx: getting last section: %w", err)
	}
	return sectionBlockWidth(last), nil
}

// getBody returns the cached Body, creating it on first call.
//
// Mirrors Python Document._body (lazy property).
func (d *Document) getBody() (*Body, error) {
	if d.body != nil {
		return d.body, nil
	}
	bodyElm := d.element.Body()
	if bodyElm == nil {
		return nil, fmt.Errorf("docx: document has no body element")
	}
	d.body = newBody(bodyElm, &d.part.StoryPart)
	return d.body, nil
}

// --------------------------------------------------------------------------
// Phase 5: ReplaceWithTable — public API + helpers
// --------------------------------------------------------------------------

// ReplaceWithTable replaces all occurrences of the text tag old with a table
// described by td. Processes body, all headers/footers, and comments.
//
// Table width is computed automatically from context:
//   - body: page width − margins (from the last section)
//   - table cell: tcPr/tcW cell width
//   - header/footer: page width − margins (from the owning section)
//   - comment: default 9360 twips (Inches(6.5))
//
// Returns the number of replacements performed.
//
// On error, the document may be partially modified. The returned count
// reflects replacements completed before the error.
func (d *Document) ReplaceWithTable(old string, td TableData) (int, error) {
	if old == "" {
		return 0, nil
	}

	// Defensive copy once at the top level. Sub-calls (body, headers,
	// comments) receive the already-copied td — no per-container copying.
	td = td.defensiveCopy()

	// 1. Document body.
	bw, err := d.blockWidth()
	if err != nil {
		return 0, err
	}
	b, err := d.getBody()
	if err != nil {
		return 0, err
	}
	count, err := b.replaceWithTable(old, td, bw)
	if err != nil {
		return count, err
	}

	// 2. Headers/footers of all sections, with deduplication.
	seen := map[*parts.StoryPart]bool{}
	for _, sect := range d.Sections().Iter() {
		sectWidth := sectionBlockWidth(sect)
		hfs := []*baseHeaderFooter{
			&sect.Header().baseHeaderFooter,
			&sect.Footer().baseHeaderFooter,
			&sect.EvenPageHeader().baseHeaderFooter,
			&sect.EvenPageFooter().baseHeaderFooter,
			&sect.FirstPageHeader().baseHeaderFooter,
			&sect.FirstPageFooter().baseHeaderFooter,
		}
		for _, hf := range hfs {
			n, err := hf.replaceWithTableDedup(old, td, sectWidth, seen)
			if err != nil {
				return count, fmt.Errorf("docx: replacing with table in %s: %w", hf.ops.kind(), err)
			}
			count += n
		}
	}

	// 3. Comments.
	n, err := d.replaceWithTableInComments(old, td)
	if err != nil {
		return count, err
	}
	count += n

	return count, nil
}

// sectionBlockWidth computes the content width (page width minus margins) for
// a section. Falls back to US Letter defaults (9360 twips = Inches(6.5)).
// If margins exceed page width (anomalous metadata), clamps to the default.
func sectionBlockWidth(sect *Section) int {
	pageWidth := Inches(8.5).Twips()
	if pw, err := sect.PageWidth(); err == nil && pw != nil {
		pageWidth = *pw
	}
	leftMargin := Inches(1).Twips()
	if lm, err := sect.LeftMargin(); err == nil && lm != nil {
		leftMargin = *lm
	}
	rightMargin := Inches(1).Twips()
	if rm, err := sect.RightMargin(); err == nil && rm != nil {
		rightMargin = *rm
	}
	w := pageWidth - leftMargin - rightMargin
	if w <= 0 {
		w = Inches(6.5).Twips() // fallback for anomalous metadata
	}
	return w
}

// replaceWithTableInComments replaces tags with tables in all comments.
// Returns 0 if no comments part exists (avoids creating one as a side effect).
// Mirrors replaceTextInComments (document.go:228).
func (d *Document) replaceWithTableInComments(old string, td TableData) (int, error) {
	if !d.part.HasCommentsPart() {
		return 0, nil
	}
	comments, err := d.Comments()
	if err != nil {
		return 0, fmt.Errorf("docx: replacing with table in comments: %w", err)
	}
	// Comment width: use default (no page layout info in comments).
	defaultWidth := Inches(6.5).Twips()
	count := 0
	for _, c := range comments.Iter() {
		// Comment embeds BlockItemContainer (comments.go:138).
		n, err := c.BlockItemContainer.replaceWithTable(old, td, defaultWidth)
		if err != nil {
			return count, err
		}
		count += n
	}
	return count, nil
}

// --------------------------------------------------------------------------
// Body snapshot — rollback protection for ReplaceWithContent
// --------------------------------------------------------------------------

// bodySnapshot captures the state of the <w:body> element before mutation,
// allowing rollback if an error occurs during content replacement.
//
// The snapshot is a deep copy of the entire body element tree via
// etree.Element.Copy(). This is the same mechanism used throughout
// the project for element cloning.
//
// Limitations (documented for callers):
//   - Orphan parts: if an error occurs after AddPart() for images or
//     generic parts, those parts remain in the OpcPackage (RemovePart
//     does not exist). Orphan parts are harmless — Word ignores them.
//   - Styles/numbering: resources imported in Phase 1 are not rolled back.
//     Unused styles and numbering definitions are harmless.
//   - For full rollback including orphan parts, callers may use the
//     save-before / reopen-on-error pattern:
//
//     buf := new(bytes.Buffer)
//     doc.Save(buf)
//     count, err := doc.ReplaceWithContent(tag, cd)
//     if err != nil {
//         doc, _ = OpenBytes(buf.Bytes())
//     }
type bodySnapshot struct {
	bodyEl *etree.Element // deep copy of <w:body>
}

// snapshotBody creates a deep copy of the current <w:body> element.
func (d *Document) snapshotBody() (*bodySnapshot, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, err
	}
	return &bodySnapshot{bodyEl: b.Element().Copy()}, nil
}

// restoreBody replaces the current <w:body> with the snapshot copy,
// effectively undoing any mutations applied after the snapshot was taken.
//
// API chain (verified):
//
//	d.element         → *oxml.CT_Document (embeds oxml.Element)
//	d.element.RawElement()  → *etree.Element  (<w:document>)
//	d.element.Body()        → *oxml.CT_Body
//	body.RawElement()       → *etree.Element  (<w:body>)
func (d *Document) restoreBody(snap *bodySnapshot) {
	docEl := d.element.RawElement() // <w:document> etree element
	oldBody := d.element.Body().RawElement()
	docEl.RemoveChild(oldBody)
	docEl.AddChild(snap.bodyEl)
	d.body = nil // invalidate cached Body proxy
}

// --------------------------------------------------------------------------
// Phase 5: ReplaceWithContent — public API + helpers
// --------------------------------------------------------------------------

// ReplaceWithContent replaces all occurrences of the text tag old with the
// body content of the source document cd.Source. Processes body, all
// headers/footers, and comments of the target document.
//
// Only block-level elements (paragraphs, tables) from the source body are
// inserted. Headers, footers, and section properties of the source are
// excluded.
//
// Images and external hyperlinks from the source are imported into the
// target document with relationship remapping. Image parts are
// deduplicated via SHA-256 hash.
//
// Numbering definitions (lists) referenced by the source content are
// imported into the target with fresh IDs. Styles (paragraph, character,
// table) are merged using UseDestinationStyles strategy: if a style
// exists in the target, the target definition is used; otherwise, the
// source definition is deep-copied including its dependency chain
// (basedOn, next, link). Footnotes and endnotes from the source are
// imported with fresh IDs; their bodies are processed for styles,
// numbering, and relationships.
//
// Returns the number of replacements performed.
//
// On error, the document body is rolled back to its pre-call state via an
// internal snapshot. Resources imported in Phase 1 (styles, numbering,
// footnotes, endnotes) and any orphan parts are not rolled back but are
// harmless — Word ignores unused definitions.
func (d *Document) ReplaceWithContent(old string, cd ContentData) (count int, err error) {
	if old == "" {
		return 0, nil
	}
	if cd.Source == nil {
		return 0, fmt.Errorf("docx: ContentData.Source is nil")
	}

	// Snapshot body before any mutation so we can restore on error.
	snap, snapErr := d.snapshotBody()
	if snapErr != nil {
		return 0, fmt.Errorf("docx: creating body snapshot: %w", snapErr)
	}
	defer func() {
		if err != nil && snap != nil {
			d.restoreBody(snap)
		}
	}()

	// ResourceImporter coordinates resource transfer for this call.
	// Shared across body, headers, footers, and comments so that styles,
	// numbering, footnotes, and part blobs are imported exactly once.
	ri := newResourceImporter(cd.Source, d, d.wmlPkg, cd.Format, cd.Options)

	// Phase 1: import resources from source (once).
	// Numbering must run before styles (styles may reference numId).
	if err := ri.importNumbering(); err != nil {
		return 0, fmt.Errorf("docx: importing numbering: %w", err)
	}
	if err := ri.importStyles(); err != nil {
		return 0, fmt.Errorf("docx: importing styles: %w", err)
	}
	if err := ri.importFootnotes(); err != nil {
		return 0, fmt.Errorf("docx: importing footnotes: %w", err)
	}
	if err := ri.importEndnotes(); err != nil {
		return 0, fmt.Errorf("docx: importing endnotes: %w", err)
	}

	// Prepare content once — remap relationships from source → target body part.
	bodyPrep, err := prepareContentElements(cd.Source, &d.part.StoryPart, ri)
	if err != nil {
		return 0, fmt.Errorf("docx: preparing content for body: %w", err)
	}

	// 1. Document body.
	bw, err := d.blockWidth()
	if err != nil {
		return 0, err
	}
	b, err := d.getBody()
	if err != nil {
		return 0, err
	}
	count, err = b.replaceWithContent(old, bodyPrep, bw)
	if err != nil {
		return count, err
	}

	// 2. Headers/footers of all sections, with deduplication.
	seen := map[*parts.StoryPart]bool{}
	for _, sect := range d.Sections().Iter() {
		sectWidth := sectionBlockWidth(sect)
		hfs := []*baseHeaderFooter{
			&sect.Header().baseHeaderFooter,
			&sect.Footer().baseHeaderFooter,
			&sect.EvenPageHeader().baseHeaderFooter,
			&sect.EvenPageFooter().baseHeaderFooter,
			&sect.FirstPageHeader().baseHeaderFooter,
			&sect.FirstPageFooter().baseHeaderFooter,
		}
		for _, hf := range hfs {
			n, err := hf.replaceWithContentDedup(old, cd.Source, ri, sectWidth, seen)
			if err != nil {
				return count, fmt.Errorf("docx: replacing with content in %s: %w",
					hf.ops.kind(), err)
			}
			count += n
		}
	}

	// 3. Comments.
	n, err := d.replaceWithContentInComments(old, cd.Source, ri)
	if err != nil {
		return count, err
	}
	count += n

	return count, nil
}

// replaceWithContentInComments replaces tags with source document content in
// all comments. Returns 0 if no comments part exists (avoids creating one as
// a side effect). Mirrors replaceWithTableInComments.
func (d *Document) replaceWithContentInComments(
	old string,
	sourceDoc *Document,
	ri *ResourceImporter,
) (int, error) {
	if !d.part.HasCommentsPart() {
		return 0, nil
	}
	comments, err := d.Comments()
	if err != nil {
		return 0, fmt.Errorf("docx: replacing with content in comments: %w", err)
	}
	defaultWidth := Inches(6.5).Twips()
	count := 0
	for _, c := range comments.Iter() {
		prep, err := prepareContentElements(sourceDoc, c.BlockItemContainer.part, ri)
		if err != nil {
			return count, err
		}
		n, err := c.BlockItemContainer.replaceWithContent(old, prep, defaultWidth)
		if err != nil {
			return count, err
		}
		count += n
	}
	return count, nil
}
