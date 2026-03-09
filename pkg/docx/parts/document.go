package parts

import (
	"fmt"
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// DocumentPart is the main document part of a WordprocessingML package.
// It acts as broker to other parts such as image, core properties, and
// style parts. It also acts as a convenient delegate when a mid-document
// object needs a service involving a remote ancestor.
//
// Mirrors Python DocumentPart(StoryPart).
type DocumentPart struct {
	StoryPart
	// numberingPart is the only lazyproperty-cached part in Python.
	// _styles_part, _settings_part, _comments_part are @property (not
	// lazyproperty) in Python — they re-check the relationship each call.
	// The relationship graph itself acts as the cache.
	numberingPart *NumberingPart

	// Bookmark ID counter — unique per document (not per story part).
	lastBookmarkID      int
	bookmarkIDScanned   bool
	bookmarkNameCounter int // monotonic suffix counter for name dedup
}

// NewDocumentPart creates a DocumentPart wrapping the given XmlPart.
func NewDocumentPart(xp *opc.XmlPart) *DocumentPart {
	dp := &DocumentPart{
		StoryPart: StoryPart{XmlPart: xp},
	}
	// The document part is its own document part.
	dp.StoryPart.SetDocumentPart(dp)
	return dp
}

// --------------------------------------------------------------------------
// Body access
// --------------------------------------------------------------------------

// Body returns the CT_Body element of this document.
func (dp *DocumentPart) Body() (*oxml.CT_Body, error) {
	el := dp.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: document element is nil")
	}
	doc := &oxml.CT_Document{Element: oxml.WrapElement(el)}
	body := doc.Body()
	if body == nil {
		return nil, fmt.Errorf("parts: document has no body element")
	}
	return body, nil
}

// --------------------------------------------------------------------------
// Bookmark ID allocation
// --------------------------------------------------------------------------

// NextBookmarkID returns a fresh document-unique bookmark ID.
// On the first call it scans ALL story parts (body, headers, footers,
// comments) for the maximum existing w:id on bookmarkStart/bookmarkEnd
// elements. Subsequent calls increment the cached counter.
func (dp *DocumentPart) NextBookmarkID() int {
	if !dp.bookmarkIDScanned {
		dp.lastBookmarkID = dp.scanMaxBookmarkID()
		dp.bookmarkIDScanned = true
	}
	dp.lastBookmarkID++
	return dp.lastBookmarkID
}

// NextBookmarkNameSuffix returns a monotonically increasing suffix number
// for bookmark name deduplication (e.g. "_imp1", "_imp2"). Document-scoped
// to guarantee uniqueness across multiple renumberBookmarks calls.
func (dp *DocumentPart) NextBookmarkNameSuffix() int {
	dp.bookmarkNameCounter++
	return dp.bookmarkNameCounter
}

// scanMaxBookmarkID finds the maximum w:id value on bookmarkStart and
// bookmarkEnd elements across all story parts of this document.
func (dp *DocumentPart) scanMaxBookmarkID() int {
	maxID := 0

	// 1. Main document body.
	if el := dp.Element(); el != nil {
		if v := collectMaxBookmarkID(el); v > maxID {
			maxID = v
		}
	}

	// 2. All headers.
	for _, rel := range dp.Rels().AllByRelType(opc.RTHeader) {
		if hp, ok := rel.TargetPart.(*HeaderPart); ok {
			if el := hp.Element(); el != nil {
				if v := collectMaxBookmarkID(el); v > maxID {
					maxID = v
				}
			}
		}
	}

	// 3. All footers.
	for _, rel := range dp.Rels().AllByRelType(opc.RTFooter) {
		if fp, ok := rel.TargetPart.(*FooterPart); ok {
			if el := fp.Element(); el != nil {
				if v := collectMaxBookmarkID(el); v > maxID {
					maxID = v
				}
			}
		}
	}

	// 4. Comments.
	rel, err := dp.Rels().GetByRelType(opc.RTComments)
	if err == nil && rel.TargetPart != nil {
		if cp, ok := rel.TargetPart.(*CommentsPart); ok {
			if el := cp.Element(); el != nil {
				if v := collectMaxBookmarkID(el); v > maxID {
					maxID = v
				}
			}
		}
	}

	return maxID
}

// collectMaxBookmarkID scans an element tree for the maximum w:id value
// on bookmarkStart and bookmarkEnd elements.
func collectMaxBookmarkID(root *etree.Element) int {
	maxID := 0
	stack := []*etree.Element{root}
	for len(stack) > 0 {
		el := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if el.Space == "w" && (el.Tag == "bookmarkStart" || el.Tag == "bookmarkEnd") {
			v := el.SelectAttrValue("w:id", "")
			if v != "" {
				if id, err := strconv.Atoi(v); err == nil && id > maxID {
					maxID = id
				}
			}
		}
		stack = append(stack, el.ChildElements()...)
	}
	return maxID
}

// --------------------------------------------------------------------------
// Header / Footer
// --------------------------------------------------------------------------

// AddHeaderPart creates a new header part, relates it to this document part,
// and returns the header part and its relationship ID.
//
// Mirrors Python DocumentPart.add_header_part:
//
//	header_part = HeaderPart.new(self.package)
//	rId = self.relate_to(header_part, RT.HEADER)
//	return header_part, rId
func (dp *DocumentPart) AddHeaderPart() (*HeaderPart, string, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, "", fmt.Errorf("parts: document part has no package")
	}
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		return nil, "", fmt.Errorf("parts: creating header part: %w", err)
	}
	// Go-specific: add to parts map so NextPartname sees it for subsequent calls.
	// Python discovers parts via relationship graph traversal in next_partname.
	pkg.AddPart(hp)
	rel := dp.Rels().GetOrAdd(opc.RTHeader, hp)
	return hp, rel.RID, nil
}

// AddFooterPart creates a new footer part, relates it to this document part,
// and returns the footer part and its relationship ID.
//
// Mirrors Python DocumentPart.add_footer_part.
func (dp *DocumentPart) AddFooterPart() (*FooterPart, string, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, "", fmt.Errorf("parts: document part has no package")
	}
	fp, err := NewFooterPart(pkg)
	if err != nil {
		return nil, "", fmt.Errorf("parts: creating footer part: %w", err)
	}
	pkg.AddPart(fp)
	rel := dp.Rels().GetOrAdd(opc.RTFooter, fp)
	return fp, rel.RID, nil
}

// DropHeaderPart removes the header part relationship identified by rId.
// Uses reference-count-aware deletion matching Python DocumentPart.drop_header_part.
//
// Mirrors Python: self.drop_rel(rId)
func (dp *DocumentPart) DropHeaderPart(rId string) {
	dp.DropRel(rId)
}

// HeaderPartByRID returns the HeaderPart related by the given rId.
//
// Mirrors Python DocumentPart.header_part(rId) → self.related_parts[rId].
func (dp *DocumentPart) HeaderPartByRID(rId string) (*HeaderPart, error) {
	rel := dp.Rels().GetByRID(rId)
	if rel == nil {
		return nil, fmt.Errorf("parts: no relationship %q", rId)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: relationship %q has no target part", rId)
	}
	hp, ok := rel.TargetPart.(*HeaderPart)
	if !ok {
		return nil, fmt.Errorf("parts: relationship %q target is %T, want *HeaderPart", rId, rel.TargetPart)
	}
	return hp, nil
}

// FooterPartByRID returns the FooterPart related by the given rId.
//
// Mirrors Python DocumentPart.footer_part(rId) → self.related_parts[rId].
func (dp *DocumentPart) FooterPartByRID(rId string) (*FooterPart, error) {
	rel := dp.Rels().GetByRID(rId)
	if rel == nil {
		return nil, fmt.Errorf("parts: no relationship %q", rId)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: relationship %q has no target part", rId)
	}
	fp, ok := rel.TargetPart.(*FooterPart)
	if !ok {
		return nil, fmt.Errorf("parts: relationship %q target is %T, want *FooterPart", rId, rel.TargetPart)
	}
	return fp, nil
}

// --------------------------------------------------------------------------
// StylesPart — @property in Python (NOT lazyproperty), re-checks each call
// --------------------------------------------------------------------------

// StylesPart returns the StylesPart for this document, creating a default
// one if not present. NOT cached in a struct field — the relationship graph
// acts as the cache, matching Python's @property behavior.
//
// Mirrors Python DocumentPart._styles_part property.
func (dp *DocumentPart) StylesPart() (*StylesPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTStyles)
	if err == nil && rel.TargetPart != nil {
		if sp, ok := rel.TargetPart.(*StylesPart); ok {
			return sp, nil
		}
	}
	// Not found — create default, relate, return.
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	sp, err := DefaultStylesPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default styles part: %w", err)
	}
	pkg.AddPart(sp)
	dp.Rels().GetOrAdd(opc.RTStyles, sp)
	return sp, nil
}

// Styles returns the CT_Styles element from the styles part.
//
// Mirrors Python DocumentPart.styles → self._styles_part.styles.
func (dp *DocumentPart) Styles() (*oxml.CT_Styles, error) {
	sp, err := dp.StylesPart()
	if err != nil {
		return nil, err
	}
	return sp.Styles()
}

// --------------------------------------------------------------------------
// NumberingPart — @lazyproperty in Python (the ONLY cached one)
// --------------------------------------------------------------------------

// NumberingPart returns the NumberingPart for this document. Unlike styles
// and settings, numbering does not auto-create a default part in the Python
// source (NumberingPart.new() raises NotImplementedError). We only resolve
// existing relationships. Cached per Python lazyproperty.
//
// Mirrors Python DocumentPart.numbering_part (lazyproperty).
func (dp *DocumentPart) NumberingPart() (*NumberingPart, error) {
	if dp.numberingPart != nil {
		return dp.numberingPart, nil
	}
	rel, err := dp.Rels().GetByRelType(opc.RTNumbering)
	if err != nil {
		return nil, fmt.Errorf("parts: no numbering part: %w", err)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: numbering relationship has no target part")
	}
	np, ok := rel.TargetPart.(*NumberingPart)
	if !ok {
		return nil, fmt.Errorf("parts: numbering target is %T, want *NumberingPart", rel.TargetPart)
	}
	dp.numberingPart = np
	return np, nil
}

// GetOrAddNumberingPart returns the NumberingPart for this document,
// creating a default empty one if not present. This is used during
// numbering import to ensure the target document has a numbering part.
func (dp *DocumentPart) GetOrAddNumberingPart() (*NumberingPart, error) {
	if np, err := dp.NumberingPart(); err == nil {
		return np, nil
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default numbering part: %w", err)
	}
	pkg.AddPart(np)
	dp.Rels().GetOrAdd(opc.RTNumbering, np)
	dp.numberingPart = np
	return np, nil
}

// --------------------------------------------------------------------------
// FootnotesPart
// --------------------------------------------------------------------------

// FootnotesPart returns the FootnotesPart for this document, or an error
// if no footnotes part exists. Does NOT auto-create a default.
func (dp *DocumentPart) FootnotesPart() (*FootnotesPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTFootnotes)
	if err != nil {
		return nil, fmt.Errorf("parts: no footnotes part: %w", err)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: footnotes relationship has no target part")
	}
	fp, ok := rel.TargetPart.(*FootnotesPart)
	if !ok {
		return nil, fmt.Errorf("parts: footnotes target is %T, want *FootnotesPart", rel.TargetPart)
	}
	return fp, nil
}

// GetOrAddFootnotesPart returns the FootnotesPart for this document,
// creating a default one (with separator/continuationSeparator) if not
// present. Used during footnote import.
func (dp *DocumentPart) GetOrAddFootnotesPart() (*FootnotesPart, error) {
	if fp, err := dp.FootnotesPart(); err == nil {
		return fp, nil
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	fp, err := DefaultFootnotesPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default footnotes part: %w", err)
	}
	pkg.AddPart(fp)
	dp.Rels().GetOrAdd(opc.RTFootnotes, fp)
	return fp, nil
}

// --------------------------------------------------------------------------
// EndnotesPart
// --------------------------------------------------------------------------

// EndnotesPart returns the EndnotesPart for this document, or an error
// if no endnotes part exists. Does NOT auto-create a default.
func (dp *DocumentPart) EndnotesPart() (*EndnotesPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTEndnotes)
	if err != nil {
		return nil, fmt.Errorf("parts: no endnotes part: %w", err)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: endnotes relationship has no target part")
	}
	ep, ok := rel.TargetPart.(*EndnotesPart)
	if !ok {
		return nil, fmt.Errorf("parts: endnotes target is %T, want *EndnotesPart", rel.TargetPart)
	}
	return ep, nil
}

// GetOrAddEndnotesPart returns the EndnotesPart for this document,
// creating a default one (with separator/continuationSeparator) if not
// present. Used during endnote import.
func (dp *DocumentPart) GetOrAddEndnotesPart() (*EndnotesPart, error) {
	if ep, err := dp.EndnotesPart(); err == nil {
		return ep, nil
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	ep, err := DefaultEndnotesPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default endnotes part: %w", err)
	}
	pkg.AddPart(ep)
	dp.Rels().GetOrAdd(opc.RTEndnotes, ep)
	return ep, nil
}

// --------------------------------------------------------------------------
// SettingsPart — @property in Python (NOT lazyproperty)
// --------------------------------------------------------------------------

// SettingsPart returns the SettingsPart for this document, creating a
// default one if not present.
//
// Mirrors Python DocumentPart._settings_part property.
func (dp *DocumentPart) SettingsPart() (*SettingsPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTSettings)
	if err == nil && rel.TargetPart != nil {
		if sp, ok := rel.TargetPart.(*SettingsPart); ok {
			return sp, nil
		}
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	sp, err := DefaultSettingsPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default settings part: %w", err)
	}
	pkg.AddPart(sp)
	dp.Rels().GetOrAdd(opc.RTSettings, sp)
	return sp, nil
}

// Settings returns the CT_Settings element from the settings part.
//
// Mirrors Python DocumentPart.settings → self._settings_part.settings.
func (dp *DocumentPart) Settings() (*oxml.CT_Settings, error) {
	sp, err := dp.SettingsPart()
	if err != nil {
		return nil, err
	}
	return sp.SettingsElement()
}

// --------------------------------------------------------------------------
// CommentsPart — @property in Python (NOT lazyproperty)
// --------------------------------------------------------------------------

// HasCommentsPart reports whether a comments part already exists.
// Unlike CommentsPart(), this does not create one if absent.
func (dp *DocumentPart) HasCommentsPart() bool {
	rel, err := dp.Rels().GetByRelType(opc.RTComments)
	return err == nil && rel.TargetPart != nil
}

// CommentsPart returns the CommentsPart for this document, creating a
// default one if not present.
//
// Mirrors Python DocumentPart._comments_part property.
func (dp *DocumentPart) CommentsPart() (*CommentsPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTComments)
	if err == nil && rel.TargetPart != nil {
		if cp, ok := rel.TargetPart.(*CommentsPart); ok {
			return cp, nil
		}
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	cp, err := DefaultCommentsPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default comments part: %w", err)
	}
	pkg.AddPart(cp)
	dp.Rels().GetOrAdd(opc.RTComments, cp)
	return cp, nil
}

// CommentsElement returns the CT_Comments element from the comments part.
//
// Mirrors Python DocumentPart.comments (element access portion — the domain
// Comments object is MR-11).
func (dp *DocumentPart) CommentsElement() (*oxml.CT_Comments, error) {
	cp, err := dp.CommentsPart()
	if err != nil {
		return nil, err
	}
	return cp.CommentsElement()
}

// --------------------------------------------------------------------------
// CoreProperties
// --------------------------------------------------------------------------

// CoreProperties returns the CorePropertiesPart for this document. If the
// package has no core properties part, a default one is created and related.
//
// Mirrors Python Package._core_properties_part (lazy creation).
func (dp *DocumentPart) CoreProperties() (*CorePropertiesPart, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}

	part, err := pkg.RelatedPart(opc.RTCoreProperties)
	if err == nil {
		cpp, ok := part.(*CorePropertiesPart)
		if ok {
			return cpp, nil
		}
		// Part exists but was loaded as wrong type (shouldn't happen with factory)
		return nil, fmt.Errorf("parts: core properties part is %T, expected *CorePropertiesPart", part)
	}

	// Not found — create default and relate to package
	// Mirrors Python: self.relate_to(core_properties_part, RT.CORE_PROPERTIES)
	cpp, err := DefaultCorePropertiesPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default core properties: %w", err)
	}
	pkg.RelateTo(cpp, opc.RTCoreProperties)
	return cpp, nil
}

// --------------------------------------------------------------------------
// Style delegation
// --------------------------------------------------------------------------

// GetStyle returns the style matching styleID and styleType.
// If styleID is nil, the default style for styleType is returned.
// If styleID does not match a defined style of styleType, the default
// style for styleType is returned.
//
// Mirrors Python DocumentPart.get_style → self.styles.get_by_id(style_id, style_type).
//
// Python Styles._get_by_id:
//
//	style = self._element.get_by_id(style_id)
//	if style is None or style.type != style_type:
//	    return self.default(style_type)
func (dp *DocumentPart) GetStyle(styleID *string, styleType enum.WdStyleType) (*oxml.CT_Style, error) {
	ss, err := dp.Styles()
	if err != nil {
		return nil, err
	}
	if styleID == nil {
		return ss.DefaultFor(styleType)
	}
	s := ss.GetByID(*styleID)
	xmlType, err := styleType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("parts: invalid style type: %w", err)
	}
	if s == nil || s.Type() != xmlType {
		// Fall back to default for the style type (matches Python _get_by_id).
		return ss.DefaultFor(styleType)
	}
	return s, nil
}

// styledObject is satisfied by domain-level style objects (e.g. docx.BaseStyle).
// Standard Go consumer-side interface — parts doesn't import docx.
type styledObject interface {
	StyleID() string
	Type() (enum.WdStyleType, error)
}

// GetStyleID returns the style_id string for styleOrName of styleType.
//
// styleOrName must be one of:
//   - string: a style UI name (e.g. "Heading 1"), resolved via BabelFish
//   - styledObject: any value implementing StyleID() and Type() (e.g. docx.BaseStyle)
//   - nil: returns nil (inherit default style)
//
// Returns nil when the style resolves to the default style for styleType.
//
// Mirrors Python DocumentPart.get_style_id → self.styles.get_style_id.
func (dp *DocumentPart) GetStyleID(styleOrName any, styleType enum.WdStyleType) (*string, error) {
	if styleOrName == nil {
		return nil, nil
	}
	switch v := styleOrName.(type) {
	case string:
		ss, err := dp.Styles()
		if err != nil {
			return nil, err
		}
		return ss.GetStyleIDByName(v, styleType)
	case styledObject:
		// Validate type (Python: _get_style_id_from_style raises ValueError).
		st, err := v.Type()
		if err != nil {
			return nil, err
		}
		if st != styleType {
			return nil, fmt.Errorf("parts: assigned style is type %v, need type %v", st, styleType)
		}
		// Default check (Python: if style == self.default(style_type): return None).
		ss, err := dp.Styles()
		if err != nil {
			return nil, err
		}
		def, err := ss.DefaultFor(styleType)
		if err != nil {
			return nil, err
		}
		if def != nil && def.StyleId() == v.StyleID() {
			return nil, nil
		}
		id := v.StyleID()
		return &id, nil
	default:
		return nil, fmt.Errorf("parts: GetStyleID expects string, style object, or nil, got %T", styleOrName)
	}
}

// --------------------------------------------------------------------------
// InlineShapes (element access only — domain object is MR-11)
// --------------------------------------------------------------------------

// InlineShapeElements returns all wp:inline elements found within the
// document body. This provides the raw element access; the domain
// InlineShapes proxy is created in MR-11.
//
// Mirrors the element query underlying Python DocumentPart.inline_shapes.
func (dp *DocumentPart) InlineShapeElements() ([]*etree.Element, error) {
	body, err := dp.Body()
	if err != nil {
		return nil, err
	}
	var inlines []*etree.Element
	findInlines(body.RawElement(), &inlines)
	return inlines, nil
}

// findInlines recursively collects wp:inline elements.
func findInlines(el *etree.Element, result *[]*etree.Element) {
	if el.Tag == "inline" && (el.Space == "wp" || el.Space == oxml.NsWp) {
		*result = append(*result, el)
	}
	for _, child := range el.ChildElements() {
		findInlines(child, result)
	}
}
