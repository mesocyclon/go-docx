package oxml

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ===========================================================================
// CT_SectPr — custom methods
// ===========================================================================

// Clone returns a deep copy of this sectPr element with all rsid attributes removed.
func (sp *CT_SectPr) Clone() *CT_SectPr {
	copied := sp.e.Copy()
	// Remove rsid* attributes
	var toRemove []string
	for _, attr := range copied.Attr {
		key := attr.Key
		if len(key) >= 4 && key[:4] == "rsid" {
			toRemove = append(toRemove, attr.FullKey())
		}
		if attr.Space != "" {
			qKey := attr.Space + ":" + attr.Key
			if len(attr.Key) >= 4 && attr.Key[:4] == "rsid" {
				toRemove = append(toRemove, qKey)
			}
		}
	}
	for _, k := range toRemove {
		copied.RemoveAttr(k)
	}
	return &CT_SectPr{Element{e: copied}}
}

// --- Page size ---

// PageWidth returns the page width in twips from pgSz/@w:w, or nil.
func (sp *CT_SectPr) PageWidth() (*int, error) {
	pgSz := sp.PgSz()
	if pgSz == nil {
		return nil, nil
	}
	return pgSz.W()
}

// SetPageWidth sets the page width in twips.
func (sp *CT_SectPr) SetPageWidth(twips *int) error {
	if twips == nil {
		pgSz := sp.PgSz()
		if pgSz != nil {
			if err := pgSz.SetW(nil); err != nil {
				return err
			}
		}
		return nil
	}
	if err := sp.GetOrAddPgSz().SetW(twips); err != nil {
		return err
	}
	return nil
}

// PageHeight returns the page height in twips from pgSz/@w:h, or nil.
func (sp *CT_SectPr) PageHeight() (*int, error) {
	pgSz := sp.PgSz()
	if pgSz == nil {
		return nil, nil
	}
	return pgSz.H()
}

// SetPageHeight sets the page height in twips.
func (sp *CT_SectPr) SetPageHeight(twips *int) error {
	if twips == nil {
		pgSz := sp.PgSz()
		if pgSz != nil {
			if err := pgSz.SetH(nil); err != nil {
				return err
			}
		}
		return nil
	}
	if err := sp.GetOrAddPgSz().SetH(twips); err != nil {
		return err
	}
	return nil
}

// --- Orientation ---

// Orientation returns the page orientation. Defaults to PORTRAIT when not present.
func (sp *CT_SectPr) Orientation() (enum.WdOrientation, error) {
	pgSz := sp.PgSz()
	if pgSz == nil {
		return enum.WdOrientationPortrait, nil
	}
	v, err := pgSz.Orient()
	if err != nil {
		return enum.WdOrientation(0), err
	}
	if v == enum.WdOrientation(0) {
		return enum.WdOrientationPortrait, nil
	}
	return v, nil
}

// SetOrientation sets the page orientation.
func (sp *CT_SectPr) SetOrientation(v enum.WdOrientation) error {
	pgSz := sp.GetOrAddPgSz()
	if v == enum.WdOrientationPortrait {
		return pgSz.SetOrient(enum.WdOrientation(0)) // removes attr, defaulting to portrait
	}
	return pgSz.SetOrient(v)
}

// --- Start type ---

// StartType returns the section start type. Defaults to NEW_PAGE when not present.
func (sp *CT_SectPr) StartType() (enum.WdSectionStart, error) {
	t := sp.Type()
	if t == nil {
		return enum.WdSectionStartNewPage, nil
	}
	// Check if val attribute is actually present
	_, ok := t.GetAttr("w:val")
	if !ok {
		return enum.WdSectionStartNewPage, nil
	}
	return t.Val()
}

// SetStartType sets the section start type. Passing WdSectionStartNewPage removes
// the type element (since NEW_PAGE is the default).
func (sp *CT_SectPr) SetStartType(v enum.WdSectionStart) error {
	if v == enum.WdSectionStartNewPage {
		sp.RemoveType()
		return nil
	}
	xmlVal, err := v.ToXml()
	if err != nil {
		return fmt.Errorf("oxml: invalid section start type: %w", err)
	}
	t := sp.GetOrAddType()
	// Use SetAttr directly because the generated SetVal treats zero (Continuous)
	// as "remove attribute", but we actually need to write w:val="continuous".
	t.SetAttr("w:val", xmlVal)
	return nil
}

// --- Title page ---

// TitlePgVal returns true if the first page has different header/footer.
func (sp *CT_SectPr) TitlePgVal() bool {
	tp := sp.TitlePg()
	if tp == nil {
		return false
	}
	return tp.Val()
}

// SetTitlePgVal sets the titlePg flag. Passing false removes the element.
func (sp *CT_SectPr) SetTitlePgVal(v bool) error {
	if !v {
		sp.RemoveTitlePg()
		return nil
	}
	if err := sp.GetOrAddTitlePg().SetVal(true); err != nil {
		return err
	}
	return nil
}

// --- Margins ---

// TopMargin returns the top margin in twips, or nil if not present.
func (sp *CT_SectPr) TopMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Top()
}

// SetTopMargin sets the top margin in twips. Passing nil removes the attribute.
func (sp *CT_SectPr) SetTopMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetTop(twips)
}

// BottomMargin returns the bottom margin in twips, or nil.
func (sp *CT_SectPr) BottomMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Bottom()
}

// SetBottomMargin sets the bottom margin in twips.
func (sp *CT_SectPr) SetBottomMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetBottom(twips)
}

// LeftMargin returns the left margin in twips, or nil.
func (sp *CT_SectPr) LeftMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Left()
}

// SetLeftMargin sets the left margin in twips.
func (sp *CT_SectPr) SetLeftMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetLeft(twips)
}

// RightMargin returns the right margin in twips, or nil.
func (sp *CT_SectPr) RightMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Right()
}

// SetRightMargin sets the right margin in twips.
func (sp *CT_SectPr) SetRightMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetRight(twips)
}

// HeaderMargin returns the header distance from top edge in twips, or nil.
func (sp *CT_SectPr) HeaderMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Header()
}

// SetHeaderMargin sets the header margin in twips.
func (sp *CT_SectPr) SetHeaderMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetHeader(twips)
}

// FooterMargin returns the footer distance from bottom edge in twips, or nil.
func (sp *CT_SectPr) FooterMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Footer()
}

// SetFooterMargin sets the footer margin in twips.
func (sp *CT_SectPr) SetFooterMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetFooter(twips)
}

// GutterMargin returns the gutter in twips, or nil.
func (sp *CT_SectPr) GutterMargin() (*int, error) {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil, nil
	}
	return pgMar.Gutter()
}

// SetGutterMargin sets the gutter in twips.
func (sp *CT_SectPr) SetGutterMargin(twips *int) error {
	pgMar := sp.GetOrAddPgMar()
	return pgMar.SetGutter(twips)
}

// --- Header/Footer references ---

// AddHeaderRef adds a headerReference with the given type and relationship ID.
func (sp *CT_SectPr) AddHeaderRef(hfType enum.WdHeaderFooterIndex, rId string) (*CT_HdrFtrRef, error) {
	ref := sp.AddHeaderReference()
	if err := ref.SetType(hfType); err != nil {
		return nil, fmt.Errorf("AddHeaderRef: %w", err)
	}
	if err := ref.SetRId(rId); err != nil {
		return nil, err
	}
	return ref, nil
}

// AddFooterRef adds a footerReference with the given type and relationship ID.
func (sp *CT_SectPr) AddFooterRef(hfType enum.WdHeaderFooterIndex, rId string) (*CT_HdrFtrRef, error) {
	ref := sp.AddFooterReference()
	if err := ref.SetType(hfType); err != nil {
		return nil, fmt.Errorf("AddFooterRef: %w", err)
	}
	if err := ref.SetRId(rId); err != nil {
		return nil, err
	}
	return ref, nil
}

// GetHeaderRef returns the headerReference of the given type, or nil.
// Returns an error if hfType has no XML representation.
func (sp *CT_SectPr) GetHeaderRef(hfType enum.WdHeaderFooterIndex) (*CT_HdrFtrRef, error) {
	xmlVal, err := hfType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("oxml: invalid header/footer type: %w", err)
	}
	for _, ref := range sp.HeaderReferenceList() {
		v, ok := ref.GetAttr("w:type")
		if ok && v == xmlVal {
			return ref, nil
		}
	}
	return nil, nil
}

// GetFooterRef returns the footerReference of the given type, or nil.
// Returns an error if hfType has no XML representation.
func (sp *CT_SectPr) GetFooterRef(hfType enum.WdHeaderFooterIndex) (*CT_HdrFtrRef, error) {
	xmlVal, err := hfType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("oxml: invalid header/footer type: %w", err)
	}
	for _, ref := range sp.FooterReferenceList() {
		v, ok := ref.GetAttr("w:type")
		if ok && v == xmlVal {
			return ref, nil
		}
	}
	return nil, nil
}

// RemoveHeaderRef removes the headerReference of the given type and returns its rId.
// Returns "" if not found.
func (sp *CT_SectPr) RemoveHeaderRef(hfType enum.WdHeaderFooterIndex) (string, error) {
	ref, err := sp.GetHeaderRef(hfType)
	if err != nil {
		return "", fmt.Errorf("oxml: getting header ref: %w", err)
	}
	if ref == nil {
		return "", nil
	}
	rId, err := ref.RId()
	if err != nil {
		return "", fmt.Errorf("oxml: reading header rId: %w", err)
	}
	sp.e.RemoveChild(ref.e)
	return rId, nil
}

// RemoveFooterRef removes the footerReference of the given type and returns its rId.
// Returns "" if not found.
func (sp *CT_SectPr) RemoveFooterRef(hfType enum.WdHeaderFooterIndex) (string, error) {
	ref, err := sp.GetFooterRef(hfType)
	if err != nil {
		return "", fmt.Errorf("oxml: getting footer ref: %w", err)
	}
	if ref == nil {
		return "", nil
	}
	rId, err := ref.RId()
	if err != nil {
		return "", fmt.Errorf("oxml: reading footer rId: %w", err)
	}
	sp.e.RemoveChild(ref.e)
	return rId, nil
}

// BodyElement returns the w:body element that contains this sectPr.
//
// A sectPr can be either:
//   - paragraph-based: w:body/w:p/w:pPr/w:sectPr → parent is pPr
//   - body-based:      w:body/w:sectPr            → parent is body
//
// Mirrors Python _SectBlockElementIterator navigating parent::w:pPr/parent::w:p
// vs self::w:sectPr[parent::w:body].
func (sp *CT_SectPr) BodyElement() *etree.Element {
	parent := sp.e.Parent()
	if parent == nil {
		return nil
	}
	if parent.Space == "w" && parent.Tag == "body" {
		return parent
	}
	if parent.Space == "w" && parent.Tag == "pPr" {
		if p := parent.Parent(); p != nil {
			return p.Parent()
		}
	}
	return nil
}

// PrecedingSectPr returns the sectPr immediately preceding this one, or nil.
// Searches via preceding-sibling within the w:body, accounting for both
// paragraph-based sectPr (w:p/w:pPr/w:sectPr) and body-based sectPr (w:body/w:sectPr).
func (sp *CT_SectPr) PrecedingSectPr() *CT_SectPr {
	body := sp.BodyElement()
	if body == nil {
		return nil
	}

	// Gather all sectPr elements in document order
	var allSectPrs []*CT_SectPr
	for _, child := range body.ChildElements() {
		// Check p/pPr/sectPr
		if child.Space == "w" && child.Tag == "p" {
			for _, pChild := range child.ChildElements() {
				if pChild.Space == "w" && pChild.Tag == "pPr" {
					for _, ppChild := range pChild.ChildElements() {
						if ppChild.Space == "w" && ppChild.Tag == "sectPr" {
							allSectPrs = append(allSectPrs, &CT_SectPr{Element{e: ppChild}})
						}
					}
				}
			}
		}
		// Check body/sectPr
		if child.Space == "w" && child.Tag == "sectPr" {
			allSectPrs = append(allSectPrs, &CT_SectPr{Element{e: child}})
		}
	}

	for i, s := range allSectPrs {
		if s.e == sp.e && i > 0 {
			return allSectPrs[i-1]
		}
	}
	return nil
}

// ===========================================================================
// CT_HdrFtr — custom methods
// ===========================================================================

// InnerContentElements returns all w:p and w:tbl direct children in document order.
func (hf *CT_HdrFtr) InnerContentElements() []BlockItem {
	var result []BlockItem
	for _, child := range hf.e.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			result = append(result, &CT_P{Element{e: child}})
		} else if child.Space == "w" && child.Tag == "tbl" {
			result = append(result, &CT_Tbl{Element{e: child}})
		}
	}
	return result
}
