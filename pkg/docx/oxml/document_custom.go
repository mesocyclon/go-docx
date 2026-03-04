package oxml

import "github.com/beevik/etree"

// ===========================================================================
// CT_Document — custom methods
// ===========================================================================

// SectPrList returns all w:sectPr elements directly accessible from the document.
// This includes w:body/w:p/w:pPr/w:sectPr (paragraph section breaks) and
// w:body/w:sectPr (body-level final section). Results are in document order,
// with the body-level sectPr always last.
func (doc *CT_Document) SectPrList() []*CT_SectPr {
	body := doc.Body()
	if body == nil {
		return nil
	}

	var result []*CT_SectPr
	for _, child := range body.e.ChildElements() {
		// Paragraph-based sectPr: w:p/w:pPr/w:sectPr
		if child.Space == "w" && child.Tag == "p" {
			for _, pChild := range child.ChildElements() {
				if pChild.Space == "w" && pChild.Tag == "pPr" {
					for _, ppChild := range pChild.ChildElements() {
						if ppChild.Space == "w" && ppChild.Tag == "sectPr" {
							result = append(result, &CT_SectPr{Element{e: ppChild}})
						}
					}
				}
			}
		}
		// Body-level sectPr
		if child.Space == "w" && child.Tag == "sectPr" {
			result = append(result, &CT_SectPr{Element{e: child}})
		}
	}
	return result
}

// ===========================================================================
// CT_Body — custom methods
// ===========================================================================

// InnerContentElements returns all <w:p> and <w:tbl> direct children in document order.
// Elements inside wrapper elements (w:ins, w:sdt, etc.) are not included.
func (b *CT_Body) InnerContentElements() []BlockItem {
	var result []BlockItem
	for _, child := range b.e.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			result = append(result, &CT_P{Element{e: child}})
		} else if child.Space == "w" && child.Tag == "tbl" {
			result = append(result, &CT_Tbl{Element{e: child}})
		}
	}
	return result
}

// ClearContent removes all content child elements from this <w:body>,
// leaving the <w:sectPr> element if present.
func (b *CT_Body) ClearContent() {
	var toRemove []*etree.Element
	for _, child := range b.e.ChildElements() {
		if !(child.Space == "w" && child.Tag == "sectPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		b.e.RemoveChild(child)
	}
}

// AddSectionBreak adds a new section at the end of the document and returns
// the sentinel w:sectPr (which now controls the new last section).
//
// The previously-last w:sectPr is cloned into a new paragraph at the end,
// becoming the second-to-last section. Header and footer references are
// removed from the sentinel sectPr so they inherit from the prior section.
func (b *CT_Body) AddSectionBreak() *CT_SectPr {
	// Get the sentinel sectPr at file-end
	sentinelSectPr := b.GetOrAddSectPr()

	// Clone it and add to a new paragraph (becomes second-to-last section)
	clonedSectPr := &CT_SectPr{Element{e: sentinelSectPr.e.Copy()}}
	newP := b.AddP()
	newP.SetSectPr(clonedSectPr)

	// Remove header/footer references from the sentinel sectPr
	var hdrFtrRefs []*etree.Element
	for _, child := range sentinelSectPr.e.ChildElements() {
		if child.Space == "w" && (child.Tag == "headerReference" || child.Tag == "footerReference") {
			hdrFtrRefs = append(hdrFtrRefs, child)
		}
	}
	for _, ref := range hdrFtrRefs {
		sentinelSectPr.e.RemoveChild(ref)
	}

	return sentinelSectPr
}

// SetSectPr replaces or adds the w:sectPr child element.
func (b *CT_Body) SetSectPr(sectPr *CT_SectPr) {
	b.RemoveSectPr()
	b.e.AddChild(sectPr.e)
}
