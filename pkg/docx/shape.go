package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// InlineShapes is a sequence of InlineShape instances found in a document body.
//
// Mirrors Python InlineShapes(Parented).
type InlineShapes struct {
	body *etree.Element // CT_Body element
	part *parts.StoryPart
}

// newInlineShapes creates a new InlineShapes proxy.
func newInlineShapes(body *etree.Element, part *parts.StoryPart) *InlineShapes {
	return &InlineShapes{body: body, part: part}
}

// Len returns the number of inline shapes in the document.
func (iss *InlineShapes) Len() int {
	return len(iss.inlineList())
}

// Get returns the inline shape at the given index.
func (iss *InlineShapes) Get(idx int) (*InlineShape, error) {
	list := iss.inlineList()
	if idx < 0 || idx >= len(list) {
		return nil, errIndexOutOfRange("InlineShapes", idx, len(list))
	}
	return &InlineShape{inline: list[idx], part: iss.part}, nil
}

// Iter returns all inline shapes in the document.
func (iss *InlineShapes) Iter() []*InlineShape {
	list := iss.inlineList()
	result := make([]*InlineShape, len(list))
	for i, il := range list {
		result[i] = &InlineShape{inline: il, part: iss.part}
	}
	return result
}

// inlineList walks the body tree to find all wp:inline elements
// (equivalent to Python's //w:p/w:r/w:drawing/wp:inline xpath).
func (iss *InlineShapes) inlineList() []*oxml.CT_Inline {
	var result []*oxml.CT_Inline
	for _, p := range iss.body.ChildElements() {
		if !(p.Space == "w" && p.Tag == "p") {
			continue
		}
		for _, r := range p.ChildElements() {
			if !(r.Space == "w" && r.Tag == "r") {
				continue
			}
			for _, drawing := range r.ChildElements() {
				if !(drawing.Space == "w" && drawing.Tag == "drawing") {
					continue
				}
				for _, inline := range drawing.ChildElements() {
					if inline.Space == "wp" && inline.Tag == "inline" {
						result = append(result, &oxml.CT_Inline{Element: oxml.WrapElement(inline)})
					}
				}
			}
		}
	}
	return result
}

// InlineShape is a proxy for a <wp:inline> element representing an inline graphical object.
//
// Mirrors Python InlineShape.
type InlineShape struct {
	inline *oxml.CT_Inline
	part   *parts.StoryPart
}

// newInlineShape creates a new InlineShape proxy.
func newInlineShape(elm *oxml.CT_Inline, part *parts.StoryPart) *InlineShape {
	return &InlineShape{inline: elm, part: part}
}

// Height returns the display height of this inline shape as a Length (EMU).
//
// Mirrors Python InlineShape.height (getter).
func (is *InlineShape) Height() (Length, error) {
	cy, err := is.inline.ExtentCy()
	if err != nil {
		return 0, err
	}
	return Length(cy), nil
}

// SetHeight sets the display height of this inline shape.
//
// Mirrors Python InlineShape.height (setter), which unconditionally sets
// both extent.cy and graphic.graphicData.pic.spPr.cy. In the XML schema
// graphic and graphicData are OneAndOnlyOne (required), so errors there
// indicate structural corruption. pic is ZeroOrOne — absent for charts
// and diagrams, where only the extent dimension matters.
func (is *InlineShape) SetHeight(v Length) error {
	cy := int64(v)
	if err := is.inline.SetExtentCy(cy); err != nil {
		return err
	}
	// Also update the spPr transform.
	graphic, err := is.inline.Graphic()
	if err != nil {
		return err
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return err
	}
	pic := gd.Pic()
	if pic == nil {
		return nil // non-picture shape (chart, diagram) — extent is sufficient
	}
	spPr, err := pic.SpPr()
	if err != nil {
		return err
	}
	return spPr.SetCy(cy)
}

// Width returns the display width of this inline shape as a Length (EMU).
//
// Mirrors Python InlineShape.width (getter).
func (is *InlineShape) Width() (Length, error) {
	cx, err := is.inline.ExtentCx()
	if err != nil {
		return 0, err
	}
	return Length(cx), nil
}

// SetWidth sets the display width of this inline shape.
//
// Mirrors Python InlineShape.width (setter). See SetHeight for rationale.
func (is *InlineShape) SetWidth(v Length) error {
	cx := int64(v)
	if err := is.inline.SetExtentCx(cx); err != nil {
		return err
	}
	// Also update the spPr transform.
	graphic, err := is.inline.Graphic()
	if err != nil {
		return err
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return err
	}
	pic := gd.Pic()
	if pic == nil {
		return nil // non-picture shape (chart, diagram) — extent is sufficient
	}
	spPr, err := pic.SpPr()
	if err != nil {
		return err
	}
	return spPr.SetCx(cx)
}

// Type returns the type of this inline shape (PICTURE, LINKED_PICTURE, CHART,
// SMART_ART, or NOT_IMPLEMENTED).
//
// Mirrors Python InlineShape.type.
func (is *InlineShape) Type() (enum.WdInlineShapeType, error) {
	graphic, err := is.inline.Graphic()
	if err != nil {
		return 0, fmt.Errorf("docx: accessing inline graphic: %w", err)
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return 0, fmt.Errorf("docx: accessing graphic data: %w", err)
	}
	uri, err := gd.Uri()
	if err != nil {
		return 0, fmt.Errorf("docx: reading graphic URI: %w", err)
	}

	switch uri {
	case oxml.NsPicture:
		pic := gd.Pic()
		if pic == nil {
			return enum.WdInlineShapeTypePicture, nil
		}
		bf, err := pic.BlipFill()
		if err != nil {
			return 0, fmt.Errorf("docx: accessing blip fill: %w", err)
		}
		blip := bf.Blip()
		if blip == nil {
			return enum.WdInlineShapeTypePicture, nil
		}
		if blip.Link() != "" {
			return enum.WdInlineShapeTypeLinkedPicture, nil
		}
		return enum.WdInlineShapeTypePicture, nil
	case oxml.NsChart:
		return enum.WdInlineShapeTypeChart, nil
	case oxml.NsDiagram:
		return enum.WdInlineShapeTypeSmartArt, nil
	default:
		return enum.WdInlineShapeTypeNotImplemented, nil
	}
}
