package oxml

import "fmt"

// ===========================================================================
// CT_Inline — custom methods
// ===========================================================================

// NewPicInline creates a new <wp:inline> element containing a <pic:pic> element
// for an inline image. shape_id is an integer identifier, rId is the relationship
// id for the image part, filename is the original image name, cx and cy are the
// image dimensions in EMU.
func NewPicInline(shapeId int, rId, filename string, cx, cy int64) (*CT_Inline, error) {
	pic, err := newPicture(0, filename, rId, cx, cy)
	if err != nil {
		return nil, err
	}
	return newInline(cx, cy, shapeId, pic)
}

// newInline creates a <wp:inline> skeleton and fills it with the given values.
func newInline(cx, cy int64, shapeId int, pic *CT_Picture) (*CT_Inline, error) {
	xml := fmt.Sprintf(
		`<wp:inline ` +
			`xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
			`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
			`xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<wp:extent cx="914400" cy="914400"/>` +
			`<wp:docPr id="777" name="unnamed"/>` +
			`<wp:cNvGraphicFramePr>` +
			`<a:graphicFrameLocks noChangeAspect="1"/>` +
			`</wp:cNvGraphicFramePr>` +
			`<a:graphic>` +
			`<a:graphicData uri="URI not set"/>` +
			`</a:graphic>` +
			`</wp:inline>`,
	)
	el, err := ParseXml([]byte(xml))
	if err != nil {
		return nil, fmt.Errorf("oxml: failed to parse inline XML: %w", err)
	}
	inline := &CT_Inline{Element{e: el}}

	// Set extent dimensions
	extent, err := inline.Extent()
	if err != nil {
		return nil, fmt.Errorf("oxml: inline missing extent: %w", err)
	}
	if err := extent.SetCx(cx); err != nil {
		return nil, err
	}
	if err := extent.SetCy(cy); err != nil {
		return nil, err
	}

	// Set docPr
	docPr, err := inline.DocPr()
	if err != nil {
		return nil, fmt.Errorf("oxml: inline missing docPr: %w", err)
	}
	if err := docPr.SetId(shapeId); err != nil {
		return nil, err
	}
	if err := docPr.SetName(fmt.Sprintf("Picture %d", shapeId)); err != nil {
		return nil, err
	}

	// Set graphic data URI and insert the picture element
	graphic, err := inline.Graphic()
	if err != nil {
		return nil, fmt.Errorf("oxml: inline missing graphic: %w", err)
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return nil, fmt.Errorf("oxml: graphic missing graphicData: %w", err)
	}
	if err := gd.SetUri(NsPicture); err != nil {
		return nil, err
	}
	// Insert pic:pic into graphicData
	gd.e.AddChild(pic.e)

	return inline, nil
}

// newPicture creates a minimum viable <pic:pic> element.
func newPicture(picId int, filename, rId string, cx, cy int64) (*CT_Picture, error) {
	xml := fmt.Sprintf(
		`<pic:pic ` +
			`xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
			`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<pic:nvPicPr>` +
			`<pic:cNvPr id="777" name="unnamed"/>` +
			`<pic:cNvPicPr/>` +
			`</pic:nvPicPr>` +
			`<pic:blipFill>` +
			`<a:blip/>` +
			`<a:stretch>` +
			`<a:fillRect/>` +
			`</a:stretch>` +
			`</pic:blipFill>` +
			`<pic:spPr>` +
			`<a:xfrm>` +
			`<a:off x="0" y="0"/>` +
			`<a:ext cx="914400" cy="914400"/>` +
			`</a:xfrm>` +
			`<a:prstGeom prst="rect"/>` +
			`</pic:spPr>` +
			`</pic:pic>`,
	)
	el, err := ParseXml([]byte(xml))
	if err != nil {
		return nil, fmt.Errorf("oxml: failed to parse pic XML: %w", err)
	}
	pic := &CT_Picture{Element{e: el}}

	// Set picture properties
	nvPicPr, err := pic.NvPicPr()
	if err != nil {
		return nil, fmt.Errorf("oxml: pic missing nvPicPr: %w", err)
	}
	cNvPr, err := nvPicPr.CNvPr()
	if err != nil {
		return nil, fmt.Errorf("oxml: nvPicPr missing cNvPr: %w", err)
	}
	if err := cNvPr.SetId(picId); err != nil {
		return nil, err
	}
	if err := cNvPr.SetName(filename); err != nil {
		return nil, err
	}
	blipFill, err := pic.BlipFill()
	if err != nil {
		return nil, fmt.Errorf("oxml: pic missing blipFill: %w", err)
	}
	blip := blipFill.Blip()
	if blip == nil {
		return nil, fmt.Errorf("oxml: blipFill missing blip element")
	}
	if err := blip.SetEmbed(rId); err != nil {
		return nil, err
	}
	spPr, err := pic.SpPr()
	if err != nil {
		return nil, fmt.Errorf("oxml: pic missing spPr: %w", err)
	}
	if err := spPr.SetCx(cx); err != nil {
		return nil, err
	}
	if err := spPr.SetCy(cy); err != nil {
		return nil, err
	}

	return pic, nil
}

// ExtentCx returns the width of the inline image in EMU.
func (i *CT_Inline) ExtentCx() (int64, error) {
	extent, err := i.Extent()
	if err != nil {
		return 0, fmt.Errorf("ExtentCx: %w", err)
	}
	v, err := extent.Cx()
	if err != nil {
		return 0, fmt.Errorf("ExtentCx: %w", err)
	}
	return v, nil
}

// ExtentCy returns the height of the inline image in EMU.
func (i *CT_Inline) ExtentCy() (int64, error) {
	extent, err := i.Extent()
	if err != nil {
		return 0, fmt.Errorf("ExtentCy: %w", err)
	}
	v, err := extent.Cy()
	if err != nil {
		return 0, fmt.Errorf("ExtentCy: %w", err)
	}
	return v, nil
}

// SetExtentCx sets the width of the inline image in EMU.
func (i *CT_Inline) SetExtentCx(v int64) error {
	extent, err := i.Extent()
	if err != nil {
		return fmt.Errorf("SetExtentCx: %w", err)
	}
	if err := extent.SetCx(v); err != nil {
		return err
	}
	return nil
}

// SetExtentCy sets the height of the inline image in EMU.
func (i *CT_Inline) SetExtentCy(v int64) error {
	extent, err := i.Extent()
	if err != nil {
		return fmt.Errorf("SetExtentCy: %w", err)
	}
	if err := extent.SetCy(v); err != nil {
		return err
	}
	return nil
}

// ===========================================================================
// CT_ShapeProperties — custom methods
// ===========================================================================

// Cx returns the shape width in EMU via xfrm/ext/@cx, or nil if not present.
func (sp *CT_ShapeProperties) Cx() (*int64, error) {
	xfrm := sp.Xfrm()
	if xfrm == nil {
		return nil, nil
	}
	return xfrm.CxVal()
}

// SetCx sets the shape width in EMU via xfrm/ext/@cx.
func (sp *CT_ShapeProperties) SetCx(v int64) error {
	xfrm := sp.GetOrAddXfrm()
	if err := xfrm.SetCxVal(v); err != nil {
		return err
	}
	return nil
}

// Cy returns the shape height in EMU via xfrm/ext/@cy, or nil if not present.
func (sp *CT_ShapeProperties) Cy() (*int64, error) {
	xfrm := sp.Xfrm()
	if xfrm == nil {
		return nil, nil
	}
	return xfrm.CyVal()
}

// SetCy sets the shape height in EMU via xfrm/ext/@cy.
func (sp *CT_ShapeProperties) SetCy(v int64) error {
	xfrm := sp.GetOrAddXfrm()
	if err := xfrm.SetCyVal(v); err != nil {
		return err
	}
	return nil
}

// ===========================================================================
// CT_Transform2D — custom methods
// ===========================================================================

// CxVal returns the width in EMU from ext/@cx, or nil if ext is not present.
func (t *CT_Transform2D) CxVal() (*int64, error) {
	ext := t.Ext()
	if ext == nil {
		return nil, nil
	}
	v, err := ext.Cx()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetCxVal sets the width in EMU on ext/@cx, creating ext if needed.
func (t *CT_Transform2D) SetCxVal(v int64) error {
	ext := t.GetOrAddExt()
	if err := ext.SetCx(v); err != nil {
		return err
	}
	return nil
}

// CyVal returns the height in EMU from ext/@cy, or nil if ext is not present.
func (t *CT_Transform2D) CyVal() (*int64, error) {
	ext := t.Ext()
	if ext == nil {
		return nil, nil
	}
	v, err := ext.Cy()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetCyVal sets the height in EMU on ext/@cy, creating ext if needed.
func (t *CT_Transform2D) SetCyVal(v int64) error {
	ext := t.GetOrAddExt()
	if err := ext.SetCy(v); err != nil {
		return err
	}
	return nil
}
