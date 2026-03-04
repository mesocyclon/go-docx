package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/image"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Drawing is a container for a DrawingML object within a run.
//
// Mirrors Python Drawing(Parented) from docx/drawing/__init__.py.
type Drawing struct {
	drawing *oxml.CT_Drawing
	part    *parts.StoryPart
}

// newDrawing creates a new Drawing proxy.
func newDrawing(drawing *oxml.CT_Drawing, part *parts.StoryPart) *Drawing {
	return &Drawing{drawing: drawing, part: part}
}

// HasPicture returns true when the drawing contains an embedded picture.
// A drawing can also contain a chart, SmartArt, or drawing canvas.
//
// Checks for both inline and floating pictures:
//
//	wp:inline/a:graphic/a:graphicData/pic:pic
//	wp:anchor/a:graphic/a:graphicData/pic:pic
//
// Mirrors Python Drawing.has_picture.
func (d *Drawing) HasPicture() bool {
	for _, child := range d.drawing.RawElement().ChildElements() {
		if child.Tag == "inline" || child.Tag == "anchor" {
			if findPicInGraphicData(child) {
				return true
			}
		}
	}
	return false
}

// ImagePart returns the ImagePart for the embedded picture in this drawing.
// Returns an error if the drawing does not contain a picture.
//
// Mirrors Python Drawing.image which returns image_part.image.
func (d *Drawing) ImagePart() (*parts.ImagePart, error) {
	rId := d.pictureRId()
	if rId == "" {
		return nil, fmt.Errorf("docx: drawing does not contain a picture")
	}
	if d.part == nil {
		return nil, fmt.Errorf("docx: drawing has no story part (required for image resolution)")
	}
	rels := d.part.Rels()
	if rels == nil {
		return nil, fmt.Errorf("docx: drawing part has no relationships")
	}
	relParts := rels.RelatedParts()
	p, ok := relParts[rId]
	if !ok {
		return nil, fmt.Errorf("docx: no related part for rId %q", rId)
	}
	ip, ok := p.(*parts.ImagePart)
	if !ok {
		return nil, fmt.Errorf("docx: related part for rId %q is not an ImagePart", rId)
	}
	return ip, nil
}

// Image returns the image.Image metadata for the embedded picture. Returns
// an error if the drawing does not contain a picture or if the image blob
// cannot be parsed.
//
// This is a convenience method; it is equivalent to:
//
//	ip, _ := drawing.ImagePart()
//	img, _ := image.FromBlob(ip.Blob())
//
// Mirrors Python Drawing.image (which returns image_part.image).
func (d *Drawing) Image() (*image.Image, error) {
	ip, err := d.ImagePart()
	if err != nil {
		return nil, err
	}
	blob, err := ip.Blob()
	if err != nil {
		return nil, fmt.Errorf("docx: reading image blob: %w", err)
	}
	img, err := image.FromBlob(blob, ip.Filename())
	if err != nil {
		return nil, fmt.Errorf("docx: parsing image: %w", err)
	}
	return img, nil
}

// pictureRId finds the r:embed attribute on a:blip inside pic:blipFill.
// Path: drawing → inline/anchor → graphic → graphicData → pic → blipFill → blip → @r:embed.
//
// Mirrors Python: self._drawing.xpath(".//pic:blipFill/a:blip/@r:embed")
func (d *Drawing) pictureRId() string {
	for _, child := range d.drawing.RawElement().ChildElements() {
		if child.Tag == "inline" || child.Tag == "anchor" {
			if rId := findBlipRId(child); rId != "" {
				return rId
			}
		}
	}
	return ""
}

// findBlipRId walks the known path graphic → graphicData → pic → blipFill → blip
// and returns the r:embed attribute value from the blip element.
func findBlipRId(el *etree.Element) string {
	for _, blip := range walkPath(el, "graphic", "graphicData", "pic", "blipFill", "blip") {
		for _, attr := range blip.Attr {
			if attr.Key == "embed" && (attr.Space == "r" || attr.Space == "") {
				return attr.Value
			}
		}
	}
	return ""
}

// findPicInGraphicData walks graphic → graphicData and checks for a pic:pic child.
func findPicInGraphicData(el *etree.Element) bool {
	return len(walkPath(el, "graphic", "graphicData", "pic")) > 0
}

// walkPath performs an iterative level-by-level descent through child elements
// matching successive tags. Returns all elements that match the full path.
//
// For example, walkPath(el, "graphic", "graphicData", "pic") finds all elements
// reachable by el → graphic → graphicData → pic without any recursion.
func walkPath(root *etree.Element, tags ...string) []*etree.Element {
	current := []*etree.Element{root}
	for _, tag := range tags {
		var next []*etree.Element
		for _, el := range current {
			for _, child := range el.ChildElements() {
				if child.Tag == tag {
					next = append(next, child)
				}
			}
		}
		current = next
	}
	return current
}

// CT_Drawing returns the underlying oxml element.
func (d *Drawing) CT_Drawing() *oxml.CT_Drawing { return d.drawing }
