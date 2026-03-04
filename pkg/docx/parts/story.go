// Package parts implements WML-specific part types (document, styles, headers, etc.)
// that extend the generic OPC part infrastructure.
package parts

import (
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/internal/xmlutil"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/image"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// StoryPart is the base for document-body, header, and footer parts.
// A story part is one that can contain textual content. These all share
// content behaviours like paragraphs, tables, and images.
//
// Mirrors Python StoryPart(XmlPart).
type StoryPart struct {
	*opc.XmlPart
	docPart  *DocumentPart // cached, mirrors Python lazyproperty _document_part
	lastID   int           // cached max ID from last scan / allocation
	idScaned bool          // true after first NextID call
}

// NewStoryPart creates a StoryPart wrapping the given XmlPart.
func NewStoryPart(xp *opc.XmlPart) *StoryPart {
	return &StoryPart{XmlPart: xp}
}

// GetOrAddImage returns (rId, imagePart) for the image identified by imgPart.
// The caller is expected to resolve the image through WmlPackage first (which
// handles SHA-256 deduplication), then this method wires the relationship.
//
// Mirrors Python StoryPart.get_or_add_image:
//
//	image_part = package.get_or_add_image_part(image_descriptor)
//	rId = self.relate_to(image_part, RT.IMAGE)
//	return rId, image_part.image
func (sp *StoryPart) GetOrAddImage(imgPart *ImagePart) (string, *ImagePart) {
	rId := sp.Rels().GetOrAdd(opc.RTImage, imgPart).RID
	return rId, imgPart
}

// NewPicInline creates a new CT_Inline element containing the image specified
// by imgPart, scaled to the given width/height.
//
// Mirrors Python StoryPart.new_pic_inline EXACTLY:
//
//	rId, image = self.get_or_add_image(image_descriptor)
//	cx, cy = image.scaled_dimensions(width, height)
//	shape_id, filename = self.next_id, image.filename
//	return CT_Inline.new_pic_inline(shape_id, rId, filename, cx, cy)
func (sp *StoryPart) NewPicInline(imgPart *ImagePart, width, height *int64) (*oxml.CT_Inline, error) {
	rId, ip := sp.GetOrAddImage(imgPart)
	cx, cy, err := ip.ScaledDimensions(width, height)
	if err != nil {
		return nil, fmt.Errorf("parts: computing scaled dimensions: %w", err)
	}
	shapeID := sp.NextID()
	filename := ip.Filename()
	return oxml.NewPicInline(shapeID, rId, filename, cx, cy)
}

// GetStyle returns the style in this document matching styleID.
// Returns the default style for styleType if styleID is nil or does not
// match a defined style of styleType.
//
// Mirrors Python StoryPart.get_style — delegates to _document_part.get_style.
func (sp *StoryPart) GetStyle(styleID *string, styleType enum.WdStyleType) (*oxml.CT_Style, error) {
	dp, err := sp.documentPart()
	if err != nil {
		return nil, err
	}
	return dp.GetStyle(styleID, styleType)
}

// GetStyleID returns the style_id string for styleOrName of styleType.
// Returns nil if the style resolves to the default style for styleType or if
// styleOrName is itself nil.
//
// styleOrName accepts the same types as [DocumentPart.GetStyleID]:
// string (style UI name), styledObject, or nil.
//
// Mirrors Python StoryPart.get_style_id — delegates to _document_part.get_style_id.
func (sp *StoryPart) GetStyleID(styleOrName any, styleType enum.WdStyleType) (*string, error) {
	dp, err := sp.documentPart()
	if err != nil {
		return nil, err
	}
	return dp.GetStyleID(styleOrName, styleType)
}

// NextID returns the next available positive integer id value in this story
// XML document. The first call scans the tree to find the current maximum;
// subsequent calls simply increment the counter. Gaps in the existing id
// sequence are not filled.
//
// Mirrors Python StoryPart.next_id.
func (sp *StoryPart) NextID() int {
	if !sp.idScaned {
		el := sp.Element()
		if el != nil {
			sp.lastID = collectMaxID(el)
		}
		sp.idScaned = true
	}
	sp.lastID++
	return sp.lastID
}

// documentPart returns the main DocumentPart for the package this story part
// belongs to. The result is cached after the first call.
//
// Mirrors Python StoryPart._document_part (lazyproperty).
func (sp *StoryPart) documentPart() (*DocumentPart, error) {
	if sp.docPart != nil {
		return sp.docPart, nil
	}
	pkg := sp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: story part has no package")
	}
	mainPart, err := pkg.MainDocumentPart()
	if err != nil {
		return nil, fmt.Errorf("parts: resolving document part: %w", err)
	}
	dp, ok := mainPart.(*DocumentPart)
	if !ok {
		return nil, fmt.Errorf("parts: main document part is %T, want *DocumentPart", mainPart)
	}
	sp.docPart = dp
	return dp, nil
}

// SetDocumentPart sets the cached document part reference. Used by
// DocumentPart to set itself as its own document part.
func (sp *StoryPart) SetDocumentPart(dp *DocumentPart) {
	sp.docPart = dp
}

// wmlPackage returns the WML package for this part by resolving it from
// the underlying OpcPackage.AppPackage(). This mirrors Python where
// Part._package IS the Package(OpcPackage) subclass — any Part can call
// self._package.get_or_add_image_part() because _package is the WML-level
// package. In Go the equivalent is Package().AppPackage().(*WmlPackage).
func (sp *StoryPart) wmlPackage() *WmlPackage {
	pkg := sp.Package()
	if pkg == nil {
		return nil
	}
	wp, _ := pkg.AppPackage().(*WmlPackage)
	return wp
}

// GetOrAddImageFromReader creates or deduplicates an image from the given
// reader and returns (rId, ImagePart). This is the stream-based version that
// mirrors the Python StoryPart.get_or_add_image(image_descriptor) flow:
//
//	image_part = package.get_or_add_image_part(image_descriptor)
//	rId = self.relate_to(image_part, RT.IMAGE)
//	return rId, image_part
func (sp *StoryPart) GetOrAddImageFromReader(r io.ReadSeeker) (string, *ImagePart, error) {
	wp := sp.wmlPackage()
	if wp == nil {
		return "", nil, fmt.Errorf("parts: WmlPackage not set on OpcPackage (required for image insertion)")
	}
	// Read blob
	if _, err := r.Seek(0, io.SeekStart); err != nil {
		return "", nil, fmt.Errorf("parts: seeking image stream: %w", err)
	}
	blob, err := io.ReadAll(r)
	if err != nil {
		return "", nil, fmt.Errorf("parts: reading image stream: %w", err)
	}
	// Parse image metadata
	img, err := image.FromBlob(blob, "")
	if err != nil {
		return "", nil, fmt.Errorf("parts: parsing image: %w", err)
	}
	// Create ImagePart
	ip := NewImagePartFromImage(img, blob)
	// Dedup via WmlPackage
	ip, err = wp.GetOrAddImagePart(ip)
	if err != nil {
		return "", nil, fmt.Errorf("parts: dedup image part: %w", err)
	}
	// Wire relationship
	rId := sp.Rels().GetOrAdd(opc.RTImage, ip).RID
	return rId, ip, nil
}

// NewPicInlineFromReader creates a new CT_Inline element from an image stream.
// This mirrors the Python StoryPart.new_pic_inline(image_descriptor, width, height)
// flow, where the caller provides a path or stream, not a pre-built ImagePart.
func (sp *StoryPart) NewPicInlineFromReader(r io.ReadSeeker, width, height *int64) (*oxml.CT_Inline, error) {
	rId, ip, err := sp.GetOrAddImageFromReader(r)
	if err != nil {
		return nil, err
	}
	cx, cy, err := ip.ScaledDimensions(width, height)
	if err != nil {
		return nil, fmt.Errorf("parts: computing scaled dimensions: %w", err)
	}
	shapeID := sp.NextID()
	filename := ip.Filename()
	return oxml.NewPicInline(shapeID, rId, filename, cx, cy)
}

// DropRel removes the relationship identified by rId if its reference count
// in this part's XML is less than 2. This prevents removing relationships
// that are still referenced elsewhere in the XML.
//
// Mirrors Python Part.drop_rel + XmlPart._rel_ref_count.
func (sp *StoryPart) DropRel(rId string) {
	if sp.relRefCount(rId) < 2 {
		sp.Rels().Delete(rId)
	}
}

// relRefCount returns the count of references to rId in this part's XML.
// Mirrors Python XmlPart._rel_ref_count which counts //@r:id occurrences.
func (sp *StoryPart) relRefCount(rId string) int {
	el := sp.Element()
	if el == nil {
		return 0
	}
	return countRIdRefs(el, rId)
}

// --------------------------------------------------------------------------
// internal helpers
// --------------------------------------------------------------------------

// collectMaxID iterates the element tree collecting all @id attributes that
// are purely numeric digits, returning the maximum value found (0 if none).
func collectMaxID(root *etree.Element) int {
	maxID := 0
	stack := []*etree.Element{root}
	for len(stack) > 0 {
		el := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		for _, attr := range el.Attr {
			if attr.Key == "id" && attr.Space == "" && xmlutil.IsDigits(attr.Value) {
				if v, err := strconv.Atoi(attr.Value); err == nil && v > maxID {
					maxID = v
				}
			}
		}
		stack = append(stack, el.ChildElements()...)
	}
	return maxID
}

// countRIdRefs iteratively counts attributes named r:id (or {relationship-ns}id)
// with the given value across the entire element tree.
func countRIdRefs(root *etree.Element, rId string) int {
	count := 0
	stack := []*etree.Element{root}
	for len(stack) > 0 {
		el := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		for _, attr := range el.Attr {
			if attr.Key == "id" && isRelNS(attr.Space) && attr.Value == rId {
				count++
			}
		}
		stack = append(stack, el.ChildElements()...)
	}
	return count
}

// isRelNS returns true if the namespace prefix or URI matches the OFC
// relationships namespace used for r:id attributes.
func isRelNS(space string) bool {
	return space == "r" ||
		strings.Contains(space, "officeDocument/2006/relationships")
}
