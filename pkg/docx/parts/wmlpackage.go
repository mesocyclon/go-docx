package parts

import (
	"fmt"
	"strings"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// WmlPackage adds WML-specific behaviours to OpcPackage, primarily image
// part management with SHA-256-based deduplication.
//
// Mirrors Python Package(OpcPackage).
type WmlPackage struct {
	*opc.OpcPackage
	imageParts *ImageParts // lazy-initialized
}

// NewWmlPackage wraps an OpcPackage with WML behaviours.
func NewWmlPackage(pkg *opc.OpcPackage) *WmlPackage {
	return &WmlPackage{
		OpcPackage: pkg,
		imageParts: NewImageParts(),
	}
}

// ImageParts returns the ImageParts collection for this package.
func (wp *WmlPackage) ImageParts() *ImageParts {
	return wp.imageParts
}

// GetOrAddImagePart returns the ImagePart matching the given image part
// (by SHA-256 dedup), creating a new one if needed.
//
// Mirrors Python Package.get_or_add_image_part.
// In Python this takes an image_descriptor; here we take a pre-built
// ImagePart (with blob, metadata) since the Image parsing is in MR-10.
func (wp *WmlPackage) GetOrAddImagePart(ip *ImagePart) (*ImagePart, error) {
	hash, err := ip.Hash()
	if err != nil {
		return nil, fmt.Errorf("parts: hashing new image part: %w", err)
	}
	existing, err := wp.ImageParts().GetByHash(hash)
	if err != nil {
		return nil, fmt.Errorf("parts: searching existing image parts: %w", err)
	}
	if existing != nil {
		return existing, nil
	}
	// Assign a new partname
	ext := ip.PartName().Ext()
	if ext == "" {
		ext = extFromContentType(ip.ContentType())
	}
	pn := wp.ImageParts().nextImagePartname(ext)
	ip.SetPartName(pn)
	wp.ImageParts().Append(ip)
	wp.OpcPackage.AddPart(ip)
	return ip, nil
}

// AfterUnmarshal gathers existing image parts from all relationships.
//
// Mirrors Python Package.after_unmarshal → _gather_image_parts.
func (wp *WmlPackage) AfterUnmarshal() {
	for _, rel := range wp.OpcPackage.IterRels() {
		if rel.IsExternal {
			continue
		}
		if rel.RelType != opc.RTImage {
			continue
		}
		if rel.TargetPart == nil {
			continue
		}
		ip, ok := rel.TargetPart.(*ImagePart)
		if !ok {
			continue
		}
		if wp.ImageParts().Contains(ip) {
			continue
		}
		wp.ImageParts().Append(ip)
	}
}

// --------------------------------------------------------------------------
// ImageParts — collection with SHA-256 deduplication
// --------------------------------------------------------------------------

// ImageParts is a collection of ImagePart objects with SHA-256-based
// deduplication.
//
// Mirrors Python ImageParts exactly.
type ImageParts struct {
	parts []*ImagePart
}

// NewImageParts creates an empty ImageParts collection.
func NewImageParts() *ImageParts {
	return &ImageParts{}
}

// Append adds an ImagePart to the collection.
func (ips *ImageParts) Append(ip *ImagePart) {
	ips.parts = append(ips.parts, ip)
}

// Contains returns true if ip is already in the collection (by pointer identity).
func (ips *ImageParts) Contains(ip *ImagePart) bool {
	for _, p := range ips.parts {
		if p == ip {
			return true
		}
	}
	return false
}

// Len returns the number of image parts in the collection.
func (ips *ImageParts) Len() int {
	return len(ips.parts)
}

// All returns all image parts.
func (ips *ImageParts) All() []*ImagePart {
	return ips.parts
}

// GetByHash returns the image part with a matching hash, or nil.
//
// Mirrors Python ImageParts._get_by_sha1 (upgraded to SHA-256).
func (ips *ImageParts) GetByHash(hash string) (*ImagePart, error) {
	for _, ip := range ips.parts {
		h, err := ip.Hash()
		if err != nil {
			return nil, fmt.Errorf("parts: computing hash for dedup: %w", err)
		}
		if h == hash {
			return ip, nil
		}
	}
	return nil, nil
}

// nextImagePartname returns the next available image partname starting from
// /word/media/image1.{ext}, reusing unused numbers.
//
// Mirrors Python ImageParts._next_image_partname EXACTLY:
// starts from 1, reuses unused numbers (gaps).
func (ips *ImageParts) nextImagePartname(ext string) opc.PackURI {
	usedNumbers := make(map[int]bool)
	for _, ip := range ips.parts {
		if idx, ok := ip.PartName().Idx(); ok {
			usedNumbers[idx] = true
		}
	}
	// Reuse gaps: try 1..len, then len+1
	for n := 1; n <= len(ips.parts); n++ {
		if !usedNumbers[n] {
			return opc.PackURI(fmt.Sprintf("/word/media/image%d.%s", n, ext))
		}
	}
	return opc.PackURI(fmt.Sprintf("/word/media/image%d.%s", len(ips.parts)+1, ext))
}

// extFromContentType derives a file extension from a MIME content type.
func extFromContentType(ct string) string {
	switch {
	case strings.Contains(ct, "jpeg"):
		return "jpg"
	case strings.Contains(ct, "png"):
		return "png"
	case strings.Contains(ct, "gif"):
		return "gif"
	case strings.Contains(ct, "tiff"):
		return "tiff"
	case strings.Contains(ct, "bmp"):
		return "bmp"
	default:
		return "bin"
	}
}
