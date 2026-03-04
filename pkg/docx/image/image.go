package image

import (
	"bytes"
	"crypto/sha256"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
)

// imageHeader is the interface for format-specific image header parsers.
// Mirrors Python BaseImageHeader abstract class.
type imageHeader interface {
	ContentType() string
	DefaultExt() string
	PxWidth() int
	PxHeight() int
	HorzDpi() int
	VertDpi() int
}

// Image represents a graphical image stream such as JPEG, PNG, or GIF with
// properties and methods required by ImagePart. Mirrors Python Image class.
type Image struct {
	blob        []byte
	filename    string
	contentType string
	pxWidth     int
	pxHeight    int
	horzDpi     int
	vertDpi     int
	hash        string // lazy, "" until first Hash() call
}

// FromBlob returns a new Image parsed from the image binary in blob.
// If filename is empty, a default name like "image.png" is generated.
//
// Note: Python Image.from_blob(blob) has no filename parameter — it always
// generates a default name. The filename argument here is a Go-specific
// addition for caller convenience (used by ImagePart).
func FromBlob(blob []byte, filename string) (*Image, error) {
	stream := bytes.NewReader(blob)
	return fromStream(stream, blob, filename)
}

// FromFile returns a new Image loaded from the file at path.
// Mirrors Python Image.from_file(path).
func FromFile(path string) (*Image, error) {
	blob, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("image: reading file %q: %w", path, err)
	}
	stream := bytes.NewReader(blob)
	filename := filepath.Base(path)
	return fromStream(stream, blob, filename)
}

// FromReadSeeker returns a new Image loaded from a stream.
// Mirrors Python Image.from_file(stream).
func FromReadSeeker(r io.ReadSeeker, filename string) (*Image, error) {
	if _, err := r.Seek(0, io.SeekStart); err != nil {
		return nil, fmt.Errorf("image: seeking stream: %w", err)
	}
	blob, err := io.ReadAll(r)
	if err != nil {
		return nil, fmt.Errorf("image: reading stream: %w", err)
	}
	stream := bytes.NewReader(blob)
	return fromStream(stream, blob, filename)
}

// fromStream creates an Image by detecting the format and parsing header properties.
// Mirrors Python Image._from_stream.
func fromStream(stream io.ReadSeeker, blob []byte, filename string) (*Image, error) {
	header, err := imageHeaderFactory(stream)
	if err != nil {
		return nil, err
	}
	if filename == "" {
		filename = "image." + header.DefaultExt()
	}
	return &Image{
		blob:        blob,
		filename:    filename,
		contentType: header.ContentType(),
		pxWidth:     header.PxWidth(),
		pxHeight:    header.PxHeight(),
		horzDpi:     header.HorzDpi(),
		vertDpi:     header.VertDpi(),
	}, nil
}

// Blob returns the bytes of the image file.
func (img *Image) Blob() []byte { return img.blob }

// Filename returns the original image filename.
func (img *Image) Filename() string { return img.filename }

// ContentType returns the MIME content type, e.g. "image/jpeg".
func (img *Image) ContentType() string { return img.contentType }

// PxWidth returns the horizontal pixel dimension of the image.
func (img *Image) PxWidth() int { return img.pxWidth }

// PxHeight returns the vertical pixel dimension of the image.
func (img *Image) PxHeight() int { return img.pxHeight }

// HorzDpi returns horizontal dots per inch. Defaults to 72 when not
// present in the file.
func (img *Image) HorzDpi() int { return img.horzDpi }

// VertDpi returns vertical dots per inch. Defaults to 72 when not
// present in the file.
func (img *Image) VertDpi() int { return img.vertDpi }

// Ext returns the file extension without the leading dot, e.g. "jpg".
// Mirrors Python Image.ext: os.path.splitext(filename)[1][1:]
func (img *Image) Ext() string {
	ext := filepath.Ext(img.filename)
	if ext == "" {
		return ""
	}
	return ext[1:] // remove leading dot, preserve original case
}

// Width returns the native width in EMU, calculated from px_width and horz_dpi.
// Mirrors Python Image.width → Inches(px_width / horz_dpi).
func (img *Image) Width() int64 {
	return int64(float64(img.pxWidth) / float64(img.horzDpi) * 914400)
}

// Height returns the native height in EMU, calculated from px_height and vert_dpi.
// Mirrors Python Image.height → Inches(px_height / vert_dpi).
func (img *Image) Height() int64 {
	return int64(float64(img.pxHeight) / float64(img.vertDpi) * 914400)
}

// ScaledDimensions returns (cx, cy) in EMU representing scaled dimensions.
//
// Rules (mirrors Python Image.scaled_dimensions exactly):
//   - Both nil → native dimensions
//   - width nil → scale width proportionally from height
//   - height nil → scale height proportionally from width
//   - Both non-nil → return as-is
func (img *Image) ScaledDimensions(width, height *int64) (cx, cy int64) {
	if width == nil && height == nil {
		return img.Width(), img.Height()
	}

	if width == nil {
		scalingFactor := float64(*height) / float64(img.Height())
		w := int64(math.Round(float64(img.Width()) * scalingFactor))
		return w, *height
	}

	if height == nil {
		scalingFactor := float64(*width) / float64(img.Width())
		h := int64(math.Round(float64(img.Height()) * scalingFactor))
		return *width, h
	}

	return *width, *height
}

// Hash returns the hex-encoded SHA-256 hash of the image blob. The value is
// computed lazily and cached.
func (img *Image) Hash() string {
	if img.hash == "" {
		sum := sha256.Sum256(img.blob)
		img.hash = fmt.Sprintf("%x", sum)
	}
	return img.hash
}

// signatures is the table of magic bytes for recognizing image formats.
// Mirrors Python SIGNATURES exactly.
var signatures = []struct {
	parser func(io.ReadSeeker) (imageHeader, error)
	offset int
	magic  []byte
}{
	{parsePNG, 0, []byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A}},
	{parseJFIF, 6, []byte("JFIF")},
	{parseExif, 6, []byte("Exif")},
	{parseGIF, 0, []byte("GIF87a")},
	{parseGIF, 0, []byte("GIF89a")},
	{parseTIFF, 0, []byte{0x4D, 0x4D, 0x00, 0x2A}}, // big-endian
	{parseTIFF, 0, []byte{0x49, 0x49, 0x2A, 0x00}}, // little-endian
	{parseBMP, 0, []byte("BM")},
}

// imageHeaderFactory returns the appropriate imageHeader by matching magic bytes.
// Mirrors Python _ImageHeaderFactory exactly.
func imageHeaderFactory(stream io.ReadSeeker) (imageHeader, error) {
	if _, err := stream.Seek(0, io.SeekStart); err != nil {
		return nil, fmt.Errorf("image: seeking to start: %w", err)
	}
	var header [32]byte
	n, err := stream.Read(header[:])
	if err != nil && err != io.EOF {
		return nil, fmt.Errorf("image: reading header bytes: %w", err)
	}

	for _, sig := range signatures {
		end := sig.offset + len(sig.magic)
		if end > n {
			continue
		}
		if bytes.Equal(header[sig.offset:end], sig.magic) {
			return sig.parser(stream)
		}
	}

	return nil, ErrUnrecognizedImage
}
