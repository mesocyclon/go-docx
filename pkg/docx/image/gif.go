package image

import (
	"encoding/binary"
	"fmt"
	"io"
)

// gifHeader holds parsed properties of a GIF image.
// GIF does not support DPI metadata; both default to 72.
type gifHeader struct {
	pxWidth_  int
	pxHeight_ int
}

func (h *gifHeader) ContentType() string { return MimeGIF }
func (h *gifHeader) DefaultExt() string  { return "gif" }
func (h *gifHeader) PxWidth() int        { return h.pxWidth_ }
func (h *gifHeader) PxHeight() int       { return h.pxHeight_ }
func (h *gifHeader) HorzDpi() int        { return 72 }
func (h *gifHeader) VertDpi() int        { return 72 }

// parseGIF parses a GIF image stream and returns its header properties.
// Mirrors Python Gif.from_stream exactly.
func parseGIF(stream io.ReadSeeker) (imageHeader, error) {
	// Seek past the 6-byte GIF signature ("GIF87a" or "GIF89a")
	if _, err := stream.Seek(6, io.SeekStart); err != nil {
		return nil, fmt.Errorf("image/gif: seek: %w", err)
	}
	var dims [4]byte
	if _, err := io.ReadFull(stream, dims[:]); err != nil {
		return nil, fmt.Errorf("%w: reading GIF dimensions", ErrUnexpectedEOF)
	}
	pxWidth := int(binary.LittleEndian.Uint16(dims[0:2]))
	pxHeight := int(binary.LittleEndian.Uint16(dims[2:4]))
	return &gifHeader{pxWidth_: pxWidth, pxHeight_: pxHeight}, nil
}
