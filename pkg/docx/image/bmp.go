package image

import (
	"io"
	"math"
)

// bmpHeader holds parsed properties of a BMP image.
type bmpHeader struct {
	pxWidth_  int
	pxHeight_ int
	horzDpi_  int
	vertDpi_  int
}

func (h *bmpHeader) ContentType() string { return MimeBMP }
func (h *bmpHeader) DefaultExt() string  { return "bmp" }
func (h *bmpHeader) PxWidth() int        { return h.pxWidth_ }
func (h *bmpHeader) PxHeight() int       { return h.pxHeight_ }
func (h *bmpHeader) HorzDpi() int        { return h.horzDpi_ }
func (h *bmpHeader) VertDpi() int        { return h.vertDpi_ }

// parseBMP parses a BMP image stream and returns its header properties.
// Mirrors Python Bmp.from_stream exactly.
func parseBMP(stream io.ReadSeeker) (imageHeader, error) {
	sr := NewStreamReader(stream, false, 0) // little-endian

	pxWidth, err := sr.ReadLong(0x12, 0)
	if err != nil {
		return nil, err
	}
	pxHeight, err := sr.ReadLong(0x16, 0)
	if err != nil {
		return nil, err
	}
	horzPxPerMeter, err := sr.ReadLong(0x26, 0)
	if err != nil {
		return nil, err
	}
	vertPxPerMeter, err := sr.ReadLong(0x2A, 0)
	if err != nil {
		return nil, err
	}

	return &bmpHeader{
		pxWidth_:  int(pxWidth),
		pxHeight_: int(pxHeight),
		horzDpi_:  bmpDpi(horzPxPerMeter),
		vertDpi_:  bmpDpi(vertPxPerMeter),
	}, nil
}

// bmpDpi returns dots per inch from pixels per meter, defaulting to 96
// if pxPerMeter is zero. Mirrors Python Bmp._dpi exactly.
func bmpDpi(pxPerMeter uint32) int {
	if pxPerMeter == 0 {
		return 96
	}
	return int(math.Round(float64(pxPerMeter) * 0.0254))
}
