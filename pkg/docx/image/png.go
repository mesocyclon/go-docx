package image

import (
	"fmt"
	"io"
	"math"
)

// pngHeader holds parsed properties of a PNG image.
type pngHeader struct {
	pxWidth_  int
	pxHeight_ int
	horzDpi_  int
	vertDpi_  int
}

func (h *pngHeader) ContentType() string { return MimePNG }
func (h *pngHeader) DefaultExt() string  { return "png" }
func (h *pngHeader) PxWidth() int        { return h.pxWidth_ }
func (h *pngHeader) PxHeight() int       { return h.pxHeight_ }
func (h *pngHeader) HorzDpi() int        { return h.horzDpi_ }
func (h *pngHeader) VertDpi() int        { return h.vertDpi_ }

// parsePNG parses a PNG image stream and returns its header properties.
// Mirrors Python Png.from_stream exactly.
func parsePNG(stream io.ReadSeeker) (imageHeader, error) {
	chunks, err := parsePNGChunks(stream)
	if err != nil {
		return nil, err
	}

	ihdr := chunks.ihdr
	if ihdr == nil {
		return nil, fmt.Errorf("%w: no IHDR chunk in PNG image", ErrInvalidImageStream)
	}

	horzDpi := 72
	vertDpi := 72
	if chunks.phys != nil {
		horzDpi = pngDpi(chunks.phys.unitsSpecifier, chunks.phys.horzPxPerUnit)
		vertDpi = pngDpi(chunks.phys.unitsSpecifier, chunks.phys.vertPxPerUnit)
	}

	return &pngHeader{
		pxWidth_:  int(ihdr.pxWidth),
		pxHeight_: int(ihdr.pxHeight),
		horzDpi_:  horzDpi,
		vertDpi_:  vertDpi,
	}, nil
}

// pngDpi returns dots per inch calculated from unit specifier and pixels per unit.
// Mirrors Python _PngParser._dpi exactly.
func pngDpi(unitsSpecifier byte, pxPerUnit uint32) int {
	if unitsSpecifier == 1 && pxPerUnit > 0 {
		return int(math.Round(float64(pxPerUnit) * 0.0254))
	}
	return 72
}

// pngChunks holds the interesting chunks parsed from a PNG stream.
type pngChunks struct {
	ihdr *ihdrChunk
	phys *physChunk
}

// ihdrChunk holds data from the PNG IHDR chunk.
type ihdrChunk struct {
	pxWidth  uint32
	pxHeight uint32
}

// physChunk holds data from the PNG pHYs chunk.
type physChunk struct {
	horzPxPerUnit  uint32
	vertPxPerUnit  uint32
	unitsSpecifier byte
}

// parsePNGChunks iterates over PNG chunks, extracting IHDR and pHYs data.
// Mirrors Python _Chunks.from_stream + _ChunkParser + _ChunkFactory.
func parsePNGChunks(stream io.ReadSeeker) (*pngChunks, error) {
	sr := NewStreamReader(stream, true, 0) // PNG is big-endian
	result := &pngChunks{}

	// Start after the 8-byte PNG signature
	chunkOffset := int64(8)

	for {
		// Read chunk data length (4 bytes) and type (4 bytes)
		chunkDataLen, err := sr.ReadLong(chunkOffset, 0)
		if err != nil {
			return nil, fmt.Errorf("image/png: reading chunk header: %w", err)
		}
		chunkType, err := sr.ReadStr(4, chunkOffset, 4)
		if err != nil {
			return nil, fmt.Errorf("image/png: reading chunk type: %w", err)
		}

		dataOffset := chunkOffset + 8

		switch chunkType {
		case pngChunkIHDR:
			w, err := sr.ReadLong(dataOffset, 0)
			if err != nil {
				return nil, fmt.Errorf("image/png: reading IHDR width: %w", err)
			}
			h, err := sr.ReadLong(dataOffset, 4)
			if err != nil {
				return nil, fmt.Errorf("image/png: reading IHDR height: %w", err)
			}
			result.ihdr = &ihdrChunk{pxWidth: w, pxHeight: h}

		case pngChunkPHYs:
			hppu, err := sr.ReadLong(dataOffset, 0)
			if err != nil {
				return nil, fmt.Errorf("image/png: reading pHYs horz: %w", err)
			}
			vppu, err := sr.ReadLong(dataOffset, 4)
			if err != nil {
				return nil, fmt.Errorf("image/png: reading pHYs vert: %w", err)
			}
			us, err := sr.ReadByteAt(dataOffset, 8)
			if err != nil {
				return nil, fmt.Errorf("image/png: reading pHYs units: %w", err)
			}
			result.phys = &physChunk{
				horzPxPerUnit:  hppu,
				vertPxPerUnit:  vppu,
				unitsSpecifier: us,
			}
		}

		if chunkType == pngChunkIEND {
			break
		}

		// Advance: 4 (length) + 4 (type) + data + 4 (CRC)
		chunkOffset += 4 + 4 + int64(chunkDataLen) + 4
	}

	return result, nil
}
