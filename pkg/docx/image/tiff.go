package image

import (
	"fmt"
	"io"
	"math"
)

// tiffHeader holds parsed properties of a TIFF image.
type tiffHeader struct {
	pxWidth_  int
	pxHeight_ int
	horzDpi_  int
	vertDpi_  int
}

func (h *tiffHeader) ContentType() string { return MimeTIFF }
func (h *tiffHeader) DefaultExt() string  { return "tiff" }
func (h *tiffHeader) PxWidth() int        { return h.pxWidth_ }
func (h *tiffHeader) PxHeight() int       { return h.pxHeight_ }
func (h *tiffHeader) HorzDpi() int        { return h.horzDpi_ }
func (h *tiffHeader) VertDpi() int        { return h.vertDpi_ }

// parseTIFF parses a TIFF image stream and returns its header properties.
// Handles both big-endian (MM) and little-endian (II) byte orderings.
// Mirrors Python Tiff.from_stream exactly.
func parseTIFF(stream io.ReadSeeker) (imageHeader, error) {
	parser, err := newTiffParser(stream)
	if err != nil {
		return nil, err
	}
	return &tiffHeader{
		pxWidth_:  parser.pxWidth(),
		pxHeight_: parser.pxHeight(),
		horzDpi_:  parser.horzDpi(),
		vertDpi_:  parser.vertDpi(),
	}, nil
}

// tiffParser parses a TIFF image stream to extract image properties from
// its main image file directory (IFD). Mirrors Python _TiffParser.
type tiffParser struct {
	entries *ifdEntries
}

// newTiffParser creates a tiffParser from the given stream.
// Mirrors Python _TiffParser.parse.
func newTiffParser(stream io.ReadSeeker) (*tiffParser, error) {
	bigEndian, err := detectTiffEndian(stream)
	if err != nil {
		return nil, err
	}
	sr := NewStreamReader(stream, bigEndian, 0)

	ifd0Offset, err := sr.ReadLong(4, 0)
	if err != nil {
		return nil, fmt.Errorf("image/tiff: reading IFD0 offset: %w", err)
	}

	entries, err := parseIfdEntries(sr, int64(ifd0Offset))
	if err != nil {
		return nil, fmt.Errorf("image/tiff: parsing IFD entries: %w", err)
	}
	return &tiffParser{entries: entries}, nil
}

func (tp *tiffParser) pxWidth() int {
	v, ok := tp.entries.get(tiffTagImageWidth)
	if !ok {
		return 0
	}
	return int(v.(int64))
}

func (tp *tiffParser) pxHeight() int {
	v, ok := tp.entries.get(tiffTagImageLength)
	if !ok {
		return 0
	}
	return int(v.(int64))
}

func (tp *tiffParser) horzDpi() int {
	return tp.dpi(tiffTagXResolution)
}

func (tp *tiffParser) vertDpi() int {
	return tp.dpi(tiffTagYResolution)
}

// dpi calculates dots per inch for the given resolution tag.
// Mirrors Python _TiffParser._dpi exactly.
func (tp *tiffParser) dpi(resolutionTag uint16) int {
	if !tp.entries.contains(resolutionTag) {
		return 72
	}

	// resolution unit defaults to 2 (inches)
	resolutionUnit := int64(2)
	if v, ok := tp.entries.get(tiffTagResolutionUnit); ok {
		resolutionUnit = v.(int64)
	}

	if resolutionUnit == 1 { // aspect ratio only
		return 72
	}

	// resolution_unit == 2 for inches, 3 for centimeters
	unitsPerInch := 1.0
	if resolutionUnit == 3 {
		unitsPerInch = 2.54
	}

	dotsPerUnit, _ := tp.entries.get(resolutionTag)
	var dpu float64
	switch v := dotsPerUnit.(type) {
	case float64:
		dpu = v
	case int64:
		dpu = float64(v)
	default:
		return 72
	}
	return int(math.Round(dpu * unitsPerInch))
}

// detectTiffEndian reads the first 2 bytes of the stream to determine endianness.
// Returns true for big-endian (MM), false for little-endian (II).
// Mirrors Python _TiffParser._detect_endian.
func detectTiffEndian(stream io.ReadSeeker) (bool, error) {
	if _, err := stream.Seek(0, io.SeekStart); err != nil {
		return false, fmt.Errorf("image/tiff: seek: %w", err)
	}
	var buf [2]byte
	if _, err := io.ReadFull(stream, buf[:]); err != nil {
		return false, fmt.Errorf("%w: reading TIFF endian marker", ErrUnexpectedEOF)
	}
	return string(buf[:]) == "MM", nil
}

// ifdEntries holds the parsed entries from a TIFF IFD.
// Mirrors Python _IfdEntries.
type ifdEntries struct {
	entries map[uint16]interface{}
}

func (e *ifdEntries) contains(tag uint16) bool {
	_, ok := e.entries[tag]
	return ok
}

func (e *ifdEntries) get(tag uint16) (interface{}, bool) {
	v, ok := e.entries[tag]
	return v, ok
}

// parseIfdEntries parses IFD entries from sr at the given offset.
// Mirrors Python _IfdEntries.from_stream.
func parseIfdEntries(sr *StreamReader, offset int64) (*ifdEntries, error) {
	entryCount, err := sr.ReadShort(offset, 0)
	if err != nil {
		return nil, fmt.Errorf("image/tiff: reading IFD entry count: %w", err)
	}

	entries := make(map[uint16]interface{})
	for i := 0; i < int(entryCount); i++ {
		dirEntryOffset := offset + 2 + int64(i)*12
		tag, value, err := parseIfdEntry(sr, dirEntryOffset)
		if err != nil {
			// Skip entries we can't parse (matches Python behavior for
			// unimplemented field types)
			continue
		}
		entries[tag] = value
	}
	return &ifdEntries{entries: entries}, nil
}

// parseIfdEntry parses a single IFD entry at offset in sr.
// Mirrors Python _IfdEntryFactory + _IfdEntry.from_stream and subclasses.
func parseIfdEntry(sr *StreamReader, offset int64) (uint16, interface{}, error) {
	tagCode, err := sr.ReadShort(offset, 0)
	if err != nil {
		return 0, nil, err
	}
	fieldType, err := sr.ReadShort(offset, 2)
	if err != nil {
		return 0, nil, err
	}
	valueCount, err := sr.ReadLong(offset, 4)
	if err != nil {
		return 0, nil, err
	}
	valueOffset, err := sr.ReadLong(offset, 8)
	if err != nil {
		return 0, nil, err
	}

	var value interface{}

	switch fieldType {
	case tiffFieldASCII:
		// Read string at value_offset (length = value_count - 1 for NUL terminator)
		s, err := sr.ReadStr(int(valueCount)-1, int64(valueOffset), 0)
		if err != nil {
			return 0, nil, err
		}
		value = s

	case tiffFieldSHORT:
		if valueCount == 1 {
			// Value is stored inline at offset+8
			v, err := sr.ReadShort(offset, 8)
			if err != nil {
				return 0, nil, err
			}
			value = int64(v)
		} else {
			return 0, nil, fmt.Errorf("image/tiff: multi-value SHORT not implemented")
		}

	case tiffFieldLONG:
		if valueCount == 1 {
			// Value is stored inline at offset+8
			v, err := sr.ReadLong(offset, 8)
			if err != nil {
				return 0, nil, err
			}
			value = int64(v)
		} else {
			return 0, nil, fmt.Errorf("image/tiff: multi-value LONG not implemented")
		}

	case tiffFieldRATIONAL:
		if valueCount == 1 {
			numerator, err := sr.ReadLong(int64(valueOffset), 0)
			if err != nil {
				return 0, nil, err
			}
			denominator, err := sr.ReadLong(int64(valueOffset), 4)
			if err != nil {
				return 0, nil, err
			}
			if denominator == 0 {
				value = float64(0)
			} else {
				value = float64(numerator) / float64(denominator)
			}
		} else {
			return 0, nil, fmt.Errorf("image/tiff: multi-value RATIONAL not implemented")
		}

	default:
		return 0, nil, fmt.Errorf("image/tiff: unimplemented field type %d", fieldType)
	}

	return tagCode, value, nil
}
