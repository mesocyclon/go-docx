package image

import (
	"bytes"
	"fmt"
	"io"
	"math"
)

// jpegHeader holds parsed properties of a JPEG image.
type jpegHeader struct {
	contentType_ string
	defaultExt_  string
	pxWidth_     int
	pxHeight_    int
	horzDpi_     int
	vertDpi_     int
}

func (h *jpegHeader) ContentType() string { return h.contentType_ }
func (h *jpegHeader) DefaultExt() string  { return h.defaultExt_ }
func (h *jpegHeader) PxWidth() int        { return h.pxWidth_ }
func (h *jpegHeader) PxHeight() int       { return h.pxHeight_ }
func (h *jpegHeader) HorzDpi() int        { return h.horzDpi_ }
func (h *jpegHeader) VertDpi() int        { return h.vertDpi_ }

// parseJFIF parses a JFIF JPEG image stream.
// Mirrors Python Jfif.from_stream exactly.
func parseJFIF(stream io.ReadSeeker) (imageHeader, error) {
	markers, err := parseJfifMarkers(stream)
	if err != nil {
		return nil, err
	}

	sof, err := markers.sof()
	if err != nil {
		return nil, err
	}
	app0, err := markers.app0()
	if err != nil {
		return nil, err
	}

	return &jpegHeader{
		contentType_: MimeJPEG,
		defaultExt_:  "jpg",
		pxWidth_:     sof.pxWidth,
		pxHeight_:    sof.pxHeight,
		horzDpi_:     app0.horzDpi(),
		vertDpi_:     app0.vertDpi(),
	}, nil
}

// parseExif parses an Exif JPEG image stream.
// Mirrors Python Exif.from_stream exactly.
func parseExif(stream io.ReadSeeker) (imageHeader, error) {
	markers, err := parseJfifMarkers(stream)
	if err != nil {
		return nil, err
	}

	sof, err := markers.sof()
	if err != nil {
		return nil, err
	}
	app1, err := markers.app1()
	if err != nil {
		return nil, err
	}

	return &jpegHeader{
		contentType_: MimeJPEG,
		defaultExt_:  "jpg",
		pxWidth_:     sof.pxWidth,
		pxHeight_:    sof.pxHeight,
		horzDpi_:     app1.horzDpi,
		vertDpi_:     app1.vertDpi,
	}, nil
}

// --- Marker collection ---

// jfifMarkers is a sequence of markers in a JPEG file, truncated at the
// first SOS marker. Mirrors Python _JfifMarkers.
type jfifMarkers struct {
	markers []jfifMarker
}

// jfifMarker is the interface for all JPEG marker types.
type jfifMarker interface {
	markerCode() byte
	segmentLength() int
}

// parseJfifMarkers scans the JPEG stream for markers up to and including SOS.
// Mirrors Python _JfifMarkers.from_stream exactly.
func parseJfifMarkers(stream io.ReadSeeker) (*jfifMarkers, error) {
	sr := NewStreamReader(stream, true, 0) // JPEG is big-endian
	finder := &markerFinder{stream: sr}

	var markers []jfifMarker
	start := int64(0)

	for {
		code, segmentOffset, err := finder.next(start)
		if err != nil {
			return nil, err
		}

		marker, err := markerFactory(code, sr, segmentOffset)
		if err != nil {
			return nil, err
		}

		markers = append(markers, marker)

		if code == markerSOS || code == markerEOI {
			break
		}
		start = segmentOffset + int64(marker.segmentLength())
	}

	return &jfifMarkers{markers: markers}, nil
}

// sof returns the first SOF (start of frame) marker.
// Mirrors Python _JfifMarkers.sof.
func (jm *jfifMarkers) sof() (*sofMarker, error) {
	for _, m := range jm.markers {
		if sofMarkerCodes[m.markerCode()] {
			if sm, ok := m.(*sofMarker); ok {
				return sm, nil
			}
		}
	}
	return nil, fmt.Errorf("image/jpeg: no start of frame (SOFn) marker in image")
}

// app0 returns the first APP0 marker.
// Mirrors Python _JfifMarkers.app0.
func (jm *jfifMarkers) app0() (*app0Marker, error) {
	for _, m := range jm.markers {
		if m.markerCode() == markerAPP0 {
			if am, ok := m.(*app0Marker); ok {
				return am, nil
			}
		}
	}
	return nil, fmt.Errorf("image/jpeg: no APP0 marker in image")
}

// app1 returns the first APP1 marker.
// Mirrors Python _JfifMarkers.app1.
func (jm *jfifMarkers) app1() (*app1Marker, error) {
	for _, m := range jm.markers {
		if m.markerCode() == markerAPP1 {
			if am, ok := m.(*app1Marker); ok {
				return am, nil
			}
		}
	}
	return nil, fmt.Errorf("image/jpeg: no APP1 marker in image")
}

// --- Marker Finder ---

// markerFinder scans a JPEG stream for the next marker.
// Mirrors Python _MarkerFinder.
type markerFinder struct {
	stream *StreamReader
}

// next returns (markerCode, segmentOffset) for the next marker after start.
// segmentOffset points to the position immediately after the 2-byte marker code.
// Mirrors Python _MarkerFinder.next exactly.
func (mf *markerFinder) next(start int64) (byte, int64, error) {
	position := start
	for {
		// Skip to next 0xFF byte
		ffPos, err := mf.offsetOfNextFFByte(position)
		if err != nil {
			return 0, 0, err
		}

		// Skip over any 0xFF padding bytes, get the non-FF byte
		nonFFPos, b, err := mf.nextNonFFByte(ffPos + 1)
		if err != nil {
			return 0, 0, err
		}

		// FF 00 is not a marker, restart scan
		if b == 0x00 {
			position = nonFFPos + 1
			continue
		}

		// This is a real marker
		return b, nonFFPos + 1, nil
	}
}

// offsetOfNextFFByte returns the offset of the next 0xFF byte starting at start.
// Mirrors Python _MarkerFinder._offset_of_next_ff_byte.
func (mf *markerFinder) offsetOfNextFFByte(start int64) (int64, error) {
	if err := mf.stream.SeekTo(start, 0); err != nil {
		return 0, err
	}
	for {
		b, err := mf.readByte()
		if err != nil {
			return 0, err
		}
		if b == 0xFF {
			pos, err := mf.stream.Tell()
			if err != nil {
				return 0, err
			}
			return pos - 1, nil
		}
	}
}

// nextNonFFByte returns (offset, byte) of the next non-0xFF byte starting at start.
// Mirrors Python _MarkerFinder._next_non_ff_byte.
func (mf *markerFinder) nextNonFFByte(start int64) (int64, byte, error) {
	if err := mf.stream.SeekTo(start, 0); err != nil {
		return 0, 0, err
	}
	for {
		b, err := mf.readByte()
		if err != nil {
			return 0, 0, err
		}
		if b != 0xFF {
			pos, err := mf.stream.Tell()
			if err != nil {
				return 0, 0, err
			}
			return pos - 1, b, nil
		}
	}
}

func (mf *markerFinder) readByte() (byte, error) {
	var buf [1]byte
	n, err := mf.stream.Read(buf[:])
	if n == 0 {
		if err != nil {
			return 0, fmt.Errorf("image/jpeg: %w: scanning for marker", ErrUnexpectedEOF)
		}
		return 0, fmt.Errorf("image/jpeg: %w: scanning for marker", ErrUnexpectedEOF)
	}
	return buf[0], nil
}

// --- Marker Factory ---

// markerFactory creates the appropriate marker type for the given code.
// Mirrors Python _MarkerFactory exactly.
func markerFactory(code byte, sr *StreamReader, offset int64) (jfifMarker, error) {
	switch {
	case code == markerAPP0:
		return parseApp0Marker(code, sr, offset)
	case code == markerAPP1:
		return parseApp1Marker(code, sr, offset)
	case sofMarkerCodes[code]:
		return parseSofMarker(code, sr, offset)
	default:
		return parseBaseMarker(code, sr, offset)
	}
}

// --- Base Marker ---

// baseMarker is the default marker type with no special parsing.
// Mirrors Python _Marker.
type baseMarker struct {
	code_          byte
	offset_        int64
	segmentLength_ int
}

func (m *baseMarker) markerCode() byte   { return m.code_ }
func (m *baseMarker) segmentLength() int { return m.segmentLength_ }

// parseBaseMarker creates a generic marker.
// Mirrors Python _Marker.from_stream.
func parseBaseMarker(code byte, sr *StreamReader, offset int64) (*baseMarker, error) {
	segLen := 0
	if !standaloneMarkers[code] {
		v, err := sr.ReadShort(offset, 0)
		if err != nil {
			return nil, fmt.Errorf("image/jpeg: reading marker segment length: %w", err)
		}
		segLen = int(v)
	}
	return &baseMarker{code_: code, offset_: offset, segmentLength_: segLen}, nil
}

// --- SOF Marker ---

// sofMarker holds data from a JPEG Start-Of-Frame marker.
// Mirrors Python _SofMarker.
type sofMarker struct {
	baseMarker
	pxWidth  int
	pxHeight int
}

// parseSofMarker parses an SOFn marker.
// Mirrors Python _SofMarker.from_stream exactly.
func parseSofMarker(code byte, sr *StreamReader, offset int64) (*sofMarker, error) {
	segLen, err := sr.ReadShort(offset, 0)
	if err != nil {
		return nil, err
	}
	pxHeight, err := sr.ReadShort(offset, 3)
	if err != nil {
		return nil, err
	}
	pxWidth, err := sr.ReadShort(offset, 5)
	if err != nil {
		return nil, err
	}
	return &sofMarker{
		baseMarker: baseMarker{code_: code, offset_: offset, segmentLength_: int(segLen)},
		pxWidth:    int(pxWidth),
		pxHeight:   int(pxHeight),
	}, nil
}

// --- APP0 Marker (JFIF) ---

// app0Marker holds JFIF DPI data from an APP0 marker.
// Mirrors Python _App0Marker.
type app0Marker struct {
	baseMarker
	densityUnits byte
	xDensity     uint16
	yDensity     uint16
}

// parseApp0Marker parses an APP0 (JFIF) marker.
// Mirrors Python _App0Marker.from_stream exactly.
func parseApp0Marker(code byte, sr *StreamReader, offset int64) (*app0Marker, error) {
	segLen, err := sr.ReadShort(offset, 0)
	if err != nil {
		return nil, err
	}
	densityUnits, err := sr.ReadByteAt(offset, 9)
	if err != nil {
		return nil, err
	}
	xDensity, err := sr.ReadShort(offset, 10)
	if err != nil {
		return nil, err
	}
	yDensity, err := sr.ReadShort(offset, 12)
	if err != nil {
		return nil, err
	}
	return &app0Marker{
		baseMarker:   baseMarker{code_: code, offset_: offset, segmentLength_: int(segLen)},
		densityUnits: densityUnits,
		xDensity:     xDensity,
		yDensity:     yDensity,
	}, nil
}

// horzDpi returns horizontal DPI from JFIF APP0 data.
// Mirrors Python _App0Marker.horz_dpi / _dpi exactly.
func (m *app0Marker) horzDpi() int {
	return app0Dpi(m.densityUnits, m.xDensity)
}

// vertDpi returns vertical DPI from JFIF APP0 data.
// Mirrors Python _App0Marker.vert_dpi / _dpi exactly.
func (m *app0Marker) vertDpi() int {
	return app0Dpi(m.densityUnits, m.yDensity)
}

// app0Dpi converts a JFIF density value to DPI.
// Mirrors Python _App0Marker._dpi exactly.
func app0Dpi(densityUnits byte, density uint16) int {
	switch densityUnits {
	case 1: // dots per inch
		return int(density)
	case 2: // dots per centimeter
		return int(math.Round(float64(density) * 2.54))
	default: // no unit or unknown
		return 72
	}
}

// --- APP1 Marker (Exif) ---

// app1Marker holds Exif DPI data from an APP1 marker.
// Mirrors Python _App1Marker.
type app1Marker struct {
	baseMarker
	horzDpi int
	vertDpi int
}

// parseApp1Marker parses an APP1 (Exif) marker.
// Mirrors Python _App1Marker.from_stream exactly.
func parseApp1Marker(code byte, sr *StreamReader, offset int64) (*app1Marker, error) {
	segLen, err := sr.ReadShort(offset, 0)
	if err != nil {
		return nil, err
	}

	horzDpi, vertDpi := 72, 72

	if !isNonExifApp1(sr, offset) {
		tiffHdr, err := tiffFromExifSegment(sr, offset, int(segLen))
		if err != nil {
			return nil, fmt.Errorf("image/jpeg: parsing Exif TIFF IFD: %w", err)
		}
		horzDpi = tiffHdr.HorzDpi()
		vertDpi = tiffHdr.VertDpi()
	}

	return &app1Marker{
		baseMarker: baseMarker{code_: code, offset_: offset, segmentLength_: int(segLen)},
		horzDpi:    horzDpi,
		vertDpi:    vertDpi,
	}, nil
}

// isNonExifApp1 checks if the APP1 segment is NOT an Exif segment.
// Mirrors Python _App1Marker._is_non_Exif_APP1_segment.
func isNonExifApp1(sr *StreamReader, offset int64) bool {
	if err := sr.SeekTo(offset+2, 0); err != nil {
		return true
	}
	var sig [6]byte
	n, err := sr.Read(sig[:])
	if err != nil || n < 6 {
		return true
	}
	return !bytes.Equal(sig[:], []byte("Exif\x00\x00"))
}

// tiffFromExifSegment extracts a TIFF IFD from an Exif APP1 segment.
// Mirrors Python _App1Marker._tiff_from_exif_segment exactly.
func tiffFromExifSegment(sr *StreamReader, offset int64, segmentLength int) (imageHeader, error) {
	// Seek to the TIFF data within the APP1 segment (offset + 8 skips
	// the 2-byte segment length and 6-byte "Exif\x00\x00" signature)
	if err := sr.SeekTo(offset+8, 0); err != nil {
		return nil, err
	}

	tiffLen := segmentLength - 8
	if tiffLen <= 0 {
		return nil, fmt.Errorf("image/jpeg: Exif segment too short")
	}

	buf := make([]byte, tiffLen)
	n, err := sr.Read(buf)
	if err != nil || n < tiffLen {
		return nil, fmt.Errorf("image/jpeg: reading Exif TIFF data: %w", ErrUnexpectedEOF)
	}

	substream := bytes.NewReader(buf)
	return parseTIFF(substream)
}
