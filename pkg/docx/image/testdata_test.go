package image

import (
	"bytes"
	"encoding/binary"
	"hash/crc32"
	"math"
)

// --- Test image builders ---
// These functions generate minimal valid binary images for testing.
// They are placed in a _test.go file to avoid shipping with the library.

// buildMinimalPNG creates a minimal valid PNG with the given dimensions
// and optional pHYs chunk. If pxPerUnit > 0 and unitsSpecifier == 1,
// it represents pixels per meter.
func buildMinimalPNG(width, height uint32, pxPerUnit uint32, unitsSpecifier byte) []byte {
	var buf bytes.Buffer

	// PNG signature
	buf.Write([]byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A})

	// IHDR chunk: width(4) + height(4) + bitDepth(1) + colorType(1) +
	// compression(1) + filter(1) + interlace(1) = 13 bytes
	ihdrData := make([]byte, 13)
	binary.BigEndian.PutUint32(ihdrData[0:], width)
	binary.BigEndian.PutUint32(ihdrData[4:], height)
	ihdrData[8] = 8  // bit depth
	ihdrData[9] = 2  // color type RGB
	ihdrData[10] = 0 // compression
	ihdrData[11] = 0 // filter
	ihdrData[12] = 0 // interlace
	writePNGChunk(&buf, "IHDR", ihdrData)

	// pHYs chunk (optional)
	if pxPerUnit > 0 || unitsSpecifier > 0 {
		physData := make([]byte, 9)
		binary.BigEndian.PutUint32(physData[0:], pxPerUnit)
		binary.BigEndian.PutUint32(physData[4:], pxPerUnit)
		physData[8] = unitsSpecifier
		writePNGChunk(&buf, "pHYs", physData)
	}

	// IDAT chunk (minimal: one row of pixels, compressed)
	// For simplicity, write a minimal valid deflate stream
	// A single-row 1-pixel image with filter byte 0
	idatData := []byte{0x08, 0xD7, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01}
	writePNGChunk(&buf, "IDAT", idatData)

	// IEND chunk
	writePNGChunk(&buf, "IEND", nil)

	return buf.Bytes()
}

// buildMinimalPNGWithSeparateDPI builds a PNG with different horizontal and
// vertical DPI values.
func buildMinimalPNGWithSeparateDPI(width, height uint32, horzPxPerUnit, vertPxPerUnit uint32, unitsSpecifier byte) []byte {
	var buf bytes.Buffer

	buf.Write([]byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A})

	ihdrData := make([]byte, 13)
	binary.BigEndian.PutUint32(ihdrData[0:], width)
	binary.BigEndian.PutUint32(ihdrData[4:], height)
	ihdrData[8] = 8
	ihdrData[9] = 2
	writePNGChunk(&buf, "IHDR", ihdrData)

	physData := make([]byte, 9)
	binary.BigEndian.PutUint32(physData[0:], horzPxPerUnit)
	binary.BigEndian.PutUint32(physData[4:], vertPxPerUnit)
	physData[8] = unitsSpecifier
	writePNGChunk(&buf, "pHYs", physData)

	idatData := []byte{0x08, 0xD7, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01}
	writePNGChunk(&buf, "IDAT", idatData)

	writePNGChunk(&buf, "IEND", nil)

	return buf.Bytes()
}

func writePNGChunk(buf *bytes.Buffer, chunkType string, data []byte) {
	length := make([]byte, 4)
	binary.BigEndian.PutUint32(length, uint32(len(data)))
	buf.Write(length)
	buf.WriteString(chunkType)
	if len(data) > 0 {
		buf.Write(data)
	}
	// CRC over type + data
	crcData := append([]byte(chunkType), data...)
	crc := crc32.ChecksumIEEE(crcData)
	crcBytes := make([]byte, 4)
	binary.BigEndian.PutUint32(crcBytes, crc)
	buf.Write(crcBytes)
}

// buildMinimalJFIF creates a minimal valid JFIF JPEG with given dimensions and DPI.
func buildMinimalJFIF(width, height uint16, densityUnits byte, xDensity, yDensity uint16) []byte {
	var buf bytes.Buffer

	// SOI marker
	buf.Write([]byte{0xFF, 0xD8})

	// APP0 (JFIF) marker
	buf.Write([]byte{0xFF, 0xE0})
	app0 := make([]byte, 16)
	binary.BigEndian.PutUint16(app0[0:], 16) // segment length
	copy(app0[2:7], "JFIF\x00")              // identifier
	app0[7] = 1                              // major version
	app0[8] = 1                              // minor version
	app0[9] = densityUnits
	binary.BigEndian.PutUint16(app0[10:], xDensity)
	binary.BigEndian.PutUint16(app0[12:], yDensity)
	app0[14] = 0 // thumbnail width
	app0[15] = 0 // thumbnail height
	buf.Write(app0)

	// SOF0 marker
	buf.Write([]byte{0xFF, 0xC0})
	sof := make([]byte, 11)
	binary.BigEndian.PutUint16(sof[0:], 11) // segment length
	sof[2] = 8                              // data precision
	binary.BigEndian.PutUint16(sof[3:], height)
	binary.BigEndian.PutUint16(sof[5:], width)
	sof[7] = 3 // num components
	sof[8] = 1 // component ID
	sof[9] = 0x11
	sof[10] = 0
	buf.Write(sof)

	// SOS marker (terminates marker parsing)
	buf.Write([]byte{0xFF, 0xDA})
	sos := make([]byte, 2)
	binary.BigEndian.PutUint16(sos, 2) // minimal segment length
	buf.Write(sos)

	// EOI marker
	buf.Write([]byte{0xFF, 0xD9})

	return buf.Bytes()
}

// buildMinimalExifJPEG creates a minimal Exif JPEG with an embedded TIFF IFD
// for DPI information.
func buildMinimalExifJPEG(width, height uint16, dpi uint32) []byte {
	var buf bytes.Buffer

	// SOI marker
	buf.Write([]byte{0xFF, 0xD8})

	// APP1 (Exif) marker with embedded TIFF
	tiffData := buildMinimalTIFFBytes(0, 0, dpi, true) // big-endian TIFF
	app1Data := make([]byte, 0, 8+len(tiffData))
	// segment length (2 bytes): includes itself + exif sig (6) + tiff
	segLen := 2 + 6 + len(tiffData)
	app1Data = binary.BigEndian.AppendUint16(app1Data, uint16(segLen))
	app1Data = append(app1Data, "Exif\x00\x00"...)
	app1Data = append(app1Data, tiffData...)

	buf.Write([]byte{0xFF, 0xE1})
	buf.Write(app1Data)

	// SOF0 marker
	buf.Write([]byte{0xFF, 0xC0})
	sof := make([]byte, 11)
	binary.BigEndian.PutUint16(sof[0:], 11)
	sof[2] = 8
	binary.BigEndian.PutUint16(sof[3:], height)
	binary.BigEndian.PutUint16(sof[5:], width)
	sof[7] = 3
	sof[8] = 1
	sof[9] = 0x11
	sof[10] = 0
	buf.Write(sof)

	// SOS marker
	buf.Write([]byte{0xFF, 0xDA})
	sos := make([]byte, 2)
	binary.BigEndian.PutUint16(sos, 2)
	buf.Write(sos)

	// EOI marker
	buf.Write([]byte{0xFF, 0xD9})

	return buf.Bytes()
}

// buildMinimalExifJPEGNonExif creates a JPEG whose APP1 segment is detected
// as Exif by the 4-byte signature table ("Exif" at offset 6) but fails the
// stricter 6-byte "Exif\x00\x00" check in isNonExifApp1, causing DPI to
// default to 72.
func buildMinimalExifJPEGNonExif(width, height uint16) []byte {
	var buf bytes.Buffer

	// SOI
	buf.Write([]byte{0xFF, 0xD8})

	// APP1 whose content starts with "Exif" (matches 4-byte signature at
	// offset 6) but whose 5th and 6th bytes are NOT \x00\x00, so the
	// 6-byte Exif check inside parseApp1Marker fails → DPI defaults to 72.
	buf.Write([]byte{0xFF, 0xE1})
	app1Content := []byte("Exif\x01\x02some non-exif payload padding")
	segLen := make([]byte, 2)
	binary.BigEndian.PutUint16(segLen, uint16(2+len(app1Content)))
	buf.Write(segLen)
	buf.Write(app1Content)

	// SOF0
	buf.Write([]byte{0xFF, 0xC0})
	sof := make([]byte, 11)
	binary.BigEndian.PutUint16(sof[0:], 11)
	sof[2] = 8
	binary.BigEndian.PutUint16(sof[3:], height)
	binary.BigEndian.PutUint16(sof[5:], width)
	sof[7] = 3
	sof[8] = 1
	sof[9] = 0x11
	sof[10] = 0
	buf.Write(sof)

	// SOS
	buf.Write([]byte{0xFF, 0xDA})
	sos := make([]byte, 2)
	binary.BigEndian.PutUint16(sos, 2)
	buf.Write(sos)

	// EOI
	buf.Write([]byte{0xFF, 0xD9})

	return buf.Bytes()
}

// buildMinimalTIFF creates a minimal valid TIFF with given dimensions and DPI.
func buildMinimalTIFF(width, height uint32, dpi uint32, bigEndian bool) []byte {
	return buildMinimalTIFFBytes(width, height, dpi, bigEndian)
}

// buildMinimalTIFFBytes generates a minimal TIFF byte stream.
func buildMinimalTIFFBytes(width, height uint32, dpi uint32, bigEndian bool) []byte {
	var order binary.ByteOrder
	if bigEndian {
		order = binary.BigEndian
	} else {
		order = binary.LittleEndian
	}

	// Determine how many IFD entries we need
	entries := 0
	if width > 0 {
		entries++ // ImageWidth
	}
	if height > 0 {
		entries++ // ImageLength
	}
	if dpi > 0 {
		entries += 3 // XResolution, YResolution, ResolutionUnit
	}

	// Layout:
	// 0-1: byte order ("MM" or "II")
	// 2-3: magic 42
	// 4-7: IFD0 offset
	// 8: IFD0 start
	//   8-9: entry count
	//   10+: entries (12 bytes each)
	//   after entries: 4 bytes next IFD offset (0)
	//   after next IFD: rational values (8 bytes each)

	ifd0Offset := uint32(8)
	entryStart := ifd0Offset + 2
	afterEntries := entryStart + uint32(entries)*12
	nextIFDOffset := afterEntries
	rationalsStart := nextIFDOffset + 4

	totalSize := rationalsStart
	if dpi > 0 {
		totalSize += 16 // two RATIONAL values (8 bytes each)
	}

	data := make([]byte, totalSize)

	// Header
	if bigEndian {
		data[0] = 'M'
		data[1] = 'M'
	} else {
		data[0] = 'I'
		data[1] = 'I'
	}
	order.PutUint16(data[2:], 42)
	order.PutUint32(data[4:], ifd0Offset)

	// IFD entry count
	order.PutUint16(data[ifd0Offset:], uint16(entries))

	entryIdx := 0
	writeEntry := func(tag uint16, fieldType uint16, count uint32, value uint32) {
		off := entryStart + uint32(entryIdx)*12
		order.PutUint16(data[off:], tag)
		order.PutUint16(data[off+2:], fieldType)
		order.PutUint32(data[off+4:], count)
		// TIFF spec: inline SHORT values occupy the first 2 bytes of
		// the 4-byte value field; remaining bytes are zero padding.
		if fieldType == tiffFieldSHORT {
			order.PutUint16(data[off+8:], uint16(value))
		} else {
			order.PutUint32(data[off+8:], value)
		}
		entryIdx++
	}

	if width > 0 {
		writeEntry(tiffTagImageWidth, tiffFieldLONG, 1, width)
	}
	if height > 0 {
		writeEntry(tiffTagImageLength, tiffFieldLONG, 1, height)
	}
	if dpi > 0 {
		// XResolution → RATIONAL at rationalsStart
		writeEntry(tiffTagXResolution, tiffFieldRATIONAL, 1, uint32(rationalsStart))
		// YResolution → RATIONAL at rationalsStart+8
		writeEntry(tiffTagYResolution, tiffFieldRATIONAL, 1, uint32(rationalsStart+8))
		// ResolutionUnit = 2 (inches), stored as SHORT inline
		writeEntry(tiffTagResolutionUnit, tiffFieldSHORT, 1, 2)

		// Write RATIONAL values: numerator / denominator
		order.PutUint32(data[rationalsStart:], dpi)
		order.PutUint32(data[rationalsStart+4:], 1) // denominator = 1
		order.PutUint32(data[rationalsStart+8:], dpi)
		order.PutUint32(data[rationalsStart+12:], 1)
	}

	// Next IFD offset = 0 (no more IFDs)
	order.PutUint32(data[nextIFDOffset:], 0)

	return data
}

// buildMinimalTIFFCm creates a TIFF with resolution unit = 3 (centimeters).
func buildMinimalTIFFCm(width, height uint32, dotsPerCm uint32, bigEndian bool) []byte {
	var order binary.ByteOrder
	if bigEndian {
		order = binary.BigEndian
	} else {
		order = binary.LittleEndian
	}

	entries := 5 // width, height, xres, yres, resunit

	ifd0Offset := uint32(8)
	entryStart := ifd0Offset + 2
	afterEntries := entryStart + uint32(entries)*12
	nextIFDOffset := afterEntries
	rationalsStart := nextIFDOffset + 4
	totalSize := rationalsStart + 16

	data := make([]byte, totalSize)

	if bigEndian {
		data[0], data[1] = 'M', 'M'
	} else {
		data[0], data[1] = 'I', 'I'
	}
	order.PutUint16(data[2:], 42)
	order.PutUint32(data[4:], ifd0Offset)
	order.PutUint16(data[ifd0Offset:], uint16(entries))

	entryIdx := 0
	writeEntry := func(tag, fieldType uint16, count, value uint32) {
		off := entryStart + uint32(entryIdx)*12
		order.PutUint16(data[off:], tag)
		order.PutUint16(data[off+2:], fieldType)
		order.PutUint32(data[off+4:], count)
		if fieldType == tiffFieldSHORT {
			order.PutUint16(data[off+8:], uint16(value))
		} else {
			order.PutUint32(data[off+8:], value)
		}
		entryIdx++
	}

	writeEntry(tiffTagImageWidth, tiffFieldLONG, 1, width)
	writeEntry(tiffTagImageLength, tiffFieldLONG, 1, height)
	writeEntry(tiffTagXResolution, tiffFieldRATIONAL, 1, uint32(rationalsStart))
	writeEntry(tiffTagYResolution, tiffFieldRATIONAL, 1, uint32(rationalsStart+8))
	writeEntry(tiffTagResolutionUnit, tiffFieldSHORT, 1, 3) // centimeters

	order.PutUint32(data[rationalsStart:], dotsPerCm)
	order.PutUint32(data[rationalsStart+4:], 1)
	order.PutUint32(data[rationalsStart+8:], dotsPerCm)
	order.PutUint32(data[rationalsStart+12:], 1)
	order.PutUint32(data[nextIFDOffset:], 0)

	return data
}

// buildMinimalTIFFAspectOnly creates a TIFF with resolution unit = 1 (aspect ratio).
func buildMinimalTIFFAspectOnly(width, height uint32, bigEndian bool) []byte {
	var order binary.ByteOrder
	if bigEndian {
		order = binary.BigEndian
	} else {
		order = binary.LittleEndian
	}

	entries := 5

	ifd0Offset := uint32(8)
	entryStart := ifd0Offset + 2
	afterEntries := entryStart + uint32(entries)*12
	nextIFDOffset := afterEntries
	rationalsStart := nextIFDOffset + 4
	totalSize := rationalsStart + 16

	data := make([]byte, totalSize)

	if bigEndian {
		data[0], data[1] = 'M', 'M'
	} else {
		data[0], data[1] = 'I', 'I'
	}
	order.PutUint16(data[2:], 42)
	order.PutUint32(data[4:], ifd0Offset)
	order.PutUint16(data[ifd0Offset:], uint16(entries))

	entryIdx := 0
	writeEntry := func(tag, fieldType uint16, count, value uint32) {
		off := entryStart + uint32(entryIdx)*12
		order.PutUint16(data[off:], tag)
		order.PutUint16(data[off+2:], fieldType)
		order.PutUint32(data[off+4:], count)
		if fieldType == tiffFieldSHORT {
			order.PutUint16(data[off+8:], uint16(value))
		} else {
			order.PutUint32(data[off+8:], value)
		}
		entryIdx++
	}

	writeEntry(tiffTagImageWidth, tiffFieldLONG, 1, width)
	writeEntry(tiffTagImageLength, tiffFieldLONG, 1, height)
	writeEntry(tiffTagXResolution, tiffFieldRATIONAL, 1, uint32(rationalsStart))
	writeEntry(tiffTagYResolution, tiffFieldRATIONAL, 1, uint32(rationalsStart+8))
	writeEntry(tiffTagResolutionUnit, tiffFieldSHORT, 1, 1) // aspect ratio only

	order.PutUint32(data[rationalsStart:], 300)
	order.PutUint32(data[rationalsStart+4:], 1)
	order.PutUint32(data[rationalsStart+8:], 300)
	order.PutUint32(data[rationalsStart+12:], 1)
	order.PutUint32(data[nextIFDOffset:], 0)

	return data
}

// buildMinimalBMP creates a minimal valid BMP with given dimensions and DPI.
func buildMinimalBMP(width, height uint32, horzPxPerMeter, vertPxPerMeter uint32) []byte {
	// Minimal BMP header is 54 bytes (14 file header + 40 DIB header)
	data := make([]byte, 54)

	// BM signature
	data[0] = 'B'
	data[1] = 'M'

	// File size
	binary.LittleEndian.PutUint32(data[2:], 54)

	// Data offset
	binary.LittleEndian.PutUint32(data[10:], 54)

	// DIB header size
	binary.LittleEndian.PutUint32(data[14:], 40)

	// Width and height
	binary.LittleEndian.PutUint32(data[0x12:], width)
	binary.LittleEndian.PutUint32(data[0x16:], height)

	// Planes
	binary.LittleEndian.PutUint16(data[0x1A:], 1)

	// Bits per pixel
	binary.LittleEndian.PutUint16(data[0x1C:], 24)

	// Horizontal and vertical pixels per meter
	binary.LittleEndian.PutUint32(data[0x26:], horzPxPerMeter)
	binary.LittleEndian.PutUint32(data[0x2A:], vertPxPerMeter)

	return data
}

// buildMinimalGIF creates a minimal valid GIF with given dimensions.
func buildMinimalGIF(width, height uint16) []byte {
	var buf bytes.Buffer

	// Header
	buf.WriteString("GIF89a")

	// Logical Screen Descriptor
	dim := make([]byte, 4)
	binary.LittleEndian.PutUint16(dim[0:], width)
	binary.LittleEndian.PutUint16(dim[2:], height)
	buf.Write(dim)

	// Packed byte: no global color table
	buf.WriteByte(0x00)
	// Background color index
	buf.WriteByte(0x00)
	// Pixel aspect ratio
	buf.WriteByte(0x00)

	// Trailer
	buf.WriteByte(0x3B)

	return buf.Bytes()
}

// pxPerMeterFromDPI converts DPI to pixels per meter.
func pxPerMeterFromDPI(dpi int) uint32 {
	return uint32(math.Round(float64(dpi) / 0.0254))
}
