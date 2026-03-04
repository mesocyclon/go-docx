package image

import (
	"bytes"
	"encoding/binary"
	"testing"
)

func TestJFIF_DensityUnits1_DPI(t *testing.T) {
	blob := buildMinimalJFIF(640, 480, 1, 96, 96) // units=1: DPI
	img, err := FromBlob(blob, "test.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 96 {
		t.Errorf("HorzDpi = %d, want 96", img.HorzDpi())
	}
	if img.VertDpi() != 96 {
		t.Errorf("VertDpi = %d, want 96", img.VertDpi())
	}
}

func TestJFIF_DensityUnits2_DPCM(t *testing.T) {
	// units=2: dots per cm. 38 dpcm ≈ 97 dpi (38 * 2.54 = 96.52 → round to 97)
	blob := buildMinimalJFIF(640, 480, 2, 38, 38)
	img, err := FromBlob(blob, "test.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	// 38 * 2.54 = 96.52 → rounds to 97
	if img.HorzDpi() != 97 {
		t.Errorf("HorzDpi = %d, want 97", img.HorzDpi())
	}
}

func TestJFIF_DensityUnits0_Default72(t *testing.T) {
	// units=0: no unit → default to 72
	blob := buildMinimalJFIF(640, 480, 0, 1, 1)
	img, err := FromBlob(blob, "test.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 72 {
		t.Errorf("HorzDpi = %d, want 72", img.HorzDpi())
	}
	if img.VertDpi() != 72 {
		t.Errorf("VertDpi = %d, want 72", img.VertDpi())
	}
}

func TestJFIF_Dimensions(t *testing.T) {
	blob := buildMinimalJFIF(1920, 1080, 1, 72, 72)
	img, err := FromBlob(blob, "test.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 1920 {
		t.Errorf("PxWidth = %d, want 1920", img.PxWidth())
	}
	if img.PxHeight() != 1080 {
		t.Errorf("PxHeight = %d, want 1080", img.PxHeight())
	}
}

func TestExif_ValidExifSignature(t *testing.T) {
	blob := buildMinimalExifJPEG(800, 600, 300)
	img, err := FromBlob(blob, "exif.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 800 {
		t.Errorf("PxWidth = %d, want 800", img.PxWidth())
	}
	if img.PxHeight() != 600 {
		t.Errorf("PxHeight = %d, want 600", img.PxHeight())
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
	if img.VertDpi() != 300 {
		t.Errorf("VertDpi = %d, want 300", img.VertDpi())
	}
	if img.ContentType() != MimeJPEG {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeJPEG)
	}
}

func TestExif_NonExifAPP1_Default72(t *testing.T) {
	blob := buildMinimalExifJPEGNonExif(640, 480)
	img, err := FromBlob(blob, "xmp.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 640 {
		t.Errorf("PxWidth = %d, want 640", img.PxWidth())
	}
	// Non-Exif APP1 → default 72 DPI
	if img.HorzDpi() != 72 {
		t.Errorf("HorzDpi = %d, want 72", img.HorzDpi())
	}
}

func TestJFIF_ContentType(t *testing.T) {
	blob := buildMinimalJFIF(1, 1, 1, 72, 72)
	img, err := FromBlob(blob, "test.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.ContentType() != MimeJPEG {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeJPEG)
	}
}

func TestJPEG_Truncated(t *testing.T) {
	// Just SOI marker, truncated
	blob := []byte{0xFF, 0xD8, 0xFF, 0xE0}
	_, err := FromBlob(blob, "truncated.jpg")
	if err == nil {
		t.Fatal("expected error for truncated JPEG")
	}
}

func TestJPEG_MarkerScanFF00(t *testing.T) {
	// Verify that FF 00 sequences are skipped during marker scanning.
	// FF 00 is placed BETWEEN APP0 and SOF0 (not before APP0) so that
	// the "JFIF" signature remains at file offset 6 for format detection.
	var buf bytes.Buffer

	// SOI
	buf.Write([]byte{0xFF, 0xD8})

	// APP0 (JFIF) — "JFIF" will be at offset 6 in the file
	buf.Write([]byte{0xFF, 0xE0})
	app0 := make([]byte, 16)
	binary.BigEndian.PutUint16(app0[0:], 16) // segment length
	copy(app0[2:7], "JFIF\x00")
	app0[7] = 1 // major version
	app0[8] = 1 // minor version
	app0[9] = 1 // density units = DPI
	binary.BigEndian.PutUint16(app0[10:], 72)
	binary.BigEndian.PutUint16(app0[12:], 72)
	buf.Write(app0)

	// FF 00 stuffed byte — NOT a marker, scanner must skip it
	buf.Write([]byte{0xFF, 0x00})

	// SOF0
	buf.Write([]byte{0xFF, 0xC0})
	sof := make([]byte, 11)
	binary.BigEndian.PutUint16(sof[0:], 11) // segment length
	sof[2] = 8                              // data precision
	binary.BigEndian.PutUint16(sof[3:], 10) // height = 10
	binary.BigEndian.PutUint16(sof[5:], 20) // width = 20
	sof[7] = 3                              // num components
	buf.Write(sof)

	// SOS
	buf.Write([]byte{0xFF, 0xDA})
	buf.Write([]byte{0x00, 0x02})

	// EOI
	buf.Write([]byte{0xFF, 0xD9})

	img, err := FromBlob(buf.Bytes(), "ff00.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 20 || img.PxHeight() != 10 {
		t.Errorf("dimensions = %dx%d, want 20x10", img.PxWidth(), img.PxHeight())
	}
}
