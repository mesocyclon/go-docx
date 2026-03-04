package image

import (
	"bytes"
	"errors"
	"math"
	"testing"
)

func TestFromBlob_PNG(t *testing.T) {
	blob := buildMinimalPNG(100, 200, pxPerMeterFromDPI(96), 1)
	img, err := FromBlob(blob, "test.png")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 100 {
		t.Errorf("PxWidth = %d, want 100", img.PxWidth())
	}
	if img.PxHeight() != 200 {
		t.Errorf("PxHeight = %d, want 200", img.PxHeight())
	}
	if img.ContentType() != MimePNG {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimePNG)
	}
	if img.Ext() != "png" {
		t.Errorf("Ext = %q, want %q", img.Ext(), "png")
	}
	if img.Filename() != "test.png" {
		t.Errorf("Filename = %q, want %q", img.Filename(), "test.png")
	}
	if img.HorzDpi() != 96 {
		t.Errorf("HorzDpi = %d, want 96", img.HorzDpi())
	}
}

func TestFromBlob_JFIF(t *testing.T) {
	blob := buildMinimalJFIF(320, 240, 1, 150, 150)
	img, err := FromBlob(blob, "photo.jpg")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 320 {
		t.Errorf("PxWidth = %d, want 320", img.PxWidth())
	}
	if img.PxHeight() != 240 {
		t.Errorf("PxHeight = %d, want 240", img.PxHeight())
	}
	if img.ContentType() != MimeJPEG {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeJPEG)
	}
	if img.Ext() != "jpg" {
		t.Errorf("Ext = %q, want %q", img.Ext(), "jpg")
	}
	if img.HorzDpi() != 150 {
		t.Errorf("HorzDpi = %d, want 150", img.HorzDpi())
	}
	if img.VertDpi() != 150 {
		t.Errorf("VertDpi = %d, want 150", img.VertDpi())
	}
}

func TestFromBlob_GIF(t *testing.T) {
	blob := buildMinimalGIF(10, 20)
	img, err := FromBlob(blob, "anim.gif")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 10 {
		t.Errorf("PxWidth = %d, want 10", img.PxWidth())
	}
	if img.PxHeight() != 20 {
		t.Errorf("PxHeight = %d, want 20", img.PxHeight())
	}
	if img.ContentType() != MimeGIF {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeGIF)
	}
	if img.HorzDpi() != 72 {
		t.Errorf("HorzDpi = %d, want 72", img.HorzDpi())
	}
	if img.VertDpi() != 72 {
		t.Errorf("VertDpi = %d, want 72", img.VertDpi())
	}
}

func TestFromBlob_BMP(t *testing.T) {
	blob := buildMinimalBMP(4, 4, pxPerMeterFromDPI(96), pxPerMeterFromDPI(96))
	img, err := FromBlob(blob, "test.bmp")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 4 {
		t.Errorf("PxWidth = %d, want 4", img.PxWidth())
	}
	if img.PxHeight() != 4 {
		t.Errorf("PxHeight = %d, want 4", img.PxHeight())
	}
	if img.ContentType() != MimeBMP {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeBMP)
	}
	if img.HorzDpi() != 96 {
		t.Errorf("HorzDpi = %d, want 96", img.HorzDpi())
	}
}

func TestFromBlob_TIFF_BigEndian(t *testing.T) {
	blob := buildMinimalTIFF(50, 60, 300, true)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 50 {
		t.Errorf("PxWidth = %d, want 50", img.PxWidth())
	}
	if img.PxHeight() != 60 {
		t.Errorf("PxHeight = %d, want 60", img.PxHeight())
	}
	if img.ContentType() != MimeTIFF {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeTIFF)
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
}

func TestFromBlob_TIFF_LittleEndian(t *testing.T) {
	blob := buildMinimalTIFF(80, 90, 150, false)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 80 {
		t.Errorf("PxWidth = %d, want 80", img.PxWidth())
	}
	if img.HorzDpi() != 150 {
		t.Errorf("HorzDpi = %d, want 150", img.HorzDpi())
	}
}

func TestFromBlob_UnrecognizedFormat(t *testing.T) {
	blob := []byte{0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07,
		0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F,
		0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17,
		0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E, 0x1F}
	_, err := FromBlob(blob, "unknown.xyz")
	if err == nil {
		t.Fatal("expected error for unrecognized format")
	}
	if !errors.Is(err, ErrUnrecognizedImage) {
		t.Errorf("expected ErrUnrecognizedImage, got %v", err)
	}
}

func TestFromReadSeeker(t *testing.T) {
	blob := buildMinimalGIF(15, 25)
	r := bytes.NewReader(blob)
	img, err := FromReadSeeker(r, "stream.gif")
	if err != nil {
		t.Fatalf("FromReadSeeker: %v", err)
	}
	if img.PxWidth() != 15 {
		t.Errorf("PxWidth = %d, want 15", img.PxWidth())
	}
}

func TestImage_DefaultFilename(t *testing.T) {
	blob := buildMinimalPNG(1, 1, 0, 0)
	img, err := FromBlob(blob, "")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.Filename() != "image.png" {
		t.Errorf("Filename = %q, want %q", img.Filename(), "image.png")
	}
}

func TestImage_Hash_Stable(t *testing.T) {
	blob := buildMinimalGIF(10, 10)
	img1, _ := FromBlob(blob, "a.gif")
	img2, _ := FromBlob(blob, "b.gif")

	if img1.Hash() == "" {
		t.Error("Hash should not be empty")
	}
	if img1.Hash() != img2.Hash() {
		t.Error("Hash should be stable for identical blobs")
	}

	// Verify calling Hash again returns cached value
	ha := img1.Hash()
	hb := img1.Hash()
	if ha != hb {
		t.Error("Hash should be cached")
	}
}

func TestImage_Hash_DifferentBlobs(t *testing.T) {
	blob1 := buildMinimalGIF(10, 10)
	blob2 := buildMinimalGIF(20, 20)
	img1, _ := FromBlob(blob1, "a.gif")
	img2, _ := FromBlob(blob2, "b.gif")

	if img1.Hash() == img2.Hash() {
		t.Error("Hash should differ for different blobs")
	}
}

func TestImage_Width_Height_EMU(t *testing.T) {
	// 72 DPI, 72x72 pixels → 1 inch = 914400 EMU
	blob := buildMinimalPNG(72, 144, 0, 0) // no pHYs → defaults to 72 DPI
	img, err := FromBlob(blob, "test.png")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}

	expectedW := int64(float64(72) / 72.0 * 914400)  // 914400
	expectedH := int64(float64(144) / 72.0 * 914400) // 1828800

	if img.Width() != expectedW {
		t.Errorf("Width = %d, want %d", img.Width(), expectedW)
	}
	if img.Height() != expectedH {
		t.Errorf("Height = %d, want %d", img.Height(), expectedH)
	}
}

func TestImage_ScaledDimensions_BothNil(t *testing.T) {
	blob := buildMinimalPNG(72, 144, 0, 0)
	img, _ := FromBlob(blob, "test.png")

	cx, cy := img.ScaledDimensions(nil, nil)
	if cx != img.Width() || cy != img.Height() {
		t.Errorf("ScaledDimensions(nil, nil) = (%d, %d), want (%d, %d)",
			cx, cy, img.Width(), img.Height())
	}
}

func TestImage_ScaledDimensions_WidthOnly(t *testing.T) {
	blob := buildMinimalPNG(100, 200, 0, 0) // 72 DPI
	img, _ := FromBlob(blob, "test.png")

	w := int64(457200) // half native width
	cx, cy := img.ScaledDimensions(&w, nil)

	if cx != w {
		t.Errorf("cx = %d, want %d", cx, w)
	}

	// Height should be scaled proportionally
	scalingFactor := float64(w) / float64(img.Width())
	expectedH := int64(math.Round(float64(img.Height()) * scalingFactor))
	if cy != expectedH {
		t.Errorf("cy = %d, want %d", cy, expectedH)
	}
}

func TestImage_ScaledDimensions_HeightOnly(t *testing.T) {
	blob := buildMinimalPNG(100, 200, 0, 0)
	img, _ := FromBlob(blob, "test.png")

	h := int64(914400) // half native height
	cx, cy := img.ScaledDimensions(nil, &h)

	if cy != h {
		t.Errorf("cy = %d, want %d", cy, h)
	}

	scalingFactor := float64(h) / float64(img.Height())
	expectedW := int64(math.Round(float64(img.Width()) * scalingFactor))
	if cx != expectedW {
		t.Errorf("cx = %d, want %d", cx, expectedW)
	}
}

func TestImage_ScaledDimensions_BothSpecified(t *testing.T) {
	blob := buildMinimalPNG(100, 200, 0, 0)
	img, _ := FromBlob(blob, "test.png")

	w := int64(500000)
	h := int64(1000000)
	cx, cy := img.ScaledDimensions(&w, &h)

	if cx != w || cy != h {
		t.Errorf("ScaledDimensions = (%d, %d), want (%d, %d)", cx, cy, w, h)
	}
}

func TestImage_Blob(t *testing.T) {
	blob := buildMinimalGIF(5, 5)
	img, _ := FromBlob(blob, "test.gif")

	if !bytes.Equal(img.Blob(), blob) {
		t.Error("Blob() should return the original blob")
	}
}

func TestImage_Ext_PreservesCase(t *testing.T) {
	blob := buildMinimalPNG(1, 1, 0, 0)

	// Python: os.path.splitext("test.PNG")[1][1:] = "PNG" (no lowercasing)
	img, _ := FromBlob(blob, "test.PNG")
	if img.Ext() != "PNG" {
		t.Errorf("Ext() = %q, want %q (should preserve case)", img.Ext(), "PNG")
	}

	img2, _ := FromBlob(blob, "test.jpg")
	if img2.Ext() != "jpg" {
		t.Errorf("Ext() = %q, want %q", img2.Ext(), "jpg")
	}
}
