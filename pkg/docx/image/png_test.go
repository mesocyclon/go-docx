package image

import (
	"errors"
	"math"
	"testing"
)

func TestPNG_WithPHYs_Meter(t *testing.T) {
	// 3780 px/m ≈ 96 DPI (3780 * 0.0254 = 96.012 → rounds to 96)
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
	if img.HorzDpi() != 96 {
		t.Errorf("HorzDpi = %d, want 96", img.HorzDpi())
	}
	if img.VertDpi() != 96 {
		t.Errorf("VertDpi = %d, want 96", img.VertDpi())
	}
}

func TestPNG_WithoutPHYs_Default72(t *testing.T) {
	blob := buildMinimalPNG(50, 50, 0, 0) // no pHYs chunk
	img, err := FromBlob(blob, "test.png")
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

func TestPNG_PHYs_UnitNotMeter_Default72(t *testing.T) {
	// units_specifier = 0 (unknown unit) → DPI defaults to 72
	blob := buildMinimalPNG(50, 50, 3780, 0)
	img, err := FromBlob(blob, "test.png")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 72 {
		t.Errorf("HorzDpi = %d, want 72", img.HorzDpi())
	}
}

func TestPNG_IHDR_Dimensions(t *testing.T) {
	blob := buildMinimalPNG(1920, 1080, 0, 0)
	img, err := FromBlob(blob, "test.png")
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

func TestPNG_MissingIHDR(t *testing.T) {
	// Manually craft a PNG with no IHDR chunk (just IEND)
	blob := []byte{
		0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, // signature
		0x00, 0x00, 0x00, 0x00, // chunk data len = 0
		'I', 'E', 'N', 'D',    // chunk type
		0xAE, 0x42, 0x60, 0x82, // CRC
	}
	_, err := FromBlob(blob, "bad.png")
	if err == nil {
		t.Fatal("expected error for missing IHDR")
	}
	if !errors.Is(err, ErrInvalidImageStream) {
		t.Errorf("expected ErrInvalidImageStream, got %v", err)
	}
}

func TestPNG_ContentType(t *testing.T) {
	blob := buildMinimalPNG(1, 1, 0, 0)
	img, err := FromBlob(blob, "test.png")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.ContentType() != MimePNG {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimePNG)
	}
}

func TestPNG_HighDPI(t *testing.T) {
	// 300 DPI
	ppm := uint32(math.Round(300.0 / 0.0254))
	blob := buildMinimalPNG(100, 100, ppm, 1)
	img, err := FromBlob(blob, "test.png")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
}

func TestPNG_SeparateHorzVertDPI(t *testing.T) {
	horzPPM := uint32(math.Round(96.0 / 0.0254))
	vertPPM := uint32(math.Round(72.0 / 0.0254))
	blob := buildMinimalPNGWithSeparateDPI(100, 200, horzPPM, vertPPM, 1)
	img, err := FromBlob(blob, "test.png")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 96 {
		t.Errorf("HorzDpi = %d, want 96", img.HorzDpi())
	}
	if img.VertDpi() != 72 {
		t.Errorf("VertDpi = %d, want 72", img.VertDpi())
	}
}
