package image

import (
	"testing"
)

func TestTIFF_BigEndian(t *testing.T) {
	blob := buildMinimalTIFF(100, 200, 300, true)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 100 {
		t.Errorf("PxWidth = %d, want 100", img.PxWidth())
	}
	if img.PxHeight() != 200 {
		t.Errorf("PxHeight = %d, want 200", img.PxHeight())
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
	if img.VertDpi() != 300 {
		t.Errorf("VertDpi = %d, want 300", img.VertDpi())
	}
}

func TestTIFF_LittleEndian(t *testing.T) {
	blob := buildMinimalTIFF(640, 480, 150, false)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 640 {
		t.Errorf("PxWidth = %d, want 640", img.PxWidth())
	}
	if img.PxHeight() != 480 {
		t.Errorf("PxHeight = %d, want 480", img.PxHeight())
	}
	if img.HorzDpi() != 150 {
		t.Errorf("HorzDpi = %d, want 150", img.HorzDpi())
	}
}

func TestTIFF_ResolutionUnitCm(t *testing.T) {
	// 118 dots/cm * 2.54 = 299.72 → rounds to 300
	blob := buildMinimalTIFFCm(100, 100, 118, true)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
}

func TestTIFF_ResolutionUnitAspectOnly(t *testing.T) {
	// Resolution unit = 1 (aspect ratio only) → DPI defaults to 72
	blob := buildMinimalTIFFAspectOnly(100, 100, true)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 72 {
		t.Errorf("HorzDpi = %d, want 72 (aspect ratio only)", img.HorzDpi())
	}
	if img.VertDpi() != 72 {
		t.Errorf("VertDpi = %d, want 72 (aspect ratio only)", img.VertDpi())
	}
}

func TestTIFF_MissingResolutionTags(t *testing.T) {
	// TIFF with only width/height, no resolution tags → default 72
	blob := buildMinimalTIFF(50, 60, 0, true)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 72 {
		t.Errorf("HorzDpi = %d, want 72 (missing tags)", img.HorzDpi())
	}
}

func TestTIFF_ContentType(t *testing.T) {
	blob := buildMinimalTIFF(1, 1, 72, true)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.ContentType() != MimeTIFF {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeTIFF)
	}
}

func TestTIFF_RATIONAL_Resolution(t *testing.T) {
	// Verify RATIONAL 300/1 = 300 DPI
	blob := buildMinimalTIFF(100, 100, 300, false)
	img, err := FromBlob(blob, "test.tiff")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
}
