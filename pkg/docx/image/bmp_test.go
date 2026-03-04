package image

import (
	"testing"
)

func TestBMP_Dimensions(t *testing.T) {
	blob := buildMinimalBMP(320, 240, pxPerMeterFromDPI(96), pxPerMeterFromDPI(96))
	img, err := FromBlob(blob, "test.bmp")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 320 {
		t.Errorf("PxWidth = %d, want 320", img.PxWidth())
	}
	if img.PxHeight() != 240 {
		t.Errorf("PxHeight = %d, want 240", img.PxHeight())
	}
}

func TestBMP_DPI_96(t *testing.T) {
	blob := buildMinimalBMP(4, 4, pxPerMeterFromDPI(96), pxPerMeterFromDPI(96))
	img, err := FromBlob(blob, "test.bmp")
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

func TestBMP_ZeroPxPerMeter_Default96(t *testing.T) {
	blob := buildMinimalBMP(4, 4, 0, 0)
	img, err := FromBlob(blob, "test.bmp")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 96 {
		t.Errorf("HorzDpi = %d, want 96 (default)", img.HorzDpi())
	}
	if img.VertDpi() != 96 {
		t.Errorf("VertDpi = %d, want 96 (default)", img.VertDpi())
	}
}

func TestBMP_ContentType(t *testing.T) {
	blob := buildMinimalBMP(1, 1, 0, 0)
	img, err := FromBlob(blob, "test.bmp")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.ContentType() != MimeBMP {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeBMP)
	}
}

func TestBMP_HighDPI(t *testing.T) {
	blob := buildMinimalBMP(100, 100, pxPerMeterFromDPI(300), pxPerMeterFromDPI(300))
	img, err := FromBlob(blob, "test.bmp")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.HorzDpi() != 300 {
		t.Errorf("HorzDpi = %d, want 300", img.HorzDpi())
	}
}
