package image

import (
	"testing"
)

func TestGIF_Dimensions(t *testing.T) {
	blob := buildMinimalGIF(320, 240)
	img, err := FromBlob(blob, "test.gif")
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

func TestGIF_DPI_Always72(t *testing.T) {
	blob := buildMinimalGIF(10, 10)
	img, err := FromBlob(blob, "test.gif")
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

func TestGIF_ContentType(t *testing.T) {
	blob := buildMinimalGIF(1, 1)
	img, err := FromBlob(blob, "test.gif")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.ContentType() != MimeGIF {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), MimeGIF)
	}
}

func TestGIF_SmallDimensions(t *testing.T) {
	blob := buildMinimalGIF(1, 1)
	img, err := FromBlob(blob, "tiny.gif")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 1 || img.PxHeight() != 1 {
		t.Errorf("dimensions = %dx%d, want 1x1", img.PxWidth(), img.PxHeight())
	}
}

func TestGIF_LargeDimensions(t *testing.T) {
	blob := buildMinimalGIF(65535, 65535)
	img, err := FromBlob(blob, "large.gif")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 65535 || img.PxHeight() != 65535 {
		t.Errorf("dimensions = %dx%d, want 65535x65535", img.PxWidth(), img.PxHeight())
	}
}

func TestGIF87a(t *testing.T) {
	// Build a GIF87a variant
	blob := buildMinimalGIF(100, 50)
	copy(blob[0:6], "GIF87a")
	img, err := FromBlob(blob, "old.gif")
	if err != nil {
		t.Fatalf("FromBlob: %v", err)
	}
	if img.PxWidth() != 100 {
		t.Errorf("PxWidth = %d, want 100", img.PxWidth())
	}
}
