package parts

import (
	"bytes"
	"crypto/sha256"
	"fmt"
	goimage "image"
	"math"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/image"
	"github.com/vortex/go-docx/pkg/docx/opc"
)

func TestImagePartHash_Stable(t *testing.T) {
	blob := []byte("test image data")
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, blob, nil)

	h1, err := ip.Hash()
	if err != nil {
		t.Fatal(err)
	}
	h2, err := ip.Hash()
	if err != nil {
		t.Fatal(err)
	}
	if h1 != h2 {
		t.Errorf("Hash not stable: %q != %q", h1, h2)
	}

	// Verify against direct computation
	expected := fmt.Sprintf("%x", sha256.Sum256(blob))
	if h1 != expected {
		t.Errorf("Hash = %q, want %q", h1, expected)
	}
}

func TestImagePartHash_SameBlob(t *testing.T) {
	blob := []byte("identical data")
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, blob, nil)
	ip2 := NewImagePart("/word/media/image2.png", opc.CTPng, blob, nil)

	h1, err := ip1.Hash()
	if err != nil {
		t.Fatal(err)
	}
	h2, err := ip2.Hash()
	if err != nil {
		t.Fatal(err)
	}
	if h1 != h2 {
		t.Error("Same blob should produce same Hash")
	}
}

func TestImagePartFilename(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, nil, nil)
	got := ip.Filename()
	if got != "image.png" {
		t.Errorf("Filename = %q, want %q", got, "image.png")
	}

	ip.SetFilename("photo.jpg")
	got = ip.Filename()
	if got != "photo.jpg" {
		t.Errorf("Filename after set = %q, want %q", got, "photo.jpg")
	}
}

func TestImagePartDefaultCx_Truncates(t *testing.T) {
	// Python: Inches(px_width / horz_dpi) → int(float * 914400) → TRUNCATES
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	cx, err := ip.DefaultCx()
	if err != nil {
		t.Fatal(err)
	}
	// int(100.0 / 96.0 * 914400) = int(952500.0) = 952500
	px, dpi := float64(100), float64(96)
	expected := int64(px / dpi * 914400)
	if cx != expected {
		t.Errorf("DefaultCx = %d, want %d (truncation)", cx, expected)
	}
}

func TestImagePartDefaultCy_Rounds_UsesVertDpi(t *testing.T) {
	// DefaultCy uses vert_dpi (fixing Python's bug where it used horz_dpi)
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 300, "",
	)
	cy, err := ip.DefaultCy()
	if err != nil {
		t.Fatal(err)
	}
	// Uses vert_dpi (300), not horz_dpi (96)
	expected := int64(math.Round(914400 * float64(200) / float64(300)))
	if cy != expected {
		t.Errorf("DefaultCy = %d, want %d (should use vert_dpi)", cy, expected)
	}
}

func TestImagePartDefaultCx_NoDPI(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, nil, nil)
	_, err := ip.DefaultCx()
	if err == nil {
		t.Error("expected error for image with no DPI")
	}
}

func TestNativeWidth_MatchesPythonImageWidth(t *testing.T) {
	// Python Image.width = Inches(px_width / horz_dpi) → int() truncation
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	w, err := ip.NativeWidth()
	if err != nil {
		t.Fatal(err)
	}
	px, dpi := float64(100), float64(96)
	expected := int64(px / dpi * 914400) // truncation
	if w != expected {
		t.Errorf("NativeWidth = %d, want %d", w, expected)
	}
}

func TestNativeHeight_UsesVertDpi(t *testing.T) {
	// Python Image.height = Inches(px_height / vert_dpi) → int() truncation
	// DefaultCy also uses vert_dpi now but with round() instead of truncation.
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 300, "",
	)
	h, err := ip.NativeHeight()
	if err != nil {
		t.Fatal(err)
	}
	// int(200.0 / 300.0 * 914400) = int(609600.0) = 609600
	px, dpi := float64(200), float64(300)
	expected := int64(px / dpi * 914400) // truncation, vert_dpi
	if h != expected {
		t.Errorf("NativeHeight = %d, want %d (should use vert_dpi)", h, expected)
	}

	// DefaultCy also uses vert_dpi but rounds instead of truncating
	cy, _ := ip.DefaultCy()
	cyExpected := int64(math.Round(914400 * float64(200) / float64(300))) // vert_dpi, rounded
	if cy != cyExpected {
		t.Errorf("DefaultCy = %d, want %d", cy, cyExpected)
	}
}

// --------------------------------------------------------------------------
// ScaledDimensions — uses NativeWidth/NativeHeight (Python Image.scaled_dimensions)
// --------------------------------------------------------------------------

func TestScaledDimensions_BothNil(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	cx, cy, err := ip.ScaledDimensions(nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	nativeW, _ := ip.NativeWidth()
	nativeH, _ := ip.NativeHeight()
	if cx != nativeW || cy != nativeH {
		t.Errorf("ScaledDimensions(nil,nil) = (%d,%d), want (%d,%d)", cx, cy, nativeW, nativeH)
	}
}

func TestScaledDimensions_BothNil_DifferentDpi(t *testing.T) {
	// When horz_dpi != vert_dpi, ScaledDimensions uses NativeWidth (horz) and
	// NativeHeight (vert), which matches DefaultCx/DefaultCy behavior.
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 300, "",
	)
	cx, cy, err := ip.ScaledDimensions(nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	nativeW, _ := ip.NativeWidth()
	nativeH, _ := ip.NativeHeight()
	if cx != nativeW || cy != nativeH {
		t.Errorf("ScaledDimensions(nil,nil) = (%d,%d), want (%d,%d)", cx, cy, nativeW, nativeH)
	}
}

func TestScaledDimensions_WidthOnly(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	w := int64(457200)
	cx, cy, err := ip.ScaledDimensions(&w, nil)
	if err != nil {
		t.Fatal(err)
	}
	if cx != w {
		t.Errorf("cx = %d, want %d", cx, w)
	}
	nativeW, _ := ip.NativeWidth()
	nativeH, _ := ip.NativeHeight()
	expectedH := int64(math.Round(float64(nativeH) * float64(w) / float64(nativeW)))
	if cy != expectedH {
		t.Errorf("cy = %d, want %d", cy, expectedH)
	}
}

func TestScaledDimensions_HeightOnly(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	h := int64(914400)
	cx, cy, err := ip.ScaledDimensions(nil, &h)
	if err != nil {
		t.Fatal(err)
	}
	if cy != h {
		t.Errorf("cy = %d, want %d", cy, h)
	}
	nativeW, _ := ip.NativeWidth()
	nativeH, _ := ip.NativeHeight()
	expectedW := int64(math.Round(float64(nativeW) * float64(h) / float64(nativeH)))
	if cx != expectedW {
		t.Errorf("cx = %d, want %d", cx, expectedW)
	}
}

func TestScaledDimensions_BothSpecified(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	w, h := int64(111), int64(222)
	cx, cy, err := ip.ScaledDimensions(&w, &h)
	if err != nil {
		t.Fatal(err)
	}
	if cx != w || cy != h {
		t.Errorf("ScaledDimensions(&111,&222) = (%d,%d), want (111,222)", cx, cy)
	}
}

// =========================================================================
// ImagePart — SetImageMeta, ensureMeta, NewImagePartFromImage
// =========================================================================

func TestImagePart_SetImageMeta(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, nil, nil)
	ip.SetImageMeta(800, 600, 150, 150)

	w, err := ip.PxWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w != 800 {
		t.Errorf("PxWidth = %d, want 800", w)
	}

	h, err := ip.PxHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h != 600 {
		t.Errorf("PxHeight = %d, want 600", h)
	}

	hd, err := ip.HorzDpi()
	if err != nil {
		t.Fatal(err)
	}
	if hd != 150 {
		t.Errorf("HorzDpi = %d, want 150", hd)
	}

	vd, err := ip.VertDpi()
	if err != nil {
		t.Fatal(err)
	}
	if vd != 150 {
		t.Errorf("VertDpi = %d, want 150", vd)
	}
}

func TestImagePart_LazyMetaFromBlob(t *testing.T) {
	// Create ImagePart with a real PNG blob (no metadata set upfront).
	// ensureMeta should parse the blob automatically.
	pngBlob := minimumPNG()
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, pngBlob, nil)

	w, err := ip.PxWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w != 1 {
		t.Errorf("PxWidth = %d, want 1 (1x1 PNG)", w)
	}

	h, err := ip.PxHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h != 1 {
		t.Errorf("PxHeight = %d, want 1", h)
	}
}

func TestImagePart_EnsureMeta_EmptyBlob(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, []byte{}, nil)
	_, err := ip.PxWidth()
	if err == nil {
		t.Error("PxWidth should error on empty blob")
	}
}

func TestImagePart_NewImagePartFromImage(t *testing.T) {
	pngBlob := minimumPNG()
	// Verify blob is valid
	_, _, err := goimage.Decode(bytes.NewReader(pngBlob))
	if err != nil {
		t.Fatal(err)
	}

	imgLib, err := image.FromBlob(pngBlob, "photo.png")
	if err != nil {
		t.Fatal(err)
	}

	ip := NewImagePartFromImage(imgLib, pngBlob)
	if ip == nil {
		t.Fatal("NewImagePartFromImage returned nil")
	}
	fn := ip.Filename()
	if fn == "" {
		t.Error("Filename should not be empty")
	}

	// Hash should be carried from Image
	h, err := ip.Hash()
	if err != nil {
		t.Fatal(err)
	}
	if h == "" {
		t.Error("Hash should not be empty")
	}
}

