package docx

import (
	"bytes"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// InlineShapes round-trip tests (MR-13)
// -----------------------------------------------------------------------

// minimalPNG returns a minimal valid 1x1 pixel PNG image (72 DPI).
// This is the simplest PNG that will pass image header parsing.
func minimalPNG() []byte {
	return []byte{
		// PNG signature
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
		// IHDR chunk: width=1, height=1, bitDepth=8, colorType=2 (RGB)
		0x00, 0x00, 0x00, 0x0D, // length = 13
		0x49, 0x48, 0x44, 0x52, // "IHDR"
		0x00, 0x00, 0x00, 0x01, // width=1
		0x00, 0x00, 0x00, 0x01, // height=1
		0x08,             // bitDepth=8
		0x02,             // colorType=2 (RGB)
		0x00, 0x00, 0x00, // compression, filter, interlace
		0x90, 0x77, 0x53, 0xDE, // CRC
		// IDAT chunk: minimal compressed pixel data
		0x00, 0x00, 0x00, 0x0C, // length = 12
		0x49, 0x44, 0x41, 0x54, // "IDAT"
		0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00,
		0x00, 0x02, 0x00, 0x01, // compressed data
		0xE2, 0x21, 0xBC, 0x33, // CRC
		// IEND chunk
		0x00, 0x00, 0x00, 0x00, // length = 0
		0x49, 0x45, 0x4E, 0x44, // "IEND"
		0xAE, 0x42, 0x60, 0x82, // CRC
	}
}

func TestInlineShapes_WithPicture_LenOne(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	_, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	shapes, err := doc.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	if shapes.Len() != 1 {
		t.Errorf("InlineShapes.Len() = %d, want 1", shapes.Len())
	}
}

func TestInlineShapes_WithPicture_Width(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	shape, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	w, err := shape.Width()
	if err != nil {
		t.Fatalf("Width(): %v", err)
	}
	if w <= 0 {
		t.Errorf("Width() = %d, expected > 0", w)
	}
}

func TestInlineShapes_WithPicture_Height(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	shape, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	h, err := shape.Height()
	if err != nil {
		t.Fatalf("Height(): %v", err)
	}
	if h <= 0 {
		t.Errorf("Height() = %d, expected > 0", h)
	}
}

func TestInlineShapes_SetWidth_RoundTrip(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	shape, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	newWidth := Inches(2)
	if err := shape.SetWidth(newWidth); err != nil {
		t.Fatalf("SetWidth: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	shapes, err := doc2.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	if shapes.Len() != 1 {
		t.Fatalf("expected 1 inline shape, got %d", shapes.Len())
	}
	shape2, err := shapes.Get(0)
	if err != nil {
		t.Fatalf("Get(0): %v", err)
	}
	w2, err := shape2.Width()
	if err != nil {
		t.Fatalf("Width(): %v", err)
	}
	if w2 != newWidth {
		t.Errorf("Width after round-trip = %d, want %d", w2, newWidth)
	}
}

func TestInlineShapes_SetHeight_RoundTrip(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	shape, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	newHeight := Inches(3)
	if err := shape.SetHeight(newHeight); err != nil {
		t.Fatalf("SetHeight: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	shapes, err := doc2.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	if shapes.Len() != 1 {
		t.Fatalf("expected 1 inline shape, got %d", shapes.Len())
	}
	shape2, err := shapes.Get(0)
	if err != nil {
		t.Fatalf("Get(0): %v", err)
	}
	h2, err := shape2.Height()
	if err != nil {
		t.Fatalf("Height(): %v", err)
	}
	if h2 != newHeight {
		t.Errorf("Height after round-trip = %d, want %d", h2, newHeight)
	}
}

func TestInlineShapes_SpecifiedWidthHeight(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	w := Inches(3).Emu()
	h := Inches(2).Emu()
	shape, err := doc.AddPicture(r, &w, &h)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	gotW, err := shape.Width()
	if err != nil {
		t.Fatalf("Width(): %v", err)
	}
	gotH, err := shape.Height()
	if err != nil {
		t.Fatalf("Height(): %v", err)
	}
	if gotW != Emu(w) {
		t.Errorf("Width() = %d, want %d", gotW, w)
	}
	if gotH != Emu(h) {
		t.Errorf("Height() = %d, want %d", gotH, h)
	}
}

func TestInlineShapes_MultiplePictures(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()

	for i := 0; i < 3; i++ {
		r := bytes.NewReader(pngData)
		_, err := doc.AddPicture(r, nil, nil)
		if err != nil {
			t.Fatalf("AddPicture[%d]: %v", i, err)
		}
	}

	shapes, err := doc.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	if shapes.Len() != 3 {
		t.Errorf("InlineShapes.Len() = %d, want 3", shapes.Len())
	}
}

func TestInlineShapes_Iter(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()

	r1 := bytes.NewReader(pngData)
	_, err := doc.AddPicture(r1, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}
	r2 := bytes.NewReader(pngData)
	_, err = doc.AddPicture(r2, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	shapes, err := doc.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	all := shapes.Iter()
	if len(all) != 2 {
		t.Errorf("Iter() len = %d, want 2", len(all))
	}
	for i, s := range all {
		w, err := s.Width()
		if err != nil {
			t.Errorf("Iter[%d].Width() error: %v", i, err)
		}
		if w <= 0 {
			t.Errorf("Iter[%d].Width() = %d, expected > 0", i, w)
		}
	}
}

func TestInlineShapes_Get_OutOfRange(t *testing.T) {
	doc := mustNewDoc(t)
	shapes, err := doc.InlineShapes()
	if err != nil {
		t.Fatalf("InlineShapes() error: %v", err)
	}
	_, err = shapes.Get(0)
	if err == nil {
		t.Error("expected error for Get(0) on empty InlineShapes")
	}
	_, err = shapes.Get(-1)
	if err == nil {
		t.Error("expected error for Get(-1)")
	}
}

func TestInlineShape_Type_Picture(t *testing.T) {
	doc := mustNewDoc(t)
	pngData := minimalPNG()
	r := bytes.NewReader(pngData)

	shape, err := doc.AddPicture(r, nil, nil)
	if err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	st, err := shape.Type()
	if err != nil {
		t.Fatal(err)
	}
	if st != enum.WdInlineShapeTypePicture {
		t.Errorf("Type() = %d, want %d (PICTURE)", st, enum.WdInlineShapeTypePicture)
	}
}
