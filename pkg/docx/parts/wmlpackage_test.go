package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

func TestImageParts_GetOrAddSameBlob_Dedup(t *testing.T) {
	ips := NewImageParts()
	blob := []byte("same image data")

	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, blob, nil)
	ips.Append(ip1)

	// Look up by hash — should find the existing part
	h, err := ip1.Hash()
	if err != nil {
		t.Fatal(err)
	}
	found, err := ips.GetByHash(h)
	if err != nil {
		t.Fatal(err)
	}
	if found != ip1 {
		t.Error("GetByHash should find existing part with same blob")
	}
}

func TestImageParts_DifferentBlobs_NoDeDup(t *testing.T) {
	ips := NewImageParts()
	blob1 := []byte("image data A")
	blob2 := []byte("image data B")

	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, blob1, nil)
	ips.Append(ip1)

	ipTmp := NewImagePart("/tmp/x.png", opc.CTPng, blob2, nil)
	h, err := ipTmp.Hash()
	if err != nil {
		t.Fatal(err)
	}
	found, err := ips.GetByHash(h)
	if err != nil {
		t.Fatal(err)
	}
	if found != nil {
		t.Error("GetByHash should not find part with different blob")
	}
}

func TestImageParts_Contains(t *testing.T) {
	ips := NewImageParts()
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("a"), nil)
	ip2 := NewImagePart("/word/media/image2.png", opc.CTPng, []byte("b"), nil)

	ips.Append(ip1)
	if !ips.Contains(ip1) {
		t.Error("Contains should return true for appended part")
	}
	if ips.Contains(ip2) {
		t.Error("Contains should return false for non-appended part")
	}
}

func TestImageParts_NextImagePartname_Sequential(t *testing.T) {
	ips := NewImageParts()
	pn := ips.nextImagePartname("png")
	if pn != "/word/media/image1.png" {
		t.Errorf("first partname = %q, want /word/media/image1.png", pn)
	}
}

func TestImageParts_NextImagePartname_ReusesGaps(t *testing.T) {
	ips := NewImageParts()
	// Simulate image1 and image3 existing (gap at image2)
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("a"), nil)
	ip3 := NewImagePart("/word/media/image3.png", opc.CTPng, []byte("b"), nil)
	ips.Append(ip1)
	ips.Append(ip3)

	pn := ips.nextImagePartname("jpg")
	// Should reuse gap at index 2
	if pn != "/word/media/image2.jpg" {
		t.Errorf("reuse gap partname = %q, want /word/media/image2.jpg", pn)
	}
}

func TestImageParts_NextImagePartname_NoGaps(t *testing.T) {
	ips := NewImageParts()
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("a"), nil)
	ip2 := NewImagePart("/word/media/image2.png", opc.CTPng, []byte("b"), nil)
	ips.Append(ip1)
	ips.Append(ip2)

	pn := ips.nextImagePartname("gif")
	if pn != "/word/media/image3.gif" {
		t.Errorf("next partname = %q, want /word/media/image3.gif", pn)
	}
}

func TestExtFromContentType(t *testing.T) {
	tests := []struct {
		ct   string
		want string
	}{
		{"image/jpeg", "jpg"},
		{"image/png", "png"},
		{"image/gif", "gif"},
		{"image/tiff", "tiff"},
		{"image/bmp", "bmp"},
		{"application/octet-stream", "bin"},
	}
	for _, tt := range tests {
		got := extFromContentType(tt.ct)
		if got != tt.want {
			t.Errorf("extFromContentType(%q) = %q, want %q", tt.ct, got, tt.want)
		}
	}
}

func TestWmlPackage_GetOrAddImagePart_Dedup(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	wp := NewWmlPackage(pkg)

	blob := []byte("test image")
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, blob, nil)

	result1, err := wp.GetOrAddImagePart(ip1)
	if err != nil {
		t.Fatal(err)
	}
	if result1 == nil {
		t.Fatal("GetOrAddImagePart returned nil")
	}

	// Add same blob again — should return existing
	ip2 := NewImagePart("/tmp/temp.png", opc.CTPng, blob, nil)
	result2, err := wp.GetOrAddImagePart(ip2)
	if err != nil {
		t.Fatal(err)
	}
	if result2 != result1 {
		t.Error("GetOrAddImagePart should dedup same blob")
	}
	if wp.ImageParts().Len() != 1 {
		t.Errorf("image parts count = %d, want 1", wp.ImageParts().Len())
	}
}

func TestWmlPackage_GetOrAddImagePart_Different(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	wp := NewWmlPackage(pkg)

	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, []byte("A"), nil)
	ip2 := NewImagePart("/tmp/temp.png", opc.CTPng, []byte("B"), nil)

	if _, err := wp.GetOrAddImagePart(ip1); err != nil {
		t.Fatal(err)
	}
	if _, err := wp.GetOrAddImagePart(ip2); err != nil {
		t.Fatal(err)
	}
	if wp.ImageParts().Len() != 2 {
		t.Errorf("image parts count = %d, want 2", wp.ImageParts().Len())
	}
}
