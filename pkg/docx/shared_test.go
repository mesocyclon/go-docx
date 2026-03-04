package docx

import (
	"errors"
	"io"
	"math"
	"testing"
)

func almostEqual(a, b, tolerance float64) bool {
	return math.Abs(a-b) < tolerance
}

func TestLengthInches(t *testing.T) {
	t.Parallel()
	l := Inches(1.0)
	if l.Emu() != 914400 {
		t.Errorf("Inches(1.0).Emu() = %d, want 914400", l.Emu())
	}
}

func TestLengthCmToInches(t *testing.T) {
	t.Parallel()
	l := Cm(2.54)
	if !almostEqual(l.Inches(), 1.0, 0.001) {
		t.Errorf("Cm(2.54).Inches() = %f, want ≈1.0", l.Inches())
	}
}

func TestLengthPt(t *testing.T) {
	t.Parallel()
	l := Pt(12)
	if l.Emu() != 152400 {
		t.Errorf("Pt(12).Emu() = %d, want 152400", l.Emu())
	}
}

func TestLengthTwips(t *testing.T) {
	t.Parallel()
	l := Twips(20)
	if l.Emu() != 12700 {
		t.Errorf("Twips(20).Emu() = %d, want 12700", l.Emu())
	}
}

func TestLengthMmToInches(t *testing.T) {
	t.Parallel()
	l := Mm(25.4)
	if !almostEqual(l.Inches(), 1.0, 0.001) {
		t.Errorf("Mm(25.4).Inches() = %f, want ≈1.0", l.Inches())
	}
}

func TestLengthRoundTrip(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name      string
		construct func() Length
		extract   func(Length) float64
		original  float64
	}{
		{"Inches", func() Length { return Inches(1.5) }, func(l Length) float64 { return l.Inches() }, 1.5},
		{"Cm", func() Length { return Cm(3.5) }, func(l Length) float64 { return l.Cm() }, 3.5},
		{"Mm", func() Length { return Mm(100.0) }, func(l Length) float64 { return l.Mm() }, 100.0},
		{"Pt", func() Length { return Pt(72.0) }, func(l Length) float64 { return l.Pt() }, 72.0},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			l := tc.construct()
			got := tc.extract(l)
			if !almostEqual(got, tc.original, 0.001) {
				t.Errorf("%s round-trip: got %f, want ≈%f", tc.name, got, tc.original)
			}
		})
	}
}

func TestLengthTwipsRoundTrip(t *testing.T) {
	t.Parallel()
	l := Twips(240)
	if l.Twips() != 240 {
		t.Errorf("Twips(240).Twips() = %d, want 240", l.Twips())
	}
}

func TestLengthEmu(t *testing.T) {
	t.Parallel()
	l := Emu(914400)
	if l.Emu() != 914400 {
		t.Errorf("Emu(914400).Emu() = %d, want 914400", l.Emu())
	}
	if !almostEqual(l.Inches(), 1.0, 0.0001) {
		t.Errorf("Emu(914400).Inches() = %f, want 1.0", l.Inches())
	}
}

func TestRGBColorString(t *testing.T) {
	t.Parallel()
	c := NewRGBColor(0x3C, 0x2F, 0x80)
	if c.String() != "3C2F80" {
		t.Errorf("RGBColor{0x3C,0x2F,0x80}.String() = %q, want %q", c.String(), "3C2F80")
	}
}

func TestRGBColorFromString(t *testing.T) {
	t.Parallel()
	c, err := RGBColorFromString("3C2F80")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	want := NewRGBColor(0x3C, 0x2F, 0x80)
	if c != want {
		t.Errorf("RGBColorFromString(\"3C2F80\") = %v, want %v", c, want)
	}
}

func TestRGBColorFromStringLowercase(t *testing.T) {
	t.Parallel()
	c, err := RGBColorFromString("3c2f80")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	want := NewRGBColor(0x3C, 0x2F, 0x80)
	if c != want {
		t.Errorf("RGBColorFromString(\"3c2f80\") = %v, want %v", c, want)
	}
}

func TestRGBColorFromStringInvalid(t *testing.T) {
	t.Parallel()
	_, err := RGBColorFromString("ZZZZZZ")
	if err == nil {
		t.Error("expected error for invalid hex string, got nil")
	}
}

func TestRGBColorFromStringWrongLength(t *testing.T) {
	t.Parallel()
	_, err := RGBColorFromString("FFF")
	if err == nil {
		t.Error("expected error for short hex string, got nil")
	}
}

func TestRGBColorRoundTrip(t *testing.T) {
	t.Parallel()
	original := "FF8800"
	c, err := RGBColorFromString(original)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if c.String() != original {
		t.Errorf("round-trip: got %q, want %q", c.String(), original)
	}
}

func TestRGBColorComponents(t *testing.T) {
	t.Parallel()
	c := NewRGBColor(10, 20, 30)
	if c.R() != 10 || c.G() != 20 || c.B() != 30 {
		t.Errorf("components: got R=%d G=%d B=%d, want 10,20,30", c.R(), c.G(), c.B())
	}
}

func TestRGBColorZeroValue(t *testing.T) {
	t.Parallel()
	var c RGBColor
	if c.String() != "000000" {
		t.Errorf("zero RGBColor.String() = %q, want %q", c.String(), "000000")
	}
}

func TestErrorTypes(t *testing.T) {
	t.Parallel()

	// Basic message formatting, no cause.
	e1 := NewInvalidXmlError(nil, "bad xml: %s", "test")
	if e1.Error() != "bad xml: test" {
		t.Errorf("InvalidXmlError.Error() = %q", e1.Error())
	}
	if e1.Unwrap() != nil {
		t.Error("Unwrap() should be nil when no cause")
	}

	e2 := NewPackageNotFoundError(nil, "not found")
	if e2.Error() != "not found" {
		t.Errorf("PackageNotFoundError.Error() = %q", e2.Error())
	}

	e3 := NewInvalidSpanError(nil, "bad span")
	if e3.Error() != "bad span" {
		t.Errorf("InvalidSpanError.Error() = %q", e3.Error())
	}

	// Wrapping a cause — errors.Is sees through.
	wrapped := NewDocxError(io.ErrUnexpectedEOF, "reading chunk")
	if !errors.Is(wrapped, io.ErrUnexpectedEOF) {
		t.Error("errors.Is should find io.ErrUnexpectedEOF through Unwrap")
	}

	// errors.As matches the typed wrapper.
	xmlErr := NewInvalidXmlError(io.EOF, "parse failed")
	var target *InvalidXmlError
	if !errors.As(xmlErr, &target) {
		t.Error("errors.As should match *InvalidXmlError")
	}
	if !errors.Is(xmlErr, io.EOF) {
		t.Error("errors.Is should find io.EOF through InvalidXmlError → DocxError.Unwrap")
	}
}
