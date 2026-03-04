package oxml

import (
	"testing"
)

func TestCT_Settings_EvenAndOddHeadersVal(t *testing.T) {
	xml := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := ParseXml([]byte(xml))
	s := &CT_Settings{Element{e: el}}

	// Default should be false
	if s.EvenAndOddHeadersVal() {
		t.Error("expected false by default")
	}

	// Set to true
	boolTrue := true
	if err := s.SetEvenAndOddHeadersVal(&boolTrue); err != nil {
		t.Fatalf("SetEvenAndOddHeadersVal: %v", err)
	}
	if !s.EvenAndOddHeadersVal() {
		t.Error("expected true after setting")
	}

	// Set to false (should remove)
	boolFalse := false
	if err := s.SetEvenAndOddHeadersVal(&boolFalse); err != nil {
		t.Fatalf("SetEvenAndOddHeadersVal: %v", err)
	}
	if s.EvenAndOddHeadersVal() {
		t.Error("expected false after unsetting")
	}

	// Set to true again then nil (should remove)
	if err := s.SetEvenAndOddHeadersVal(&boolTrue); err != nil {
		t.Fatalf("SetEvenAndOddHeadersVal: %v", err)
	}
	if err := s.SetEvenAndOddHeadersVal(nil); err != nil {
		t.Fatalf("SetEvenAndOddHeadersVal: %v", err)
	}
	if s.EvenAndOddHeadersVal() {
		t.Error("expected false after setting nil")
	}
}
