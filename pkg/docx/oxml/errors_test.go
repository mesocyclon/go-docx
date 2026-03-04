package oxml

import (
	"errors"
	"fmt"
	"testing"
)

func TestParseAttrError_ErrorsAs(t *testing.T) {
	t.Parallel()

	cause := fmt.Errorf("parse failure")
	pe := &ParseAttrError{
		Element:  "w:ind",
		Attr:     "w:left",
		RawValue: "xyz",
		Err:      cause,
	}

	// Wrap it in another error to test errors.As
	wrapped := fmt.Errorf("reading paragraph: %w", pe)

	var target *ParseAttrError
	if !errors.As(wrapped, &target) {
		t.Fatal("errors.As should match *ParseAttrError")
	}
	if target.Element != "w:ind" {
		t.Errorf("Element: got %q, want %q", target.Element, "w:ind")
	}
	if target.Attr != "w:left" {
		t.Errorf("Attr: got %q, want %q", target.Attr, "w:left")
	}
}

func TestParseAttrError_NilErr(t *testing.T) {
	t.Parallel()

	pe := &ParseAttrError{
		Element:  "w:sz",
		Attr:     "w:val",
		RawValue: "",
		Err:      nil,
	}
	// Should not panic
	_ = pe.Error()
	if pe.Unwrap() != nil {
		t.Error("Unwrap with nil Err should return nil")
	}
}
