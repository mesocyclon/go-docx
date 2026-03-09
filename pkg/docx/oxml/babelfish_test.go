package oxml

import (
	"testing"
)

func TestUI2Internal(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input, want string
	}{
		{"Heading 1", "heading 1"},
		{"Heading 9", "heading 9"},
		{"Caption", "caption"},
		{"Header", "header"},
		{"Footer", "footer"},
		{"Normal", "Normal"},           // unmapped — returned as-is
		{"My Custom Style", "My Custom Style"}, // unmapped
	}
	for _, tt := range tests {
		if got := UI2Internal(tt.input); got != tt.want {
			t.Errorf("UI2Internal(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestInternal2UI(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input, want string
	}{
		{"heading 1", "Heading 1"},
		{"caption", "Caption"},
		{"header", "Header"},
		{"footer", "Footer"},
		{"Normal", "Normal"}, // unmapped
	}
	for _, tt := range tests {
		if got := Internal2UI(tt.input); got != tt.want {
			t.Errorf("Internal2UI(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}
