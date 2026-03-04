package opc

import (
	"testing"
)

func TestNewPackURI(t *testing.T) {
	tests := []struct {
		input    string
		expected PackURI
	}{
		{"word/document.xml", "/word/document.xml"},
		{"/word/document.xml", "/word/document.xml"},
		{"/", "/"},
	}
	for _, tt := range tests {
		got := NewPackURI(tt.input)
		if got != tt.expected {
			t.Errorf("NewPackURI(%q) = %q, want %q", tt.input, got, tt.expected)
		}
	}
}

func TestPackURI_BaseURI(t *testing.T) {
	tests := []struct {
		uri      PackURI
		expected string
	}{
		{"/word/document.xml", "/word"},
		{"/word/media/image1.png", "/word/media"},
		{"/", "/"},
	}
	for _, tt := range tests {
		got := tt.uri.BaseURI()
		if got != tt.expected {
			t.Errorf("PackURI(%q).BaseURI() = %q, want %q", tt.uri, got, tt.expected)
		}
	}
}

func TestPackURI_Filename(t *testing.T) {
	tests := []struct {
		uri      PackURI
		expected string
	}{
		{"/word/document.xml", "document.xml"},
		{"/word/media/image1.png", "image1.png"},
		{"/", ""},
	}
	for _, tt := range tests {
		got := tt.uri.Filename()
		if got != tt.expected {
			t.Errorf("PackURI(%q).Filename() = %q, want %q", tt.uri, got, tt.expected)
		}
	}
}

func TestPackURI_Ext(t *testing.T) {
	tests := []struct {
		uri      PackURI
		expected string
	}{
		{"/word/document.xml", "xml"},
		{"/word/media/image1.png", "png"},
	}
	for _, tt := range tests {
		got := tt.uri.Ext()
		if got != tt.expected {
			t.Errorf("PackURI(%q).Ext() = %q, want %q", tt.uri, got, tt.expected)
		}
	}
}

func TestPackURI_RelsURI(t *testing.T) {
	tests := []struct {
		uri      PackURI
		expected PackURI
	}{
		{"/word/document.xml", "/word/_rels/document.xml.rels"},
		{"/", "/_rels/.rels"},
	}
	for _, tt := range tests {
		got := tt.uri.RelsURI()
		if got != tt.expected {
			t.Errorf("PackURI(%q).RelsURI() = %q, want %q", tt.uri, got, tt.expected)
		}
	}
}

func TestPackURI_Membername(t *testing.T) {
	tests := []struct {
		uri      PackURI
		expected string
	}{
		{"/word/document.xml", "word/document.xml"},
		{"/", ""},
	}
	for _, tt := range tests {
		got := tt.uri.Membername()
		if got != tt.expected {
			t.Errorf("PackURI(%q).Membername() = %q, want %q", tt.uri, got, tt.expected)
		}
	}
}

func TestFromRelRef(t *testing.T) {
	tests := []struct {
		base     string
		ref      string
		expected PackURI
	}{
		{"/word", "document.xml", "/word/document.xml"},
		{"/word", "media/image1.png", "/word/media/image1.png"},
		{"/word/slides", "../slideLayouts/slideLayout1.xml", "/word/slideLayouts/slideLayout1.xml"},
		{"/", "word/document.xml", "/word/document.xml"},
	}
	for _, tt := range tests {
		got := FromRelRef(tt.base, tt.ref)
		if got != tt.expected {
			t.Errorf("FromRelRef(%q, %q) = %q, want %q", tt.base, tt.ref, got, tt.expected)
		}
	}
}

func TestPackURI_RelativeRef(t *testing.T) {
	tests := []struct {
		uri      PackURI
		base     string
		expected string
	}{
		{"/word/document.xml", "/", "word/document.xml"},
		{"/word/styles.xml", "/word", "styles.xml"},
		{"/word/media/image1.png", "/word", "media/image1.png"},
	}
	for _, tt := range tests {
		got := tt.uri.RelativeRef(tt.base)
		if got != tt.expected {
			t.Errorf("PackURI(%q).RelativeRef(%q) = %q, want %q", tt.uri, tt.base, got, tt.expected)
		}
	}
}

func TestPackURI_Idx(t *testing.T) {
	tests := []struct {
		uri     PackURI
		wantIdx int
		wantOK  bool
	}{
		{"/word/media/image1.png", 1, true},
		{"/word/media/image21.jpg", 21, true},
		{"/word/header3.xml", 3, true},
		{"/word/document.xml", 0, false},     // no trailing digits
		{"/word/styles.xml", 0, false},       // no trailing digits
		{"/", 0, false},                      // empty filename
		{"/word/media/image0.png", 0, false}, // 0 is not matched by Python regex [1-9][0-9]*
	}
	for _, tt := range tests {
		gotIdx, gotOK := tt.uri.Idx()
		if gotIdx != tt.wantIdx || gotOK != tt.wantOK {
			t.Errorf("PackURI(%q).Idx() = (%d, %v), want (%d, %v)",
				tt.uri, gotIdx, gotOK, tt.wantIdx, tt.wantOK)
		}
	}
}
