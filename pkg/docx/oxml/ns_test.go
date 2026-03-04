package oxml

import (
	"testing"
)

func TestQn(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name string
		tag  string
		want string
	}{
		{
			name: "w:p resolves to wordprocessingml namespace",
			tag:  "w:p",
			want: "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p",
		},
		{
			name: "w:body resolves correctly",
			tag:  "w:body",
			want: "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body",
		},
		{
			name: "r:id resolves to relationships namespace",
			tag:  "r:id",
			want: "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
		},
		{
			name: "a:blip resolves to drawingml namespace",
			tag:  "a:blip",
			want: "{http://schemas.openxmlformats.org/drawingml/2006/main}blip",
		},
		{
			name: "wp:inline resolves to wordprocessingDrawing namespace",
			tag:  "wp:inline",
			want: "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline",
		},
		{
			name: "pic:pic resolves to picture namespace",
			tag:  "pic:pic",
			want: "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic",
		},
		{
			name: "w14:conflictMode resolves to word 2010 namespace",
			tag:  "w14:conflictMode",
			want: "{http://schemas.microsoft.com/office/word/2010/wordml}conflictMode",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			got := Qn(tt.tag)
			if got != tt.want {
				t.Errorf("Qn(%q) = %q, want %q", tt.tag, got, tt.want)
			}
		})
	}
}

func TestQnPanicsOnUnknownPrefix(t *testing.T) {
	t.Parallel()
	defer func() {
		if r := recover(); r == nil {
			t.Error("expected panic for unknown prefix, got nil")
		}
	}()
	Qn("unknown:tag")
}

func TestQnNoPrefix(t *testing.T) {
	t.Parallel()
	got := Qn("simpleTag")
	if got != "simpleTag" {
		t.Errorf("Qn(%q) = %q, want %q", "simpleTag", got, "simpleTag")
	}
}

func TestNewNSPTag(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:p")
	if tag.Prefix() != "w" {
		t.Errorf("Prefix() = %q, want %q", tag.Prefix(), "w")
	}
	if tag.LocalPart() != "p" {
		t.Errorf("LocalPart() = %q, want %q", tag.LocalPart(), "p")
	}
	if tag.NsURI() != nsmap["w"] {
		t.Errorf("NsURI() = %q, want %q", tag.NsURI(), nsmap["w"])
	}
	if tag.String() != "w:p" {
		t.Errorf("String() = %q, want %q", tag.String(), "w:p")
	}
}

func TestNewNSPTagClarkName(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:body")
	want := Qn("w:body")
	if tag.ClarkName() != want {
		t.Errorf("ClarkName() = %q, want %q", tag.ClarkName(), want)
	}
}

func TestNSPTagFromClark(t *testing.T) {
	t.Parallel()
	clark := Qn("w:p")
	tag := NSPTagFromClark(clark)
	if tag.String() != "w:p" {
		t.Errorf("String() = %q, want %q", tag.String(), "w:p")
	}
	if tag.ClarkName() != clark {
		t.Errorf("ClarkName() = %q, want %q", tag.ClarkName(), clark)
	}
}

func TestNSPTagRoundTrip(t *testing.T) {
	t.Parallel()
	tags := []string{"w:p", "w:body", "r:id", "a:blip", "wp:inline", "pic:pic", "w14:conflictMode"}
	for _, nstag := range tags {
		t.Run(nstag, func(t *testing.T) {
			t.Parallel()
			clark := NewNSPTag(nstag).ClarkName()
			back := NSPTagFromClark(clark)
			if back.String() != nstag {
				t.Errorf("round-trip failed: %q → %q → %q", nstag, clark, back.String())
			}
		})
	}
}

func TestNsPfxMap(t *testing.T) {
	t.Parallel()
	m := NsPfxMap("w", "r")
	if len(m) != 2 {
		t.Fatalf("NsPfxMap returned %d entries, want 2", len(m))
	}
	if m["w"] != nsmap["w"] {
		t.Errorf("m[w] = %q, want %q", m["w"], nsmap["w"])
	}
	if m["r"] != nsmap["r"] {
		t.Errorf("m[r] = %q, want %q", m["r"], nsmap["r"])
	}
}

func TestNSPTagnsmap(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:p")
	m := tag.NsMap()
	if len(m) != 1 {
		t.Fatalf("nsmap returned %d entries, want 1", len(m))
	}
	if m["w"] != nsmap["w"] {
		t.Errorf("m[w] = %q, want %q", m["w"], nsmap["w"])
	}
}

func TestPfxmapIsInverseOfnsmap(t *testing.T) {
	t.Parallel()
	for pfx, uri := range nsmap {
		got, ok := pfxmap[uri]
		if !ok {
			t.Errorf("Pfxmap missing URI %q (prefix %q)", uri, pfx)
			continue
		}
		if got != pfx {
			t.Errorf("Pfxmap[%q] = %q, want %q", uri, got, pfx)
		}
	}
}

// --------------------------------------------------------------------------
// Safe API tests
// --------------------------------------------------------------------------

func TestTryQn(t *testing.T) {
	t.Parallel()

	t.Run("valid prefixed tag", func(t *testing.T) {
		t.Parallel()
		got, err := TryQn("w:p")
		if err != nil {
			t.Fatalf("TryQn(\"w:p\") returned error: %v", err)
		}
		want := "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
		if got != want {
			t.Errorf("TryQn(\"w:p\") = %q, want %q", got, want)
		}
	})

	t.Run("no prefix passes through", func(t *testing.T) {
		t.Parallel()
		got, err := TryQn("simpleTag")
		if err != nil {
			t.Fatalf("TryQn(\"simpleTag\") returned error: %v", err)
		}
		if got != "simpleTag" {
			t.Errorf("TryQn(\"simpleTag\") = %q, want %q", got, "simpleTag")
		}
	})

	t.Run("unknown prefix returns error", func(t *testing.T) {
		t.Parallel()
		_, err := TryQn("unknown:tag")
		if err == nil {
			t.Error("TryQn(\"unknown:tag\") expected error, got nil")
		}
	})
}

func TestParseNSPTag(t *testing.T) {
	t.Parallel()

	t.Run("valid tag", func(t *testing.T) {
		t.Parallel()
		tag, err := ParseNSPTag("w:p")
		if err != nil {
			t.Fatalf("ParseNSPTag(\"w:p\") returned error: %v", err)
		}
		if tag.Prefix() != "w" {
			t.Errorf("Prefix() = %q, want %q", tag.Prefix(), "w")
		}
		if tag.LocalPart() != "p" {
			t.Errorf("LocalPart() = %q, want %q", tag.LocalPart(), "p")
		}
		if tag.NsURI() != nsmap["w"] {
			t.Errorf("NsURI() = %q, want %q", tag.NsURI(), nsmap["w"])
		}
	})

	t.Run("no colon returns error", func(t *testing.T) {
		t.Parallel()
		_, err := ParseNSPTag("noprefix")
		if err == nil {
			t.Error("ParseNSPTag(\"noprefix\") expected error, got nil")
		}
	})

	t.Run("unknown prefix returns error", func(t *testing.T) {
		t.Parallel()
		_, err := ParseNSPTag("zzz:tag")
		if err == nil {
			t.Error("ParseNSPTag(\"zzz:tag\") expected error, got nil")
		}
	})
}

func TestParseNSPTagFromClark(t *testing.T) {
	t.Parallel()

	t.Run("valid clark notation", func(t *testing.T) {
		t.Parallel()
		clark := "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
		tag, err := ParseNSPTagFromClark(clark)
		if err != nil {
			t.Fatalf("ParseNSPTagFromClark(%q) returned error: %v", clark, err)
		}
		if tag.String() != "w:p" {
			t.Errorf("String() = %q, want %q", tag.String(), "w:p")
		}
	})

	t.Run("empty string returns error", func(t *testing.T) {
		t.Parallel()
		_, err := ParseNSPTagFromClark("")
		if err == nil {
			t.Error("ParseNSPTagFromClark(\"\") expected error, got nil")
		}
	})

	t.Run("no opening brace returns error", func(t *testing.T) {
		t.Parallel()
		_, err := ParseNSPTagFromClark("noBrace")
		if err == nil {
			t.Error("expected error for missing brace, got nil")
		}
	})

	t.Run("no closing brace returns error", func(t *testing.T) {
		t.Parallel()
		_, err := ParseNSPTagFromClark("{http://example.com")
		if err == nil {
			t.Error("expected error for missing closing brace, got nil")
		}
	})

	t.Run("unknown URI returns error", func(t *testing.T) {
		t.Parallel()
		_, err := ParseNSPTagFromClark("{http://unknown.example.com}tag")
		if err == nil {
			t.Error("expected error for unknown URI, got nil")
		}
	})
}

func TestTryOxmlElement(t *testing.T) {
	t.Parallel()

	t.Run("valid tag creates element", func(t *testing.T) {
		t.Parallel()
		el, err := TryOxmlElement("w:p")
		if err != nil {
			t.Fatalf("TryOxmlElement(\"w:p\") returned error: %v", err)
		}
		if el.Space != "w" || el.Tag != "p" {
			t.Errorf("element Space=%q Tag=%q, want Space=\"w\" Tag=\"p\"", el.Space, el.Tag)
		}
	})

	t.Run("with extra nsDecls", func(t *testing.T) {
		t.Parallel()
		el, err := TryOxmlElement("w:p", "r")
		if err != nil {
			t.Fatalf("TryOxmlElement returned error: %v", err)
		}
		if el == nil {
			t.Fatal("expected non-nil element")
		}
		// Verify both namespace declarations exist
		hasW, hasR := false, false
		for _, attr := range el.Attr {
			if attr.Space == "xmlns" && attr.Key == "w" {
				hasW = true
			}
			if attr.Space == "xmlns" && attr.Key == "r" {
				hasR = true
			}
		}
		if !hasW {
			t.Error("missing xmlns:w declaration")
		}
		if !hasR {
			t.Error("missing xmlns:r declaration")
		}
	})

	t.Run("invalid tag returns error", func(t *testing.T) {
		t.Parallel()
		_, err := TryOxmlElement("noprefix")
		if err == nil {
			t.Error("TryOxmlElement(\"noprefix\") expected error, got nil")
		}
	})

	t.Run("unknown prefix returns error", func(t *testing.T) {
		t.Parallel()
		_, err := TryOxmlElement("zzz:tag")
		if err == nil {
			t.Error("TryOxmlElement(\"zzz:tag\") expected error, got nil")
		}
	})
}
