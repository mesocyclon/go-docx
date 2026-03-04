package oxml

import (
	"strings"
	"testing"
)

func TestParseXml(t *testing.T) {
	t.Parallel()

	t.Run("parses simple element", func(t *testing.T) {
		t.Parallel()
		xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
		el, err := ParseXml([]byte(xml))
		if err != nil {
			t.Fatalf("ParseXml error: %v", err)
		}
		if el.Tag != "p" {
			t.Errorf("Tag = %q, want %q", el.Tag, "p")
		}
		if el.Space != "w" {
			t.Errorf("Space = %q, want %q", el.Space, "w")
		}
	})

	t.Run("parses element with children", func(t *testing.T) {
		t.Parallel()
		xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:pPr/>
			<w:r><w:t>Hello</w:t></w:r>
		</w:p>`
		el, err := ParseXml([]byte(xml))
		if err != nil {
			t.Fatalf("ParseXml error: %v", err)
		}
		children := el.ChildElements()
		if len(children) != 2 {
			t.Fatalf("expected 2 children, got %d", len(children))
		}
		if children[0].Tag != "pPr" {
			t.Errorf("first child tag = %q, want pPr", children[0].Tag)
		}
		if children[1].Tag != "r" {
			t.Errorf("second child tag = %q, want r", children[1].Tag)
		}
	})

	t.Run("returns error on invalid XML", func(t *testing.T) {
		t.Parallel()
		_, err := ParseXml([]byte("<unclosed"))
		if err == nil {
			t.Error("expected error for invalid XML")
		}
	})

	t.Run("returns error on empty input", func(t *testing.T) {
		t.Parallel()
		_, err := ParseXml([]byte(""))
		if err == nil {
			t.Error("expected error for empty input")
		}
	})
}

func TestSerializeXml(t *testing.T) {
	t.Parallel()

	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r/></w:p>`
	el, err := ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml error: %v", err)
	}

	data, err := SerializeXml(el)
	if err != nil {
		t.Fatalf("SerializeXml error: %v", err)
	}

	s := string(data)
	if !strings.Contains(s, `standalone="yes"`) {
		t.Errorf("SerializeXml output missing standalone='yes': %s", s)
	}
	if !strings.Contains(s, `<?xml`) {
		t.Errorf("SerializeXml output missing XML declaration: %s", s)
	}
}

func TestParseAndSerializeRoundTrip(t *testing.T) {
	t.Parallel()

	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr/><w:r><w:t>Hello</w:t></w:r></w:p>`
	el, err := ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("first ParseXml error: %v", err)
	}

	data, err := SerializeXml(el)
	if err != nil {
		t.Fatalf("SerializeXml error: %v", err)
	}

	el2, err := ParseXml(data)
	if err != nil {
		t.Fatalf("second ParseXml error: %v", err)
	}

	// Verify structure is preserved
	if el2.Tag != "p" {
		t.Errorf("round-trip: root tag = %q, want p", el2.Tag)
	}
	children := el2.ChildElements()
	if len(children) != 2 {
		t.Fatalf("round-trip: expected 2 children, got %d", len(children))
	}
	if children[0].Tag != "pPr" {
		t.Errorf("round-trip: child[0] = %q, want pPr", children[0].Tag)
	}
	if children[1].Tag != "r" {
		t.Errorf("round-trip: child[1] = %q, want r", children[1].Tag)
	}

	// Verify text is preserved
	tEl := children[1].FindElement("./t")
	if tEl == nil {
		t.Fatal("round-trip: <w:t> not found inside <w:r>")
	}
	if tEl.Text() != "Hello" {
		t.Errorf("round-trip: text = %q, want %q", tEl.Text(), "Hello")
	}
}

func TestOxmlElement(t *testing.T) {
	t.Parallel()

	t.Run("creates element with correct tag and namespace", func(t *testing.T) {
		t.Parallel()
		el := OxmlElement("w:p")
		if el.Tag != "p" {
			t.Errorf("Tag = %q, want %q", el.Tag, "p")
		}
		if el.Space != "w" {
			t.Errorf("Space = %q, want %q", el.Space, "w")
		}
	})

	t.Run("adds namespace declaration attribute", func(t *testing.T) {
		t.Parallel()
		el := OxmlElement("w:p")
		uri, ok := HasNsDecl(el, "w")
		if !ok {
			t.Error("missing xmlns:w declaration")
		}
		if uri != nsmap["w"] {
			t.Errorf("xmlns:w = %q, want %q", uri, nsmap["w"])
		}
	})

	t.Run("adds additional namespace declarations", func(t *testing.T) {
		t.Parallel()
		el := OxmlElement("w:p", "r")
		if _, ok := HasNsDecl(el, "w"); !ok {
			t.Error("missing xmlns:w declaration")
		}
		if _, ok := HasNsDecl(el, "r"); !ok {
			t.Error("missing xmlns:r declaration")
		}
	})

	t.Run("creates element for different namespaces", func(t *testing.T) {
		t.Parallel()
		el := OxmlElement("a:blip")
		if el.Tag != "blip" || el.Space != "a" {
			t.Errorf("element = %s:%s, want a:blip", el.Space, el.Tag)
		}
	})
}

func TestOxmlElementWithAttrs(t *testing.T) {
	t.Parallel()

	attrs := map[string]string{"val": "center"}
	el := OxmlElementWithAttrs("w:jc", attrs)

	if el.Tag != "jc" || el.Space != "w" {
		t.Errorf("element = %s:%s, want w:jc", el.Space, el.Tag)
	}

	a := el.SelectAttr("val")
	if a == nil || a.Value != "center" {
		t.Errorf("attr 'val' = %v, want 'center'", a)
	}
}

func TestSerializeForReading(t *testing.T) {
	t.Parallel()

	el := OxmlElement("w:p")
	r := el.CreateElement("r")
	r.Space = "w"

	output := SerializeForReading(el)
	if output == "" {
		t.Error("SerializeForReading returned empty string")
	}
	// Should not contain XML declaration
	if strings.Contains(output, "<?xml") {
		t.Error("SerializeForReading should not contain XML declaration")
	}
}
