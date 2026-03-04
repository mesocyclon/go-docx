package oxml

import (
	"testing"

	"github.com/beevik/etree"
)

// buildTestElement creates a <w:p> with <w:pPr/>, <w:r/>, <w:r/> children for testing.
func buildTestElement() *Element {
	p := etree.NewElement("p")
	p.Space = "w"

	pPr := p.CreateElement("pPr")
	pPr.Space = "w"

	r1 := p.CreateElement("r")
	r1.Space = "w"
	r1.CreateAttr("id", "1")

	r2 := p.CreateElement("r")
	r2.Space = "w"
	r2.CreateAttr("id", "2")

	return &Element{e: p}
}

func TestElementTag(t *testing.T) {
	t.Parallel()
	el := buildTestElement()
	tag := el.Tag()
	want := Qn("w:p")
	if tag != want {
		t.Errorf("Tag() = %q, want %q", tag, want)
	}
}

func TestFindChild(t *testing.T) {
	t.Parallel()
	el := buildTestElement()

	t.Run("finds existing child", func(t *testing.T) {
		t.Parallel()
		pPr := el.FindChild("w:pPr")
		if pPr == nil {
			t.Fatal("FindChild('w:pPr') returned nil, want non-nil")
		}
		if pPr.Tag != "pPr" || pPr.Space != "w" {
			t.Errorf("found wrong element: Space=%q Tag=%q", pPr.Space, pPr.Tag)
		}
	})

	t.Run("returns nil for nonexistent child", func(t *testing.T) {
		t.Parallel()
		result := el.FindChild("w:nonexistent")
		if result != nil {
			t.Error("FindChild('w:nonexistent') should return nil")
		}
	})
}

func TestFindAllChildren(t *testing.T) {
	t.Parallel()
	el := buildTestElement()

	t.Run("finds all matching children", func(t *testing.T) {
		t.Parallel()
		runs := el.FindAllChildren("w:r")
		if len(runs) != 2 {
			t.Fatalf("FindAllChildren('w:r') returned %d, want 2", len(runs))
		}
	})

	t.Run("returns empty for no match", func(t *testing.T) {
		t.Parallel()
		result := el.FindAllChildren("w:nothing")
		if len(result) != 0 {
			t.Errorf("FindAllChildren('w:nothing') returned %d, want 0", len(result))
		}
	})

	t.Run("finds single matching child", func(t *testing.T) {
		t.Parallel()
		result := el.FindAllChildren("w:pPr")
		if len(result) != 1 {
			t.Errorf("FindAllChildren('w:pPr') returned %d, want 1", len(result))
		}
	})
}

func TestFirstChildIn(t *testing.T) {
	t.Parallel()
	el := buildTestElement()

	t.Run("returns first matching child in priority order", func(t *testing.T) {
		t.Parallel()
		child := el.FirstChildIn("w:pPr", "w:r")
		if child == nil {
			t.Fatal("FirstChildIn returned nil")
		}
		if child.Tag != "pPr" {
			t.Errorf("FirstChildIn returned %q, want pPr", child.Tag)
		}
	})

	t.Run("returns second tag if first not found", func(t *testing.T) {
		t.Parallel()
		child := el.FirstChildIn("w:nonexistent", "w:r")
		if child == nil {
			t.Fatal("FirstChildIn returned nil")
		}
		if child.Tag != "r" {
			t.Errorf("FirstChildIn returned %q, want r", child.Tag)
		}
	})

	t.Run("returns nil if none found", func(t *testing.T) {
		t.Parallel()
		child := el.FirstChildIn("w:foo", "w:bar")
		if child != nil {
			t.Error("FirstChildIn should return nil when no tags match")
		}
	})
}

func TestInsertElementBefore(t *testing.T) {
	t.Parallel()

	t.Run("inserts before first successor", func(t *testing.T) {
		el := buildTestElement()

		newChild := etree.NewElement("bookmarkStart")
		newChild.Space = "w"

		// Insert before w:r (successor). Should go between pPr and first r.
		el.InsertElementBefore(newChild, "w:r")

		children := el.e.ChildElements()
		if len(children) != 4 {
			t.Fatalf("expected 4 children, got %d", len(children))
		}
		if children[0].Tag != "pPr" {
			t.Errorf("child[0] = %q, want pPr", children[0].Tag)
		}
		if children[1].Tag != "bookmarkStart" {
			t.Errorf("child[1] = %q, want bookmarkStart", children[1].Tag)
		}
		if children[2].Tag != "r" {
			t.Errorf("child[2] = %q, want r", children[2].Tag)
		}
	})

	t.Run("appends when no successor found", func(t *testing.T) {
		el := buildTestElement()

		newChild := etree.NewElement("sectPr")
		newChild.Space = "w"

		el.InsertElementBefore(newChild, "w:nonexistent")

		children := el.e.ChildElements()
		last := children[len(children)-1]
		if last.Tag != "sectPr" {
			t.Errorf("last child = %q, want sectPr", last.Tag)
		}
	})

	t.Run("appends when no successors given", func(t *testing.T) {
		el := buildTestElement()

		newChild := etree.NewElement("custom")
		newChild.Space = "w"

		el.InsertElementBefore(newChild) // no successors

		children := el.e.ChildElements()
		last := children[len(children)-1]
		if last.Tag != "custom" {
			t.Errorf("last child = %q, want custom", last.Tag)
		}
	})

	t.Run("inserts before second successor if first not found", func(t *testing.T) {
		el := buildTestElement()

		newChild := etree.NewElement("rPr")
		newChild.Space = "w"

		// w:foo doesn't exist, but w:r does
		el.InsertElementBefore(newChild, "w:foo", "w:r")

		children := el.e.ChildElements()
		// Should be: pPr, rPr, r, r
		if len(children) != 4 {
			t.Fatalf("expected 4 children, got %d", len(children))
		}
		if children[1].Tag != "rPr" {
			t.Errorf("child[1] = %q, want rPr", children[1].Tag)
		}
	})
}

func TestRemoveAll(t *testing.T) {
	t.Parallel()
	el := buildTestElement()

	el.RemoveAll("w:r")

	children := el.e.ChildElements()
	if len(children) != 1 {
		t.Fatalf("after RemoveAll('w:r'), expected 1 child, got %d", len(children))
	}
	if children[0].Tag != "pPr" {
		t.Errorf("remaining child = %q, want pPr", children[0].Tag)
	}
}

func TestRemoveAllMultipleTags(t *testing.T) {
	t.Parallel()
	el := buildTestElement()

	el.RemoveAll("w:pPr", "w:r")

	children := el.e.ChildElements()
	if len(children) != 0 {
		t.Errorf("after RemoveAll('w:pPr', 'w:r'), expected 0 children, got %d", len(children))
	}
}

func TestRemove(t *testing.T) {
	t.Parallel()
	el := buildTestElement()

	pPr := el.FindChild("w:pPr")
	el.Remove(pPr)

	if el.FindChild("w:pPr") != nil {
		t.Error("pPr should have been removed")
	}
	if len(el.e.ChildElements()) != 2 {
		t.Errorf("expected 2 children after remove, got %d", len(el.e.ChildElements()))
	}
}

func TestGetSetAttr(t *testing.T) {
	t.Parallel()

	e := etree.NewElement("p")
	e.Space = "w"
	e.CreateAttr("val", "test123")
	el := &Element{e: e}

	t.Run("get existing attribute", func(t *testing.T) {
		t.Parallel()
		val, ok := el.GetAttr("val")
		if !ok {
			t.Fatal("GetAttr('val') returned false")
		}
		if val != "test123" {
			t.Errorf("GetAttr('val') = %q, want %q", val, "test123")
		}
	})

	t.Run("get nonexistent attribute", func(t *testing.T) {
		t.Parallel()
		_, ok := el.GetAttr("nonexistent")
		if ok {
			t.Error("GetAttr('nonexistent') should return false")
		}
	})
}

func TestSetAttr(t *testing.T) {
	t.Parallel()

	e := etree.NewElement("p")
	e.Space = "w"
	el := &Element{e: e}

	el.SetAttr("val", "hello")

	val, ok := el.GetAttr("val")
	if !ok || val != "hello" {
		t.Errorf("after SetAttr, GetAttr = (%q, %v), want ('hello', true)", val, ok)
	}
}

func TestRemoveAttr(t *testing.T) {
	t.Parallel()

	e := etree.NewElement("p")
	e.Space = "w"
	e.CreateAttr("val", "test")
	el := &Element{e: e}

	el.RemoveAttr("val")

	_, ok := el.GetAttr("val")
	if ok {
		t.Error("attribute should have been removed")
	}
}

func TestTextAndSetText(t *testing.T) {
	t.Parallel()

	e := etree.NewElement("t")
	e.Space = "w"
	e.SetText("Hello World")
	el := &Element{e: e}

	if el.Text() != "Hello World" {
		t.Errorf("Text() = %q, want %q", el.Text(), "Hello World")
	}

	el.SetText("Goodbye")
	if el.Text() != "Goodbye" {
		t.Errorf("after SetText, Text() = %q, want %q", el.Text(), "Goodbye")
	}
}

func TestAddSubElement(t *testing.T) {
	t.Parallel()

	e := etree.NewElement("p")
	e.Space = "w"
	el := &Element{e: e}

	child := el.AddSubElement("w:r")
	if child.Tag != "r" {
		t.Errorf("AddSubElement tag = %q, want %q", child.Tag, "r")
	}
	if child.Space != "w" {
		t.Errorf("AddSubElement space = %q, want %q", child.Space, "w")
	}

	children := el.e.ChildElements()
	if len(children) != 1 {
		t.Fatalf("expected 1 child, got %d", len(children))
	}
}

func TestXml(t *testing.T) {
	t.Parallel()

	e := etree.NewElement("p")
	e.Space = "w"
	r := e.CreateElement("r")
	r.Space = "w"
	el := &Element{e: e}

	xml := el.Xml()
	if xml == "" {
		t.Error("Xml() returned empty string")
	}
}

func TestNewElement(t *testing.T) {
	t.Parallel()

	t.Run("wraps non-nil element", func(t *testing.T) {
		t.Parallel()
		e := etree.NewElement("p")
		el := NewElement(e)
		if el == nil || el.e != e {
			t.Error("NewElement should wrap the etree element")
		}
	})

	t.Run("returns nil for nil", func(t *testing.T) {
		t.Parallel()
		el := NewElement(nil)
		if el != nil {
			t.Error("NewElement(nil) should return nil")
		}
	})
}
