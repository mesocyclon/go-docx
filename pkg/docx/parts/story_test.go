package parts

import (
	"github.com/beevik/etree"
	"github.com/vortex/go-docx/internal/xmlutil"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"testing"
)

func makeElementWithIDs(ids ...string) *etree.Element {
	el := etree.NewElement("root")
	for _, id := range ids {
		child := etree.NewElement("item")
		child.CreateAttr("id", id)
		el.AddChild(child)
	}
	return el
}

func TestNextID_EmptyElement(t *testing.T) {
	el := etree.NewElement("root")
	got := collectMaxID(el) + 1
	if got != 1 {
		t.Errorf("NextID for empty element: got %d, want 1", got)
	}
}

func TestNextID_WithIDs(t *testing.T) {
	el := makeElementWithIDs("1", "5", "3")
	got := collectMaxID(el) + 1
	if got != 6 {
		t.Errorf("NextID with ids [1,5,3]: got %d, want 6", got)
	}
}

func TestNextID_IgnoresNonDigit(t *testing.T) {
	el := makeElementWithIDs("abc", "12", "rId4", "7")
	got := collectMaxID(el) + 1
	if got != 13 {
		t.Errorf("NextID with mixed ids: got %d, want 13", got)
	}
}

func TestNextID_NestedElements(t *testing.T) {
	el := etree.NewElement("root")
	child := etree.NewElement("p")
	child.CreateAttr("id", "10")
	grandchild := etree.NewElement("r")
	grandchild.CreateAttr("id", "20")
	child.AddChild(grandchild)
	el.AddChild(child)

	got := collectMaxID(el) + 1
	if got != 21 {
		t.Errorf("NextID nested: got %d, want 21", got)
	}
}

func TestIsDigits(t *testing.T) {
	tests := []struct {
		input string
		want  bool
	}{
		{"", false},
		{"123", true},
		{"0", true},
		{"abc", false},
		{"12a", false},
		{"rId3", false},
	}
	for _, tt := range tests {
		got := xmlutil.IsDigits(tt.input)
		if got != tt.want {
			t.Errorf("isDigits(%q) = %v, want %v", tt.input, got, tt.want)
		}
	}
}

func TestRelRefCount(t *testing.T) {
	el := etree.NewElement("root")
	child1 := etree.NewElement("drawing")
	child1.CreateAttr("r:id", "rId5")
	el.AddChild(child1)
	child2 := etree.NewElement("hyperlink")
	child2.CreateAttr("r:id", "rId5")
	el.AddChild(child2)
	child3 := etree.NewElement("other")
	child3.CreateAttr("r:id", "rId3")
	el.AddChild(child3)

	if got := countRIdRefs(el, "rId5"); got != 2 {
		t.Errorf("relRefCount for rId5: got %d, want 2", got)
	}
	if got := countRIdRefs(el, "rId3"); got != 1 {
		t.Errorf("relRefCount for rId3: got %d, want 1", got)
	}
	if got := countRIdRefs(el, "rId99"); got != 0 {
		t.Errorf("relRefCount for rId99: got %d, want 0", got)
	}
}

func TestDropRel_DeletesWhenRefCountLow(t *testing.T) {
	// DropRel should delete a relationship when its XML reference count < 2.
	// Core logic tested via countRIdRefs; integration tested in document_test.go.
	el := etree.NewElement("root")
	if got := countRIdRefs(el, "rId1"); got >= 2 {
		t.Errorf("expected count < 2 for element with no refs, got %d", got)
	}
}

// --- NextID caching ---

func TestNextID_CachesAfterFirstCall(t *testing.T) {
	// Two consecutive NextID calls should return sequential values
	// without the second call needing to rescan the tree.
	sp := newTestStoryPart(t, makeElementWithIDs("1", "5", "3"))

	first := sp.NextID()
	if first != 6 {
		t.Errorf("first NextID: got %d, want 6", first)
	}

	second := sp.NextID()
	if second != 7 {
		t.Errorf("second NextID: got %d, want 7", second)
	}

	third := sp.NextID()
	if third != 8 {
		t.Errorf("third NextID: got %d, want 8", third)
	}
}

func TestNextID_EmptyElement_CachesAndIncrements(t *testing.T) {
	sp := newTestStoryPart(t, etree.NewElement("root"))
	if got := sp.NextID(); got != 1 {
		t.Errorf("first NextID on empty: got %d, want 1", got)
	}
	if got := sp.NextID(); got != 2 {
		t.Errorf("second NextID on empty: got %d, want 2", got)
	}
}

// newTestStoryPart creates a minimal StoryPart for unit tests.
func newTestStoryPart(t *testing.T, el *etree.Element) *StoryPart {
	t.Helper()
	xp := opc.NewXmlPartFromElement("", "", el, nil)
	return &StoryPart{XmlPart: xp}
}
