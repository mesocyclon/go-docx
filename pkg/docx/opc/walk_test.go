package opc

import (
	"fmt"
	"testing"
)

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

// makePart creates a named BasePart registered in pkg.
func makePart(pkg *OpcPackage, name PackURI) *BasePart {
	p := NewBasePart(name, CTXml, nil, pkg)
	pkg.AddPart(p)
	return p
}

// link adds an internal relationship from source's rels to target.
func link(src *Relationships, relType string, target Part) {
	ref := target.PartName().RelativeRef(src.BaseURI())
	src.Add(relType, ref, target, false)
}

// linkExt adds an external relationship on src.
func linkExt(src *Relationships, relType, url string) {
	src.Add(relType, url, nil, true)
}

// partNames extracts ordered partname strings from a Part slice.
func partNames(parts []Part) []string {
	out := make([]string, len(parts))
	for i, p := range parts {
		out[i] = string(p.PartName())
	}
	return out
}

// assertPartNames checks that actual partnames match expected (order matters).
func assertPartNames(t *testing.T, got []Part, want []string) {
	t.Helper()
	names := partNames(got)
	if len(names) != len(want) {
		t.Fatalf("len: got %d %v, want %d %v", len(names), names, len(want), want)
	}
	for i := range want {
		if names[i] != want[i] {
			t.Errorf("[%d]: got %q, want %q (full: %v)", i, names[i], want[i], names)
			return
		}
	}
}

// assertRelCount checks that the rels slice has the expected length.
func assertRelCount(t *testing.T, rels []*Relationship, want int) {
	t.Helper()
	if len(rels) != want {
		t.Errorf("rel count: got %d, want %d", len(rels), want)
	}
}

// ---------------------------------------------------------------------------
// IterParts
// ---------------------------------------------------------------------------

func TestIterParts_Empty(t *testing.T) {
	t.Parallel()
	pkg := NewOpcPackage(nil)
	parts := pkg.IterParts()
	if len(parts) != 0 {
		t.Errorf("expected 0 parts, got %d", len(parts))
	}
}

func TestIterParts_SinglePart(t *testing.T) {
	t.Parallel()
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	link(pkg.Rels(), RTOfficeDocument, a)

	assertPartNames(t, pkg.IterParts(), []string{"/a.xml"})
}

func TestIterParts_LinearChain(t *testing.T) {
	t.Parallel()
	// pkg → A → B → C  (DFS order: A, B, C)
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")
	c := makePart(pkg, "/c.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(b.Rels(), RTStyles, c)

	assertPartNames(t, pkg.IterParts(), []string{"/a.xml", "/b.xml", "/c.xml"})
}

func TestIterParts_BranchingTree(t *testing.T) {
	t.Parallel()
	// pkg → A; A → B, A → C  (DFS order: A, B, C)
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")
	c := makePart(pkg, "/c.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(a.Rels(), RTStyles, c)

	assertPartNames(t, pkg.IterParts(), []string{"/a.xml", "/b.xml", "/c.xml"})
}

func TestIterParts_Diamond_NoDuplicates(t *testing.T) {
	t.Parallel()
	// pkg → A; A → B, A → C; B → C  (C must appear only once)
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")
	c := makePart(pkg, "/c.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(a.Rels(), RTStyles, c)
	link(b.Rels(), RTStyles, c) // duplicate path to C

	parts := pkg.IterParts()
	assertPartNames(t, parts, []string{"/a.xml", "/b.xml", "/c.xml"})
}

func TestIterParts_Cycle(t *testing.T) {
	t.Parallel()
	// pkg → A → B → A  (cycle — each part appears once)
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(b.Rels(), RTStyles, a) // back-edge

	assertPartNames(t, pkg.IterParts(), []string{"/a.xml", "/b.xml"})
}

func TestIterParts_SkipsExternal(t *testing.T) {
	t.Parallel()
	// pkg → A (internal), pkg → ext (external)
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	linkExt(pkg.Rels(), RTOfficeDocument, "https://example.com")

	assertPartNames(t, pkg.IterParts(), []string{"/a.xml"})
}

func TestIterParts_MultipleRootsFromPackage(t *testing.T) {
	t.Parallel()
	// pkg → A, pkg → B; A and B are independent
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(pkg.Rels(), RTStyles, b)

	assertPartNames(t, pkg.IterParts(), []string{"/a.xml", "/b.xml"})
}

func TestIterParts_DeepChain_NoStackOverflow(t *testing.T) {
	t.Parallel()
	// Build a linear chain of depth 5000:
	// pkg → p0 → p1 → ... → p4999
	// This would blow up with default recursion limit in Python (~1000)
	// and is the primary motivation for the iterative rewrite.
	const depth = 5000
	pkg := NewOpcPackage(nil)

	parts := make([]*BasePart, depth)
	for i := 0; i < depth; i++ {
		name := PackURI(fmt.Sprintf("/p%d.xml", i))
		parts[i] = makePart(pkg, name)
	}

	link(pkg.Rels(), RTOfficeDocument, parts[0])
	for i := 0; i < depth-1; i++ {
		link(parts[i].Rels(), RTStyles, parts[i+1])
	}

	got := pkg.IterParts()
	if len(got) != depth {
		t.Fatalf("expected %d parts, got %d", depth, len(got))
	}
	// Verify order: p0, p1, ..., p4999
	if string(got[0].PartName()) != "/p0.xml" {
		t.Errorf("first part: got %q", got[0].PartName())
	}
	if string(got[depth-1].PartName()) != fmt.Sprintf("/p%d.xml", depth-1) {
		t.Errorf("last part: got %q", got[depth-1].PartName())
	}
}

// ---------------------------------------------------------------------------
// IterRels
// ---------------------------------------------------------------------------

func TestIterRels_Empty(t *testing.T) {
	t.Parallel()
	pkg := NewOpcPackage(nil)
	assertRelCount(t, pkg.IterRels(), 0)
}

func TestIterRels_IncludesExternalRels(t *testing.T) {
	t.Parallel()
	// External rels are yielded but not descended into.
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	linkExt(pkg.Rels(), RTOfficeDocument, "https://example.com")

	rels := pkg.IterRels()
	// 2 from package level + 0 from a (a has no rels)
	assertRelCount(t, rels, 2)
}

func TestIterRels_LinearChain(t *testing.T) {
	t.Parallel()
	// pkg→A→B: 1 rel at package level, 1 rel from A
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)

	rels := pkg.IterRels()
	assertRelCount(t, rels, 2)
	if rels[0].TargetPart != a {
		t.Error("first rel should point to A")
	}
	if rels[1].TargetPart != b {
		t.Error("second rel should point to B")
	}
}

func TestIterRels_Diamond_NoDuplicateRels(t *testing.T) {
	t.Parallel()
	// pkg→A; A→B, A→C; B→C
	// Total rels: 1(pkg) + 2(A) + 1(B) = 4
	// C's rels (empty) not re-visited because C already visited when B→C is followed
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")
	c := makePart(pkg, "/c.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(a.Rels(), RTStyles, c)
	link(b.Rels(), RTStyles, c) // B→C rel is yielded but C is not descended into again

	rels := pkg.IterRels()
	assertRelCount(t, rels, 4)
}

func TestIterRels_Cycle(t *testing.T) {
	t.Parallel()
	// pkg→A→B→A: rels from A and B are each yielded once
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(b.Rels(), RTStyles, a)

	rels := pkg.IterRels()
	// 1(pkg→A) + 1(A→B) + 1(B→A, yielded but A already visited)
	assertRelCount(t, rels, 3)
}

func TestIterRels_DeepChain_NoStackOverflow(t *testing.T) {
	t.Parallel()
	const depth = 5000
	pkg := NewOpcPackage(nil)

	parts := make([]*BasePart, depth)
	for i := 0; i < depth; i++ {
		parts[i] = makePart(pkg, PackURI(fmt.Sprintf("/p%d.xml", i)))
	}

	link(pkg.Rels(), RTOfficeDocument, parts[0])
	for i := 0; i < depth-1; i++ {
		link(parts[i].Rels(), RTStyles, parts[i+1])
	}

	rels := pkg.IterRels()
	// 1 from package + (depth-1) inter-part links = depth total
	assertRelCount(t, rels, depth)
}

// ---------------------------------------------------------------------------
// DFS order preservation (IterParts must match original recursive order)
// ---------------------------------------------------------------------------

func TestIterParts_DFSOrder(t *testing.T) {
	t.Parallel()
	// Build a tree and verify pre-order DFS:
	//
	//   pkg → A
	//         ├→ B
	//         │  └→ D
	//         └→ C
	//            └→ E
	//
	// Expected DFS pre-order: A, B, D, C, E
	pkg := NewOpcPackage(nil)
	a := makePart(pkg, "/a.xml")
	b := makePart(pkg, "/b.xml")
	c := makePart(pkg, "/c.xml")
	d := makePart(pkg, "/d.xml")
	e := makePart(pkg, "/e.xml")

	link(pkg.Rels(), RTOfficeDocument, a)
	link(a.Rels(), RTStyles, b)
	link(a.Rels(), RTStyles, c)
	link(b.Rels(), RTStyles, d)
	link(c.Rels(), RTStyles, e)

	assertPartNames(t, pkg.IterParts(), []string{
		"/a.xml", "/b.xml", "/d.xml", "/c.xml", "/e.xml",
	})
}
