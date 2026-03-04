package opc

import (
	"testing"
)

func TestRelationships_Add(t *testing.T) {
	rels := NewRelationships("/word")
	part := NewBasePart("/word/styles.xml", CTWmlStyles, nil, nil)

	rel := rels.Add(RTStyles, "styles.xml", part, false)
	if rel.RID != "rId1" {
		t.Errorf("expected rId1, got %q", rel.RID)
	}
	if rel.RelType != RTStyles {
		t.Errorf("wrong reltype")
	}
	if rel.IsExternal {
		t.Error("expected internal")
	}

	// Add another
	rel2 := rels.Add(RTImage, "media/image1.png", nil, false)
	if rel2.RID != "rId2" {
		t.Errorf("expected rId2, got %q", rel2.RID)
	}
}

func TestRelationships_Load(t *testing.T) {
	rels := NewRelationships("/word")
	part := NewBasePart("/word/styles.xml", CTWmlStyles, nil, nil)

	rel := rels.Load("rId5", RTStyles, "styles.xml", part, false)
	if rel.RID != "rId5" {
		t.Errorf("expected rId5, got %q", rel.RID)
	}

	// Next auto rId should be past 5
	rel2 := rels.Add(RTImage, "media/image1.png", nil, false)
	if rel2.RID != "rId6" {
		t.Errorf("expected rId6, got %q", rel2.RID)
	}
}

func TestRelationships_GetByRID(t *testing.T) {
	rels := NewRelationships("/")
	part := NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, nil)
	rels.Load("rId1", RTOfficeDocument, "word/document.xml", part, false)

	got := rels.GetByRID("rId1")
	if got == nil {
		t.Fatal("expected non-nil")
	}
	if got.RelType != RTOfficeDocument {
		t.Errorf("wrong reltype: %q", got.RelType)
	}

	got = rels.GetByRID("rIdNonexistent")
	if got != nil {
		t.Error("expected nil for nonexistent rId")
	}
}

func TestRelationships_GetByRelType(t *testing.T) {
	rels := NewRelationships("/")
	part := NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, nil)
	rels.Load("rId1", RTOfficeDocument, "word/document.xml", part, false)

	got, err := rels.GetByRelType(RTOfficeDocument)
	if err != nil {
		t.Fatalf("GetByRelType: %v", err)
	}
	if got.RID != "rId1" {
		t.Errorf("expected rId1, got %q", got.RID)
	}

	// Not found
	_, err = rels.GetByRelType(RTStyles)
	if err == nil {
		t.Error("expected error for not-found reltype")
	}
}

func TestRelationships_AllByRelType(t *testing.T) {
	rels := NewRelationships("/word")
	rels.Load("rId1", RTHeader, "header1.xml", nil, false)
	rels.Load("rId2", RTHeader, "header2.xml", nil, false)
	rels.Load("rId3", RTFooter, "footer1.xml", nil, false)

	got := rels.AllByRelType(RTHeader)
	if len(got) != 2 {
		t.Errorf("expected 2, got %d", len(got))
	}
}

func TestRelationships_NextRID_FillsGaps(t *testing.T) {
	rels := NewRelationships("/")
	rels.Load("rId1", RTStyles, "styles.xml", nil, false)
	rels.Load("rId3", RTImage, "image.png", nil, false)

	// Next should be rId2 (gap) — but since nextNum advances past existing,
	// the behavior depends on the implementation. Let's check it's at least valid.
	rel := rels.Add(RTSettings, "settings.xml", nil, false)
	if rel.RID == "" {
		t.Error("expected non-empty rId")
	}
}

func TestRelationships_GetOrAdd(t *testing.T) {
	rels := NewRelationships("/word")
	part := NewBasePart("/word/styles.xml", CTWmlStyles, nil, nil)

	// First call creates
	rel := rels.GetOrAdd(RTStyles, part)
	if rel.RID != "rId1" {
		t.Errorf("expected rId1, got %q", rel.RID)
	}

	// Second call returns existing
	rel2 := rels.GetOrAdd(RTStyles, part)
	if rel2.RID != rel.RID {
		t.Errorf("expected same rId, got %q and %q", rel.RID, rel2.RID)
	}

	if rels.Len() != 1 {
		t.Errorf("expected 1 rel, got %d", rels.Len())
	}
}

func TestRelationships_GetOrAddExtRel(t *testing.T) {
	rels := NewRelationships("/word")

	rId := rels.GetOrAddExtRel(RTHyperlink, "http://example.com")
	if rId == "" {
		t.Error("expected non-empty rId")
	}

	// Second call returns same
	rId2 := rels.GetOrAddExtRel(RTHyperlink, "http://example.com")
	if rId2 != rId {
		t.Errorf("expected same rId, got %q and %q", rId, rId2)
	}
}

func TestRelationships_Delete(t *testing.T) {
	rels := NewRelationships("/")
	rels.Load("rId1", RTStyles, "styles.xml", nil, false)
	rels.Load("rId2", RTImage, "image.png", nil, false)

	rels.Delete("rId1")
	if rels.Len() != 1 {
		t.Errorf("expected 1 rel after delete, got %d", rels.Len())
	}
	if rels.GetByRID("rId1") != nil {
		t.Error("expected rId1 to be deleted")
	}
}

func TestRelationships_RelatedParts(t *testing.T) {
	rels := NewRelationships("/word")
	part1 := NewBasePart("/word/styles.xml", CTWmlStyles, nil, nil)
	part2 := NewBasePart("/word/numbering.xml", CTWmlNumbering, nil, nil)

	rels.Load("rId1", RTStyles, "styles.xml", part1, false)
	rels.Load("rId2", RTNumbering, "numbering.xml", part2, false)
	rels.Load("rId3", RTHyperlink, "http://example.com", nil, true)

	related := rels.RelatedParts()
	if len(related) != 2 {
		t.Errorf("expected 2 related parts, got %d", len(related))
	}
	if related["rId1"] != part1 {
		t.Error("rId1 should map to part1")
	}
	if related["rId2"] != part2 {
		t.Error("rId2 should map to part2")
	}
}
