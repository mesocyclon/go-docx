package docx

import (
	"testing"
)

func TestBody_ClearContent(t *testing.T) {
	doc := mustNewDoc(t)
	doc.AddParagraph("To be cleared")
	doc.AddTable(1, 1)

	body, err := doc.getBody()
	if err != nil {
		t.Fatalf("getBody error: %v", err)
	}

	body.ClearContent()

	// After clear, paragraphs and tables should be gone,
	// but document should still be valid (sectPr preserved).
	paras := body.Paragraphs()
	if len(paras) != 0 {
		t.Errorf("expected 0 paragraphs after clear, got %d", len(paras))
	}
	tables := body.Tables()
	if len(tables) != 0 {
		t.Errorf("expected 0 tables after clear, got %d", len(tables))
	}
}

func TestBody_AddParagraphAndTable(t *testing.T) {
	doc := mustNewDoc(t)

	body, err := doc.getBody()
	if err != nil {
		t.Fatalf("getBody error: %v", err)
	}

	body.ClearContent()

	p, err := body.AddParagraph("Test")
	if err != nil {
		t.Fatalf("AddParagraph error: %v", err)
	}
	if p.Text() != "Test" {
		t.Errorf("paragraph text: want %q, got %q", "Test", p.Text())
	}

	tbl, err := body.AddTable(2, 2, 5000)
	if err != nil {
		t.Fatalf("AddTable error: %v", err)
	}
	if tbl == nil {
		t.Fatal("AddTable returned nil")
	}

	if len(body.Paragraphs()) != 1 {
		t.Errorf("paragraphs: want 1, got %d", len(body.Paragraphs()))
	}
	if len(body.Tables()) != 1 {
		t.Errorf("tables: want 1, got %d", len(body.Tables()))
	}
}
