package oxml

import (
	"testing"
)

func TestCT_Comments_AddCommentFull(t *testing.T) {
	xml := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>`
	el, _ := ParseXml([]byte(xml))
	cs := &CT_Comments{Element{e: el}}

	c, err := cs.AddCommentFull()
	if err != nil {
		t.Fatal(err)
	}
	if c == nil {
		t.Fatal("expected comment, got nil")
	}
	id, err := c.Id()
	if err != nil {
		t.Fatalf("comment id error: %v", err)
	}
	if id != 0 {
		t.Errorf("expected first comment id=0, got %d", id)
	}

	// Add another
	c2, err := cs.AddCommentFull()
	if err != nil {
		t.Fatal(err)
	}
	id2, _ := c2.Id()
	if id2 != 1 {
		t.Errorf("expected second comment id=1, got %d", id2)
	}

	// Check list
	if len(cs.CommentList()) != 2 {
		t.Errorf("expected 2 comments, got %d", len(cs.CommentList()))
	}
}

func TestCT_Comments_GetCommentByID(t *testing.T) {
	xml := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:comment w:id="5" w:author="Alice"/>` +
		`<w:comment w:id="10" w:author="Bob"/>` +
		`</w:comments>`
	el, _ := ParseXml([]byte(xml))
	cs := &CT_Comments{Element{e: el}}

	c := cs.GetCommentByID(10)
	if c == nil {
		t.Fatal("expected comment with id=10, got nil")
	}
	author, _ := c.Author()
	if author != "Bob" {
		t.Errorf("expected author 'Bob', got %q", author)
	}

	if cs.GetCommentByID(999) != nil {
		t.Error("expected nil for nonexistent comment id")
	}
}

func TestCT_Comment_InnerContentElements(t *testing.T) {
	xml := `<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="0" w:author="">` +
		`<w:p/><w:tbl/><w:p/>` +
		`</w:comment>`
	el, _ := ParseXml([]byte(xml))
	c := &CT_Comment{Element{e: el}}

	elems := c.InnerContentElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elems))
	}
	if _, ok := elems[0].(*CT_P); !ok {
		t.Error("expected first element to be CT_P")
	}
	if _, ok := elems[1].(*CT_Tbl); !ok {
		t.Error("expected second element to be CT_Tbl")
	}
	if _, ok := elems[2].(*CT_P); !ok {
		t.Error("expected third element to be CT_P")
	}
}
