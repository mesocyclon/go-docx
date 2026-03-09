package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// Mirrors Python: it_provides_access_to_its_comments_collection
func TestCommentsPart_CommentsElement(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test">
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:comment>
</w:comments>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/comments.xml", opc.CTWmlComments, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	cp := NewCommentsPart(xp)

	comments, err := cp.CommentsElement()
	if err != nil {
		t.Fatalf("CommentsElement(): %v", err)
	}
	if comments == nil {
		t.Fatal("CommentsElement() returned nil")
	}
}

// Mirrors Python: it_constructs_a_default_comments_part_to_help
func TestCommentsPart_Default(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	cp, err := DefaultCommentsPart(pkg)
	if err != nil {
		t.Fatalf("DefaultCommentsPart: %v", err)
	}
	if cp == nil {
		t.Fatal("DefaultCommentsPart returned nil")
	}
	if cp.Element() == nil {
		t.Error("default CommentsPart has nil element")
	}
	comments, err := cp.CommentsElement()
	if err != nil {
		t.Fatalf("default CommentsElement(): %v", err)
	}
	if comments == nil {
		t.Fatal("default CommentsElement() returned nil")
	}
}

func TestLoadCommentsPart_ReturnsCommentsPart(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadCommentsPart("/word/comments.xml", opc.CTWmlComments, "", blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*CommentsPart); !ok {
		t.Errorf("LoadCommentsPart returned %T, want *CommentsPart", part)
	}
}
