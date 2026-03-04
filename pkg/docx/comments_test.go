package docx

import (
	"testing"
	"time"
)

// -----------------------------------------------------------------------
// comments_test.go — Comments / Comment (Batch 1)
// Mirrors Python: tests/test_comments.py
// -----------------------------------------------------------------------

func makeTestComments(t *testing.T, innerXml string) *Comments {
	t.Helper()
	elm := makeComments(t, innerXml)
	part := testCommentsPart(t)
	return newComments(elm, part)
}

// Mirrors Python: Comments.it_knows_how_many_comments
func TestComments_Len(t *testing.T) {
	tests := []struct {
		name     string
		innerXml string
		expected int
	}{
		{"empty", ``, 0},
		{"one", `<w:comment w:id="0" w:author="A"><w:p/></w:comment>`, 1},
		{"two", `<w:comment w:id="0" w:author="A"><w:p/></w:comment><w:comment w:id="1" w:author="B"><w:p/></w:comment>`, 2},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cs := makeTestComments(t, tt.innerXml)
			if cs.Len() != tt.expected {
				t.Errorf("Len() = %d, want %d", cs.Len(), tt.expected)
			}
		})
	}
}

// Mirrors Python: Comments.it_is_iterable
func TestComments_Iter(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="0" w:author="A"><w:p/></w:comment><w:comment w:id="1" w:author="B"><w:p/></w:comment>`)
	items := cs.Iter()
	if len(items) != 2 {
		t.Fatalf("len(Iter()) = %d, want 2", len(items))
	}
	a0, err := items[0].Author()
	if err != nil {
		t.Fatal(err)
	}
	if a0 != "A" {
		t.Errorf("Iter()[0].Author() = %q, want %q", a0, "A")
	}
	a1, err := items[1].Author()
	if err != nil {
		t.Fatal(err)
	}
	if a1 != "B" {
		t.Errorf("Iter()[1].Author() = %q, want %q", a1, "B")
	}
}

// Mirrors Python: Comments.it_can_get_by_id
func TestComments_Get(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="42" w:author="John"><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:comment>`)

	// Found
	c := cs.Get(42)
	if c == nil {
		t.Fatal("Get(42) returned nil")
	}
	got, err := c.Author()
	if err != nil {
		t.Fatal(err)
	}
	if got != "John" {
		t.Errorf("Author() = %q, want %q", got, "John")
	}

	// Not found
	c2 := cs.Get(999)
	if c2 != nil {
		t.Error("Get(999) should return nil")
	}
}

// Mirrors Python: Comments.it_can_add_a_new_comment
func TestComments_AddComment(t *testing.T) {
	cs := makeTestComments(t, ``)

	c, err := cs.AddComment("Test text", "Author", strPtr("AA"))
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	if c == nil {
		t.Fatal("AddComment returned nil")
	}
	gotA, err := c.Author()
	if err != nil {
		t.Fatal(err)
	}
	if gotA != "Author" {
		t.Errorf("Author() = %q, want %q", gotA, "Author")
	}
	if c.Initials() != "AA" {
		t.Errorf("Initials() = %q, want %q", c.Initials(), "AA")
	}
	if c.Text() != "Test text" {
		t.Errorf("Text() = %q, want %q", c.Text(), "Test text")
	}
	if cs.Len() != 1 {
		t.Errorf("Len() after add = %d, want 1", cs.Len())
	}
}

// Mirrors Python: Comment.it_can_update_author
func TestComment_SetAuthor(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="0" w:author="Old"><w:p/></w:comment>`)
	c := cs.Iter()[0]

	if err := c.SetAuthor("New"); err != nil {
		t.Fatal(err)
	}
	gotN, err := c.Author()
	if err != nil {
		t.Fatal(err)
	}
	if gotN != "New" {
		t.Errorf("Author() = %q, want %q", gotN, "New")
	}
}

// Mirrors Python: Comment.it_can_update_initials
func TestComment_SetInitials(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="0" w:author="A" w:initials="OI"><w:p/></w:comment>`)
	c := cs.Iter()[0]

	if err := c.SetInitials("NI"); err != nil {
		t.Fatal(err)
	}
	if c.Initials() != "NI" {
		t.Errorf("Initials() = %q, want %q", c.Initials(), "NI")
	}
}

// Mirrors Python: Comment.it_knows_the_datetime
func TestComment_Timestamp(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="0" w:author="A" w:date="2024-01-15T10:30:00Z"><w:p/></w:comment>`)
	c := cs.Iter()[0]

	ts, err := c.Timestamp()
	if err != nil {
		t.Fatal(err)
	}
	if ts == nil {
		t.Fatal("Timestamp() returned nil")
	}
	if ts.Year() != 2024 || ts.Month() != time.January || ts.Day() != 15 {
		t.Errorf("Timestamp = %v, want 2024-01-15", ts)
	}
}

// Mirrors Python: Comment.it_can_summarize_as_text (multi-paragraph)
func TestComment_Text_MultiParagraph(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="0" w:author="A">`+
		`<w:p><w:r><w:t>Line 1</w:t></w:r></w:p>`+
		`<w:p><w:r><w:t>Line 2</w:t></w:r></w:p>`+
		`</w:comment>`)
	c := cs.Iter()[0]
	if c.Text() != "Line 1\nLine 2" {
		t.Errorf("Text() = %q, want %q", c.Text(), "Line 1\nLine 2")
	}
}

// Mirrors Python: Comment.it_provides_access_to_paragraphs
func TestComment_Paragraphs(t *testing.T) {
	cs := makeTestComments(t, `<w:comment w:id="0" w:author="A">`+
		`<w:p><w:r><w:t>para 1</w:t></w:r></w:p>`+
		`<w:p><w:r><w:t>para 2</w:t></w:r></w:p>`+
		`</w:comment>`)
	c := cs.Iter()[0]
	paras := c.Paragraphs()
	if len(paras) != 2 {
		t.Fatalf("len(Paragraphs) = %d, want 2", len(paras))
	}
	if paras[0].Text() != "para 1" {
		t.Errorf("para[0] = %q, want %q", paras[0].Text(), "para 1")
	}
}
