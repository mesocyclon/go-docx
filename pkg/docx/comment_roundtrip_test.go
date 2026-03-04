package docx

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

// -----------------------------------------------------------------------
// Comment round-trip tests (MR-13)
// -----------------------------------------------------------------------

func TestAddComment_RoundTrip_Single(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("Annotated text")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("flagged")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	initials := "JD"
	comment, err := doc.AddComment([]*Run{run}, "Review this", "John Doe", &initials)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	if comment == nil {
		t.Fatal("AddComment returned nil")
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	if comments.Len() != 1 {
		t.Fatalf("expected 1 comment after round-trip, got %d", comments.Len())
	}
}

func TestAddComment_RoundTrip_MetadataPreserved(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("Test paragraph")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("some text")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	initials := "AB"
	_, err = doc.AddComment([]*Run{run}, "Important note", "Alice Brown", &initials)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	all := comments.Iter()
	if len(all) != 1 {
		t.Fatalf("expected 1 comment, got %d", len(all))
	}

	c := all[0]
	gotA, err := c.Author()
	if err != nil {
		t.Fatal(err)
	}
	if gotA != "Alice Brown" {
		t.Errorf("Author() = %q, want %q", gotA, "Alice Brown")
	}
	if c.Initials() != "AB" {
		t.Errorf("Initials() = %q, want %q", c.Initials(), "AB")
	}
	if !strings.Contains(c.Text(), "Important note") {
		t.Errorf("Text() = %q, expected to contain %q", c.Text(), "Important note")
	}
	cid, err := c.CommentID()
	if err != nil {
		t.Fatalf("CommentID(): %v", err)
	}
	if cid < 0 {
		t.Errorf("CommentID() = %d, expected >= 0", cid)
	}
}

func TestAddComment_RoundTrip_MultipleComments(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run1, err := p.AddRun("first part")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}
	run2, err := p.AddRun("second part")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	_, err = doc.AddComment([]*Run{run1}, "Comment A", "Author1", nil)
	if err != nil {
		t.Fatalf("AddComment(1): %v", err)
	}
	_, err = doc.AddComment([]*Run{run2}, "Comment B", "Author2", nil)
	if err != nil {
		t.Fatalf("AddComment(2): %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	if comments.Len() != 2 {
		t.Errorf("expected 2 comments, got %d", comments.Len())
	}
}

func TestAddComment_RoundTrip_MultiRunRange(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	r1, err := p.AddRun("word1 ")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}
	_, err = p.AddRun("word2 ")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}
	r3, err := p.AddRun("word3")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	_, err = doc.AddComment([]*Run{r1, r3}, "Multi-run comment", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	if comments.Len() != 1 {
		t.Errorf("expected 1 comment, got %d", comments.Len())
	}
}

func TestAddComment_RoundTrip_XMLRangeMarkers(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("annotated")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	comment, err := doc.AddComment([]*Run{run}, "Test", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	commentID, err := comment.CommentID()
	if err != nil {
		t.Fatalf("CommentID: %v", err)
	}
	idStr := fmt.Sprintf("%d", commentID)

	pEl := p.CT_P().RawElement()
	var foundStart, foundEnd, foundRef bool
	for _, child := range pEl.ChildElements() {
		switch child.Tag {
		case "commentRangeStart":
			for _, attr := range child.Attr {
				if attr.Key == "id" && attr.Value == idStr {
					foundStart = true
				}
			}
		case "commentRangeEnd":
			for _, attr := range child.Attr {
				if attr.Key == "id" && attr.Value == idStr {
					foundEnd = true
				}
			}
		case "r":
			for _, grandChild := range child.ChildElements() {
				if grandChild.Tag == "commentReference" {
					for _, attr := range grandChild.Attr {
						if attr.Key == "id" && attr.Value == idStr {
							foundRef = true
						}
					}
				}
			}
		}
	}
	if !foundStart {
		t.Error("commentRangeStart not found in paragraph XML")
	}
	if !foundEnd {
		t.Error("commentRangeEnd not found in paragraph XML")
	}
	if !foundRef {
		t.Error("commentReference not found in paragraph XML")
	}
}

func TestAddComment_MultilineText(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("text")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	_, err = doc.AddComment([]*Run{run}, "Line 1\nLine 2\nLine 3", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	all := comments.Iter()
	if len(all) != 1 {
		t.Fatalf("expected 1 comment, got %d", len(all))
	}
	text := all[0].Text()
	if !strings.Contains(text, "Line 1") ||
		!strings.Contains(text, "Line 2") ||
		!strings.Contains(text, "Line 3") {
		t.Errorf("comment text = %q, expected all three lines", text)
	}
}

func TestAddComment_EmptyRuns_Error(t *testing.T) {
	doc := mustNewDoc(t)
	_, err := doc.AddComment(nil, "text", "author", nil)
	if err == nil {
		t.Error("expected error for nil runs, got nil")
	}
	_, err = doc.AddComment([]*Run{}, "text", "author", nil)
	if err == nil {
		t.Error("expected error for empty runs slice, got nil")
	}
}

func TestAddComment_EmptyText(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("content")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	comment, err := doc.AddComment([]*Run{run}, "", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	if comment == nil {
		t.Fatal("AddComment returned nil")
	}
}

func TestComment_Get_ByID(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("text")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	comment, err := doc.AddComment([]*Run{run}, "Found me", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	cid, err := comment.CommentID()
	if err != nil {
		t.Fatalf("CommentID: %v", err)
	}

	comments, err := doc.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}

	found := comments.Get(cid)
	if found == nil {
		t.Fatalf("Comments.Get(%d) returned nil", cid)
	}
	if !strings.Contains(found.Text(), "Found me") {
		t.Errorf("Get(%d).Text() = %q, expected to contain %q", cid, found.Text(), "Found me")
	}

	if comments.Get(99999) != nil {
		t.Error("expected nil for non-existent comment ID")
	}
}

func TestComment_Timestamp_RoundTrip(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("text")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	comment, err := doc.AddComment([]*Run{run}, "Timestamped", "Author", nil)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}
	ts, err := comment.Timestamp()
	if err != nil {
		t.Fatal(err)
	}
	if ts == nil {
		t.Fatal("expected non-nil Timestamp on new comment")
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	all := comments.Iter()
	if len(all) != 1 {
		t.Fatalf("expected 1 comment, got %d", len(all))
	}
	ts2, err := all[0].Timestamp()
	if err != nil {
		t.Fatal(err)
	}
	if ts2 == nil {
		t.Fatal("expected non-nil Timestamp after round-trip")
	}
	diff := ts2.Sub(*ts)
	if diff < 0 {
		diff = -diff
	}
	if diff.Seconds() > 5 {
		t.Errorf("timestamps differ by %v", diff)
	}
}

// -----------------------------------------------------------------------
// ReplaceText reaches comments (MR-13 / findRunForAtom fix)
// -----------------------------------------------------------------------

func TestDocument_ReplaceText_ReachesComments(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("body text")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("annotated")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	initials := "T"
	_, err = doc.AddComment([]*Run{run}, "COMMENT_OLD is here", "Tester", &initials)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}

	// Round-trip: save → reopen → replace → save → reopen.
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	n, err := doc2.ReplaceText("COMMENT_OLD", "COMMENT_NEW")
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if n != 1 {
		t.Errorf("expected 1 replacement in comment, got %d", n)
	}

	// Verify the comment text was actually changed.
	comments, err := doc2.Comments()
	if err != nil {
		t.Fatalf("Comments(): %v", err)
	}
	if comments.Len() != 1 {
		t.Fatalf("expected 1 comment, got %d", comments.Len())
	}
	text := comments.Iter()[0].Text()
	if !strings.Contains(text, "COMMENT_NEW") {
		t.Errorf("comment text should contain 'COMMENT_NEW', got %q", text)
	}
	if strings.Contains(text, "COMMENT_OLD") {
		t.Errorf("comment text should not contain 'COMMENT_OLD', got %q", text)
	}
}

func TestDocument_ReplaceText_NoCommentsPart_NoPanic(t *testing.T) {
	doc := mustNewDoc(t)
	if _, err := doc.AddParagraph("no comments here"); err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}

	// Round-trip without ever creating comments.
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	// Must not panic or error when comments part is absent.
	n, err := doc2.ReplaceText("anything", "else")
	if err != nil {
		t.Fatalf("ReplaceText on doc without comments: %v", err)
	}
	if n != 0 {
		t.Errorf("expected 0 replacements, got %d", n)
	}
}

func TestDocument_ReplaceText_CommentMultiParagraph(t *testing.T) {
	doc := mustNewDoc(t)
	p, err := doc.AddParagraph("text")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	run, err := p.AddRun("marked")
	if err != nil {
		t.Fatalf("AddRun: %v", err)
	}

	// Multi-line comment: "PLACEHOLDER in line one\nPLACEHOLDER in line two"
	_, err = doc.AddComment(
		[]*Run{run},
		"PLACEHOLDER in line one\nPLACEHOLDER in line two",
		"Author", nil,
	)
	if err != nil {
		t.Fatalf("AddComment: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	n, err := doc2.ReplaceText("PLACEHOLDER", "DONE")
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	// Two paragraphs, each with one occurrence.
	if n != 2 {
		t.Errorf("expected 2 replacements in multi-paragraph comment, got %d", n)
	}

	comments, _ := doc2.Comments()
	text := comments.Iter()[0].Text()
	if text != "DONE in line one\nDONE in line two" {
		t.Errorf("unexpected comment text: %q", text)
	}
}
