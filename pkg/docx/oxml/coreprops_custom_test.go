package oxml

import (
	"testing"
	"time"
)

func TestNewCoreProperties(t *testing.T) {
	cp, err := NewCoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	if cp == nil {
		t.Fatal("expected coreProperties, got nil")
	}
	// Check that it has the cp namespace
	_, ok := HasNsDecl(cp.e, "cp")
	if !ok {
		t.Error("expected xmlns:cp declaration")
	}
}

func TestCT_CoreProperties_TextProperties(t *testing.T) {
	cp, err := NewCoreProperties()
	if err != nil {
		t.Fatal(err)
	}

	// All text properties should start empty
	if got := cp.TitleText(); got != "" {
		t.Errorf("expected empty title, got %q", got)
	}

	// Set and get title
	if err := cp.SetTitleText("My Document"); err != nil {
		t.Fatal(err)
	}
	if got := cp.TitleText(); got != "My Document" {
		t.Errorf("expected 'My Document', got %q", got)
	}

	// Set and get author
	if err := cp.SetAuthorText("Alice"); err != nil {
		t.Fatal(err)
	}
	if got := cp.AuthorText(); got != "Alice" {
		t.Errorf("expected 'Alice', got %q", got)
	}

	// Set and get subject
	if err := cp.SetSubjectText("Testing"); err != nil {
		t.Fatal(err)
	}
	if got := cp.SubjectText(); got != "Testing" {
		t.Errorf("expected 'Testing', got %q", got)
	}

	// Set and get category
	if err := cp.SetCategoryText("Test Category"); err != nil {
		t.Fatal(err)
	}
	if got := cp.CategoryText(); got != "Test Category" {
		t.Errorf("expected 'Test Category', got %q", got)
	}

	// Set and get keywords
	if err := cp.SetKeywordsText("go, docx"); err != nil {
		t.Fatal(err)
	}
	if got := cp.KeywordsText(); got != "go, docx" {
		t.Errorf("expected 'go, docx', got %q", got)
	}

	// Set and get comments
	if err := cp.SetCommentsText("A test"); err != nil {
		t.Fatal(err)
	}
	if got := cp.CommentsText(); got != "A test" {
		t.Errorf("expected 'A test', got %q", got)
	}

	// Set and get lastModifiedBy
	if err := cp.SetLastModifiedByText("Bob"); err != nil {
		t.Fatal(err)
	}
	if got := cp.LastModifiedByText(); got != "Bob" {
		t.Errorf("expected 'Bob', got %q", got)
	}

	// Set and get contentStatus
	if err := cp.SetContentStatusText("Draft"); err != nil {
		t.Fatal(err)
	}
	if got := cp.ContentStatusText(); got != "Draft" {
		t.Errorf("expected 'Draft', got %q", got)
	}

	// Set and get language
	if err := cp.SetLanguageText("en-US"); err != nil {
		t.Fatal(err)
	}
	if got := cp.LanguageText(); got != "en-US" {
		t.Errorf("expected 'en-US', got %q", got)
	}

	// Set and get version
	if err := cp.SetVersionText("1.0"); err != nil {
		t.Fatal(err)
	}
	if got := cp.VersionText(); got != "1.0" {
		t.Errorf("expected '1.0', got %q", got)
	}
}

func TestCT_CoreProperties_TextProperty_255Limit(t *testing.T) {
	cp, err := NewCoreProperties()
	if err != nil {
		t.Fatal(err)
	}
	longStr := ""
	for i := 0; i < 256; i++ {
		longStr += "x"
	}
	err = cp.SetTitleText(longStr)
	if err == nil {
		t.Error("expected error for string > 255 chars")
	}
}

func TestCT_CoreProperties_DatetimeProperties(t *testing.T) {
	cp, err := NewCoreProperties()
	if err != nil {
		t.Fatal(err)
	}

	// Created should be nil initially
	gotInit, errInit := cp.CreatedDatetime()
	if errInit != nil {
		t.Fatal(errInit)
	}
	if gotInit != nil {
		t.Errorf("expected nil created, got %v", gotInit)
	}

	// Set created
	created := time.Date(2024, 1, 15, 10, 30, 0, 0, time.UTC)
	cp.SetCreatedDatetime(created)
	got, err := cp.CreatedDatetime()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil created")
	}
	if !got.Equal(created) {
		t.Errorf("expected %v, got %v", created, *got)
	}

	// Set modified
	modified := time.Date(2024, 6, 20, 14, 45, 0, 0, time.UTC)
	cp.SetModifiedDatetime(modified)
	got, err = cp.ModifiedDatetime()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil modified")
	}
	if !got.Equal(modified) {
		t.Errorf("expected %v, got %v", modified, *got)
	}

	// Set lastPrinted
	lastPrinted := time.Date(2024, 3, 1, 8, 0, 0, 0, time.UTC)
	cp.SetLastPrintedDatetime(lastPrinted)
	got, err = cp.LastPrintedDatetime()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil lastPrinted")
	}
	if !got.Equal(lastPrinted) {
		t.Errorf("expected %v, got %v", lastPrinted, *got)
	}
}

func TestCT_CoreProperties_RevisionNumber(t *testing.T) {
	cp, err := NewCoreProperties()
	if err != nil {
		t.Fatal(err)
	}

	// Default should be 0
	if got := cp.RevisionNumber(); got != 0 {
		t.Errorf("expected revision 0, got %d", got)
	}

	// Set valid revision
	if err := cp.SetRevisionNumber(5); err != nil {
		t.Fatal(err)
	}
	if got := cp.RevisionNumber(); got != 5 {
		t.Errorf("expected revision 5, got %d", got)
	}

	// Set invalid (< 1) should error
	if err := cp.SetRevisionNumber(0); err == nil {
		t.Error("expected error for revision 0")
	}
	if err := cp.SetRevisionNumber(-1); err == nil {
		t.Error("expected error for negative revision")
	}
}

func TestCT_CoreProperties_RevisionFromXml(t *testing.T) {
	// Test parsing a revision from existing XML
	xml := `<cp:coreProperties ` +
		`xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
		`xmlns:dcterms="http://purl.org/dc/terms/">` +
		`<cp:revision>42</cp:revision>` +
		`</cp:coreProperties>`
	el, _ := ParseXml([]byte(xml))
	cp := &CT_CoreProperties{Element{e: el}}

	if got := cp.RevisionNumber(); got != 42 {
		t.Errorf("expected revision 42, got %d", got)
	}

	// Non-integer revision should return 0
	xml2 := `<cp:coreProperties ` +
		`xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
		`xmlns:dcterms="http://purl.org/dc/terms/">` +
		`<cp:revision>abc</cp:revision>` +
		`</cp:coreProperties>`
	el2, _ := ParseXml([]byte(xml2))
	cp2 := &CT_CoreProperties{Element{e: el2}}
	if got := cp2.RevisionNumber(); got != 0 {
		t.Errorf("expected 0 for non-integer revision, got %d", got)
	}
}

func TestParseW3CDTF(t *testing.T) {
	tests := []struct {
		input   string
		wantErr bool
		year    int
		month   time.Month
		day     int
	}{
		{"2024-01-15T10:30:00Z", false, 2024, time.January, 15},
		{"2024-01-15T10:30:00", false, 2024, time.January, 15},
		{"2024-01-15", false, 2024, time.January, 15},
		{"2024-01", false, 2024, time.January, 1},
		{"2024", false, 2024, time.January, 1},
		{"", true, 0, 0, 0},
		{"not-a-date", true, 0, 0, 0},
	}
	for _, tt := range tests {
		dt, err := parseW3CDTF(tt.input)
		if tt.wantErr {
			if err == nil {
				t.Errorf("parseW3CDTF(%q): expected error, got nil", tt.input)
			}
			continue
		}
		if err != nil {
			t.Errorf("parseW3CDTF(%q): unexpected error: %v", tt.input, err)
			continue
		}
		if dt.Year() != tt.year || dt.Month() != tt.month || dt.Day() != tt.day {
			t.Errorf("parseW3CDTF(%q): expected %d-%d-%d, got %d-%d-%d",
				tt.input, tt.year, tt.month, tt.day, dt.Year(), dt.Month(), dt.Day())
		}
	}
}

func TestParseW3CDTF_WithOffset(t *testing.T) {
	dt, err := parseW3CDTF("2024-01-15T10:30:00-05:00")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// -05:00 means UTC+5 hours when reversed (add +5 to get UTC)
	if dt.Hour() != 15 || dt.Minute() != 30 {
		t.Errorf("expected 15:30 UTC, got %d:%d", dt.Hour(), dt.Minute())
	}
}
