package docx

import (
	"testing"
	"time"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// ===========================================================================
// Fix #1: AddTable / addP insertion order (before w:sectPr)
// ===========================================================================

// TestAddTable_InsertsBeforeSectPr verifies that AddTable places the new w:tbl
// before any trailing w:sectPr, matching Python's _insert_tbl successor logic.
func TestAddTable_InsertsBeforeSectPr(t *testing.T) {
	// Build a <w:body> with one <w:p> and a trailing <w:sectPr>
	body := etree.NewElement("body")
	body.Space = "w"

	p := body.CreateElement("p")
	p.Space = "w"

	sectPr := body.CreateElement("sectPr")
	sectPr.Space = "w"

	bic := newBlockItemContainer(body, nil)

	// Add a table — should go BEFORE sectPr
	_, err := bic.AddTable(1, 1, 5000)
	if err != nil {
		t.Fatalf("AddTable failed: %v", err)
	}

	// Verify order: p, tbl, sectPr
	children := body.ChildElements()
	if len(children) != 3 {
		t.Fatalf("expected 3 children, got %d", len(children))
	}

	tags := make([]string, len(children))
	for i, c := range children {
		tags[i] = c.Tag
	}

	if tags[0] != "p" || tags[1] != "tbl" || tags[2] != "sectPr" {
		t.Errorf("expected [p, tbl, sectPr], got %v", tags)
	}
}

// TestAddParagraph_InsertsBeforeSectPr verifies that AddParagraph places
// the new w:p before any trailing w:sectPr.
func TestAddParagraph_InsertsBeforeSectPr(t *testing.T) {
	body := etree.NewElement("body")
	body.Space = "w"

	sectPr := body.CreateElement("sectPr")
	sectPr.Space = "w"

	bic := newBlockItemContainer(body, nil)

	_, err := bic.AddParagraph("hello")
	if err != nil {
		t.Fatalf("AddParagraph failed: %v", err)
	}

	children := body.ChildElements()
	if len(children) != 2 {
		t.Fatalf("expected 2 children, got %d", len(children))
	}

	if children[0].Tag != "p" {
		t.Errorf("expected first child to be <w:p>, got <%s:%s>", children[0].Space, children[0].Tag)
	}
	if children[1].Tag != "sectPr" {
		t.Errorf("expected second child to be <w:sectPr>, got <%s:%s>", children[1].Space, children[1].Tag)
	}
}

// TestAddTable_NoSectPr_Appends verifies that without a sectPr, elements
// are simply appended (regression check for non-body containers like cells).
func TestAddTable_NoSectPr_Appends(t *testing.T) {
	tc := etree.NewElement("tc")
	tc.Space = "w"

	p := tc.CreateElement("p")
	p.Space = "w"

	bic := newBlockItemContainer(tc, nil)

	_, err := bic.AddTable(1, 1, 3000)
	if err != nil {
		t.Fatalf("AddTable failed: %v", err)
	}

	children := tc.ChildElements()
	if len(children) != 2 {
		t.Fatalf("expected 2 children, got %d", len(children))
	}

	if children[0].Tag != "p" || children[1].Tag != "tbl" {
		t.Errorf("expected [p, tbl], got [%s, %s]", children[0].Tag, children[1].Tag)
	}
}

// TestMultipleInserts_BeforeSectPr verifies that interleaved paragraphs and
// tables all stay before sectPr and maintain their relative order.
func TestMultipleInserts_BeforeSectPr(t *testing.T) {
	body := etree.NewElement("body")
	body.Space = "w"

	sectPr := body.CreateElement("sectPr")
	sectPr.Space = "w"

	bic := newBlockItemContainer(body, nil)

	bic.AddParagraph("first")
	bic.AddTable(1, 1, 5000)
	bic.AddParagraph("second")

	children := body.ChildElements()
	if len(children) != 4 {
		t.Fatalf("expected 4 children, got %d", len(children))
	}

	expected := []string{"p", "tbl", "p", "sectPr"}
	for i, exp := range expected {
		if children[i].Tag != exp {
			t.Errorf("child[%d]: expected %s, got %s", i, exp, children[i].Tag)
		}
	}
}

// ===========================================================================
// Fix #2: CoreProperties domain proxy
// ===========================================================================

// makeCorePropsXML builds a minimal core properties XML for testing.
func makeCorePropsXML(author, title string) []byte {
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<cp:coreProperties ` +
		`xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
		`xmlns:dcterms="http://purl.org/dc/terms/" ` +
		`xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
	if author != "" {
		xml += `<dc:creator>` + author + `</dc:creator>`
	}
	if title != "" {
		xml += `<dc:title>` + title + `</dc:title>`
	}
	xml += `</cp:coreProperties>`
	return []byte(xml)
}

// makeCorePropsPartForTest creates a *parts.CorePropertiesPart from XML bytes,
// matching how the part factory loads core properties in production.
func makeCorePropsPartForTest(blob []byte) *parts.CorePropertiesPart {
	xp, err := opc.NewXmlPart("/docProps/core.xml", opc.CTOpcCoreProperties, blob, nil)
	if err != nil {
		panic("test helper: " + err.Error())
	}
	return parts.NewCorePropertiesPart(xp)
}

// ctFromPart extracts CT_CoreProperties from a CorePropertiesPart.
// Mirrors Python CorePropertiesPart.core_properties → CoreProperties(self.element).
func ctFromPart(cpp *parts.CorePropertiesPart) *oxml.CT_CoreProperties {
	el := cpp.Element()
	return &oxml.CT_CoreProperties{Element: oxml.WrapElement(el)}
}

func TestCoreProperties_ReadFromPart(t *testing.T) {
	cpp := makeCorePropsPartForTest(makeCorePropsXML("Jane", "Report"))
	cp := newCoreProperties(ctFromPart(cpp))

	if got := cp.Author(); got != "Jane" {
		t.Errorf("Author: expected 'Jane', got '%s'", got)
	}
	if got := cp.Title(); got != "Report" {
		t.Errorf("Title: expected 'Report', got '%s'", got)
	}
}

func TestCoreProperties_SettersAndGetters(t *testing.T) {
	cpp := makeCorePropsPartForTest(makeCorePropsXML("", ""))
	cp := newCoreProperties(ctFromPart(cpp))

	tests := []struct {
		name  string
		set   func(string) error
		get   func() string
		value string
	}{
		{"Author", cp.SetAuthor, cp.Author, "Alice"},
		{"Title", cp.SetTitle, cp.Title, "Test Title"},
		{"Subject", cp.SetSubject, cp.Subject, "Test Subject"},
		{"Category", cp.SetCategory, cp.Category, "Test Category"},
		{"Keywords", cp.SetKeywords, cp.Keywords, "go, docx, test"},
		{"Comments", cp.SetComments, cp.Comments, "A test document"},
		{"LastModifiedBy", cp.SetLastModifiedBy, cp.LastModifiedBy, "Bob"},
		{"ContentStatus", cp.SetContentStatus, cp.ContentStatus, "Draft"},
		{"Identifier", cp.SetIdentifier, cp.Identifier, "ID-123"},
		{"Language", cp.SetLanguage, cp.Language, "en-US"},
		{"Version", cp.SetVersion, cp.Version, "1.0.0"},
	}

	for _, tt := range tests {
		if err := tt.set(tt.value); err != nil {
			t.Errorf("%s: set error: %v", tt.name, err)
			continue
		}
		if got := tt.get(); got != tt.value {
			t.Errorf("%s: expected '%s', got '%s'", tt.name, tt.value, got)
		}
	}
}

func TestCoreProperties_SharedElement_WriteBack(t *testing.T) {
	cpp := makeCorePropsPartForTest(makeCorePropsXML("Original", ""))
	cp := newCoreProperties(ctFromPart(cpp))

	// Modify via proxy
	if err := cp.SetAuthor("Modified"); err != nil {
		t.Fatalf("SetAuthor: %v", err)
	}

	// Verify shared element: part's element should reflect the change
	el := cpp.Element()
	for _, child := range el.ChildElements() {
		if child.Tag == "creator" {
			if got := child.Text(); got != "Modified" {
				t.Errorf("Part element not updated: expected 'Modified', got '%s'", got)
			}
			return
		}
	}
	t.Error("dc:creator element not found in CorePropertiesPart after SetAuthor")
}

func TestCoreProperties_DatetimeProperties(t *testing.T) {
	cpp := makeCorePropsPartForTest(makeCorePropsXML("", ""))
	cp := newCoreProperties(ctFromPart(cpp))

	cre, err := cp.Created()
	if err != nil {
		t.Fatal(err)
	}
	if cre != nil {
		t.Error("Created should be nil initially")
	}
	mod, err := cp.Modified()
	if err != nil {
		t.Fatal(err)
	}
	if mod != nil {
		t.Error("Modified should be nil initially")
	}

	now := time.Date(2025, 6, 15, 10, 30, 0, 0, time.UTC)
	cp.SetCreated(now)
	cp.SetModified(now)

	gotCre, err := cp.Created()
	if err != nil {
		t.Fatal(err)
	}
	if gotCre == nil {
		t.Fatal("Created should not be nil after set")
	} else if !gotCre.Equal(now) {
		t.Errorf("Created: expected %v, got %v", now, *gotCre)
	}

	gotMod, err := cp.Modified()
	if err != nil {
		t.Fatal(err)
	}
	if gotMod == nil {
		t.Fatal("Modified should not be nil after set")
	} else if !gotMod.Equal(now) {
		t.Errorf("Modified: expected %v, got %v", now, *gotMod)
	}
}

func TestCoreProperties_Revision(t *testing.T) {
	cpp := makeCorePropsPartForTest(makeCorePropsXML("", ""))
	cp := newCoreProperties(ctFromPart(cpp))

	if got := cp.Revision(); got != 0 {
		t.Errorf("Revision: expected 0, got %d", got)
	}

	if err := cp.SetRevision(5); err != nil {
		t.Fatalf("SetRevision: %v", err)
	}
	if got := cp.Revision(); got != 5 {
		t.Errorf("Revision: expected 5, got %d", got)
	}

	if err := cp.SetRevision(0); err == nil {
		t.Error("SetRevision(0) should return error")
	}
}

func TestCoreProperties_255CharLimit(t *testing.T) {
	cpp := makeCorePropsPartForTest(makeCorePropsXML("", ""))
	cp := newCoreProperties(ctFromPart(cpp))

	long := ""
	for i := 0; i < 256; i++ {
		long += "x"
	}

	if err := cp.SetTitle(long); err == nil {
		t.Error("SetTitle with 256 chars should return error")
	}
	if err := cp.SetTitle(long[:255]); err != nil {
		t.Errorf("SetTitle with 255 chars should succeed, got: %v", err)
	}
}

func TestDefaultCorePropertiesPart(t *testing.T) {
	cpp, err := parts.DefaultCorePropertiesPart(nil)
	if err != nil {
		t.Fatalf("DefaultCorePropertiesPart: %v", err)
	}

	cp := newCoreProperties(ctFromPart(cpp))

	if got := cp.Title(); got != "Word Document" {
		t.Errorf("Title: expected 'Word Document', got '%s'", got)
	}
	if got := cp.LastModifiedBy(); got != "go-docx" {
		t.Errorf("LastModifiedBy: expected 'go-docx', got '%s'", got)
	}
	if got := cp.Revision(); got != 1 {
		t.Errorf("Revision: expected 1, got %d", got)
	}
	modCheck, err := cp.Modified()
	if err != nil {
		t.Fatal(err)
	}
	if modCheck == nil {
		t.Error("Modified should be set in default")
	}
}
