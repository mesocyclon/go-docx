package docx

import (
	"strings"
	"testing"
	"time"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// -----------------------------------------------------------------------
// Test helpers
// -----------------------------------------------------------------------

func mustParseXml(t *testing.T, xml string) *oxml.Element {
	t.Helper()
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("ParseXml: %v", err)
	}
	return oxml.NewElement(el)
}

func makeP(t *testing.T, innerXml string) *oxml.CT_P {
	t.Helper()
	xml := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:p>`
	el := mustParseXml(t, xml)
	return &oxml.CT_P{Element: *el}
}

// testCommentsPart creates a minimal CommentsPart for testing.
func testCommentsPart(t *testing.T) *parts.CommentsPart {
	t.Helper()
	root := etree.NewElement("w:comments")
	root.CreateAttr("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
	xp := opc.NewXmlPartFromElement("/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", root, nil)
	return parts.NewCommentsPart(xp)
}

func makeR(t *testing.T, innerXml string) *oxml.CT_R {
	t.Helper()
	xml := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:r>`
	el := mustParseXml(t, xml)
	return &oxml.CT_R{Element: *el}
}

func makeTbl(t *testing.T, innerXml string) *oxml.CT_Tbl {
	t.Helper()
	xml := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + innerXml + `</w:tbl>`
	el := mustParseXml(t, xml)
	return &oxml.CT_Tbl{Element: *el}
}

// -----------------------------------------------------------------------
// shared.go — Length / RGBColor
// -----------------------------------------------------------------------

func TestLength_Conversions(t *testing.T) {
	tests := []struct {
		name string
		emu  int64
		pt   float64
		in   float64
	}{
		{"one_inch", 914400, 72.0, 1.0},
		{"one_point", 12700, 1.0, 1.0 / 72.0},
		{"zero", 0, 0.0, 0.0},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			l := Emu(tt.emu)
			if l.Emu() != tt.emu {
				t.Errorf("Emu() = %d, want %d", l.Emu(), tt.emu)
			}
			gotPt := l.Pt()
			if diff := gotPt - tt.pt; diff > 0.01 || diff < -0.01 {
				t.Errorf("Pt() = %f, want %f", gotPt, tt.pt)
			}
		})
	}
}

func TestRGBColor(t *testing.T) {
	t.Run("from_components", func(t *testing.T) {
		c := NewRGBColor(0x3C, 0x2F, 0x80)
		if c.String() != "3C2F80" {
			t.Errorf("String() = %q, want %q", c.String(), "3C2F80")
		}
		if c.R() != 0x3C || c.G() != 0x2F || c.B() != 0x80 {
			t.Errorf("components = (%d,%d,%d), want (60,47,128)", c.R(), c.G(), c.B())
		}
	})

	t.Run("from_string", func(t *testing.T) {
		c, err := RGBColorFromString("FF0000")
		if err != nil {
			t.Fatal(err)
		}
		if c.R() != 0xFF || c.G() != 0 || c.B() != 0 {
			t.Errorf("components = (%d,%d,%d), want (255,0,0)", c.R(), c.G(), c.B())
		}
	})

	t.Run("invalid_hex", func(t *testing.T) {
		_, err := RGBColorFromString("XYZ")
		if err == nil {
			t.Error("expected error for invalid hex")
		}
	})
}

// -----------------------------------------------------------------------
// font.go — Font tri-state booleans
// -----------------------------------------------------------------------

func TestFont_BoldTriState(t *testing.T) {
	tests := []struct {
		name   string
		xml    string
		expect *bool
	}{
		{
			"not_present",
			``,
			nil,
		},
		{
			"present_no_val",
			`<w:rPr><w:b/></w:rPr>`,
			boolPtr(true),
		},
		{
			"val_true",
			`<w:rPr><w:b w:val="true"/></w:rPr>`,
			boolPtr(true),
		},
		{
			"val_false",
			`<w:rPr><w:b w:val="false"/></w:rPr>`,
			boolPtr(false),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.xml)
			f := newFont(r)
			got := f.Bold()
			if got == nil && tt.expect == nil {
				return
			}
			if got == nil || tt.expect == nil || *got != *tt.expect {
				t.Errorf("Bold() = %v, want %v", got, tt.expect)
			}
		})
	}
}

func TestFont_SetBold(t *testing.T) {
	r := makeR(t, "")
	f := newFont(r)

	// Set bold to true
	if err := f.SetBold(boolPtr(true)); err != nil {
		t.Fatal(err)
	}
	got := f.Bold()
	if got == nil || !*got {
		t.Error("expected Bold()=true after SetBold(true)")
	}

	// Set bold to nil (remove)
	if err := f.SetBold(nil); err != nil {
		t.Fatal(err)
	}
	if f.Bold() != nil {
		t.Error("expected Bold()=nil after SetBold(nil)")
	}
}

func TestFont_Name(t *testing.T) {
	tests := []struct {
		name   string
		xml    string
		expect *string
	}{
		{
			"not_present",
			``,
			nil,
		},
		{
			"set",
			`<w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>`,
			strPtr("Arial"),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := makeR(t, tt.xml)
			f := newFont(r)
			got := f.Name()
			if got == nil && tt.expect == nil {
				return
			}
			if got == nil || tt.expect == nil || *got != *tt.expect {
				t.Errorf("Name() = %v, want %v", ptrStr(got), ptrStr(tt.expect))
			}
		})
	}
}

func TestFont_Size(t *testing.T) {
	// size 24 = 12pt in half-points
	r := makeR(t, `<w:rPr><w:sz w:val="24"/></w:rPr>`)
	f := newFont(r)
	got, err := f.Size()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil Size()")
	}
	// 12pt = 12 * 12700 EMU = 152400
	if got.Pt() < 11.99 || got.Pt() > 12.01 {
		t.Errorf("Size() = %f pt, want ~12.0", got.Pt())
	}
}

// -----------------------------------------------------------------------
// paragraph.go — Paragraph
// -----------------------------------------------------------------------

func TestParagraph_Text(t *testing.T) {
	p := makeP(t, `<w:r><w:t>Hello</w:t></w:r><w:r><w:t> World</w:t></w:r>`)
	para := newParagraph(p, nil)
	if got := para.Text(); got != "Hello World" {
		t.Errorf("Text() = %q, want %q", got, "Hello World")
	}
}

func TestParagraph_Runs(t *testing.T) {
	p := makeP(t, `<w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r>`)
	para := newParagraph(p, nil)
	runs := para.Runs()
	if len(runs) != 2 {
		t.Fatalf("len(Runs) = %d, want 2", len(runs))
	}
	if runs[0].Text() != "A" || runs[1].Text() != "B" {
		t.Errorf("run texts = %q, %q, want A, B", runs[0].Text(), runs[1].Text())
	}
}

func TestParagraph_AddRun(t *testing.T) {
	p := makeP(t, "")
	para := newParagraph(p, nil)
	r, err := para.AddRun("test")
	if err != nil {
		t.Fatal(err)
	}
	if r.Text() != "test" {
		t.Errorf("added run text = %q, want %q", r.Text(), "test")
	}
	if para.Text() != "test" {
		t.Errorf("paragraph text after add = %q, want %q", para.Text(), "test")
	}
}

// -----------------------------------------------------------------------
// parfmt.go — ParagraphFormat
// -----------------------------------------------------------------------

func TestParagraphFormat_Alignment(t *testing.T) {
	p := makeP(t, `<w:pPr><w:jc w:val="center"/></w:pPr>`)
	para := newParagraph(p, nil)
	pf := para.ParagraphFormat()

	got, err := pf.Alignment()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil alignment")
	}
	if *got != enum.WdParagraphAlignmentCenter {
		t.Errorf("Alignment() = %v, want CENTER", *got)
	}
}

func TestParagraphFormat_KeepWithNext(t *testing.T) {
	p := makeP(t, `<w:pPr><w:keepNext/></w:pPr>`)
	para := newParagraph(p, nil)
	pf := para.ParagraphFormat()

	got := pf.KeepWithNext()
	if got == nil || !*got {
		t.Error("expected KeepWithNext()=true")
	}
}

func TestParagraphFormat_SpaceAfter(t *testing.T) {
	// 240 twips
	p := makeP(t, `<w:pPr><w:spacing w:after="240"/></w:pPr>`)
	para := newParagraph(p, nil)
	pf := para.ParagraphFormat()

	got, err := pf.SpaceAfter()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil SpaceAfter")
	}
	if *got != 240 {
		t.Errorf("SpaceAfter() = %d, want 240", *got)
	}
}

func TestParagraphFormat_LineSpacing_Multiple(t *testing.T) {
	// w:line="480" w:lineRule="auto" → 480/240 = 2.0 (double spacing)
	p := makeP(t, `<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>`)
	para := newParagraph(p, nil)
	pf := para.ParagraphFormat()

	got, err := pf.LineSpacing()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("expected non-nil LineSpacing")
	}
	if !got.IsMultiple() {
		t.Fatalf("expected IsMultiple for MULTIPLE, got twips=%d", got.Twips())
	}
	if f := got.Multiple(); f < 1.99 || f > 2.01 {
		t.Errorf("LineSpacing().Multiple() = %f, want 2.0", f)
	}
}

func TestParagraphFormat_LineSpacingRule_Single(t *testing.T) {
	// 240 twips with MULTIPLE → SINGLE
	p := makeP(t, `<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>`)
	para := newParagraph(p, nil)
	pf := para.ParagraphFormat()

	rule, err := pf.LineSpacingRule()
	if err != nil {
		t.Fatal(err)
	}
	if rule == nil {
		t.Fatal("expected non-nil LineSpacingRule")
	}
	if *rule != enum.WdLineSpacingSingle {
		t.Errorf("LineSpacingRule() = %v, want SINGLE", *rule)
	}
}

// -----------------------------------------------------------------------
// table.go — Table
// -----------------------------------------------------------------------

func TestTable_Rows(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
		<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
	`)
	table := newTable(tbl, nil)

	rows := table.Rows()
	if rows.Len() != 2 {
		t.Errorf("Rows.Len() = %d, want 2", rows.Len())
	}
}

func TestTable_Columns(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
	`)
	table := newTable(tbl, nil)

	cols, err := table.Columns()
	if err != nil {
		t.Fatal(err)
	}
	if cols.Len() != 2 {
		t.Errorf("Columns.Len() = %d, want 2", cols.Len())
	}
}

func TestTable_CellAt(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
		<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
	`)
	table := newTable(tbl, nil)

	cell, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	if cell.Text() != "A1" {
		t.Errorf("CellAt(0,0) = %q, want %q", cell.Text(), "A1")
	}

	cell, err = table.CellAt(1, 1)
	if err != nil {
		t.Fatal(err)
	}
	if cell.Text() != "B2" {
		t.Errorf("CellAt(1,1) = %q, want %q", cell.Text(), "B2")
	}
}

func TestTable_CellAt_OutOfRange(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc><w:p/></w:tc></w:tr>
	`)
	table := newTable(tbl, nil)

	_, err := table.CellAt(5, 0)
	if err == nil {
		t.Error("expected error for out-of-range row")
	}
}

// -----------------------------------------------------------------------
// table.go — Cell merge (gridSpan + vMerge)
// -----------------------------------------------------------------------

func TestTable_HorizontalMerge(t *testing.T) {
	// 2-col table where first row has gridSpan="2"
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr>
			<w:tc>
				<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
				<w:p><w:r><w:t>merged</w:t></w:r></w:p>
			</w:tc>
		</w:tr>
	`)
	table := newTable(tbl, nil)

	// Both cell(0,0) and cell(0,1) should refer to the same cell
	c1, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	c2, err := table.CellAt(0, 1)
	if err != nil {
		t.Fatal(err)
	}
	if c1.Text() != "merged" || c2.Text() != "merged" {
		t.Errorf("merged cell text: c1=%q c2=%q, both should be 'merged'", c1.Text(), c2.Text())
	}
}

func TestTable_VerticalMerge(t *testing.T) {
	// 1-col, 2-row table where second row has <w:vMerge/> (no val = continue)
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr>
			<w:tc>
				<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
				<w:p><w:r><w:t>top</w:t></w:r></w:p>
			</w:tc>
		</w:tr>
		<w:tr>
			<w:tc>
				<w:tcPr><w:vMerge/></w:tcPr>
				<w:p/>
			</w:tc>
		</w:tr>
	`)
	table := newTable(tbl, nil)

	// Both rows should reference the same cell (the "top" cell)
	c1, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	c2, err := table.CellAt(1, 0)
	if err != nil {
		t.Fatal(err)
	}
	if c1.Text() != "top" {
		t.Errorf("CellAt(0,0) = %q, want %q", c1.Text(), "top")
	}
	if c2.Text() != "top" {
		t.Errorf("CellAt(1,0) = %q, want %q (should be same as top via vMerge continue)", c2.Text(), "top")
	}
}

// -----------------------------------------------------------------------
// section.go — Section
// -----------------------------------------------------------------------

func TestSection_PageDimensions(t *testing.T) {
	xml := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pgSz w:w="12240" w:h="15840"/>
		<w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>
	</w:sectPr>`
	el := mustParseXml(t, xml)
	sectPr := &oxml.CT_SectPr{Element: *el}
	sec := newSection(sectPr, nil)

	w, err := sec.PageWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w == nil || *w != 12240 {
		t.Errorf("PageWidth = %v, want 12240", w)
	}
	h, err := sec.PageHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h == nil || *h != 15840 {
		t.Errorf("PageHeight = %v, want 15840", h)
	}
}

func TestSection_Margins(t *testing.T) {
	xml := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>
	</w:sectPr>`
	el := mustParseXml(t, xml)
	sectPr := &oxml.CT_SectPr{Element: *el}
	sec := newSection(sectPr, nil)

	top, err := sec.TopMargin()
	if err != nil {
		t.Fatal(err)
	}
	if top == nil || *top != 1440 {
		t.Errorf("TopMargin = %v, want 1440", top)
	}
	left, err := sec.LeftMargin()
	if err != nil {
		t.Fatal(err)
	}
	if left == nil || *left != 1800 {
		t.Errorf("LeftMargin = %v, want 1800", left)
	}
}

// -----------------------------------------------------------------------
// styles.go — BabelFish
// -----------------------------------------------------------------------

func TestBabelFish_UI2Internal(t *testing.T) {
	tests := []struct {
		ui       string
		internal string
	}{
		{"Heading 1", "heading 1"},
		{"Caption", "caption"},
		{"Normal", "Normal"}, // not aliased
	}
	for _, tt := range tests {
		t.Run(tt.ui, func(t *testing.T) {
			got := UI2Internal(tt.ui)
			if got != tt.internal {
				t.Errorf("UI2Internal(%q) = %q, want %q", tt.ui, got, tt.internal)
			}
		})
	}
}

func TestBabelFish_Internal2UI(t *testing.T) {
	got := Internal2UI("heading 1")
	if got != "Heading 1" {
		t.Errorf("Internal2UI(\"heading 1\") = %q, want \"Heading 1\"", got)
	}
}

// -----------------------------------------------------------------------
// color.go — ColorFormat
// -----------------------------------------------------------------------

func TestColorFormat_RGB(t *testing.T) {
	r := makeR(t, `<w:rPr><w:color w:val="3C2F80"/></w:rPr>`)
	cf := newColorFormat(r)

	rgb, err := cf.RGB()
	if err != nil {
		t.Fatal(err)
	}
	if rgb == nil {
		t.Fatal("expected non-nil RGB")
	}
	if rgb.String() != "3C2F80" {
		t.Errorf("RGB() = %q, want %q", rgb.String(), "3C2F80")
	}
}

func TestColorFormat_Type_RGB(t *testing.T) {
	r := makeR(t, `<w:rPr><w:color w:val="FF0000"/></w:rPr>`)
	cf := newColorFormat(r)

	ct, err := cf.Type()
	if err != nil {
		t.Fatal(err)
	}
	if ct == nil {
		t.Fatal("expected non-nil Type")
	}
	if *ct != enum.MsoColorTypeRGB {
		t.Errorf("Type() = %d, want RGB(%d)", *ct, enum.MsoColorTypeRGB)
	}
}

func TestColorFormat_Type_Auto(t *testing.T) {
	r := makeR(t, `<w:rPr><w:color w:val="auto"/></w:rPr>`)
	cf := newColorFormat(r)

	ct2, err := cf.Type()
	if err != nil {
		t.Fatal(err)
	}
	if ct2 == nil {
		t.Fatal("expected non-nil Type for auto")
	}
	if *ct2 != enum.MsoColorTypeAuto {
		t.Errorf("Type() = %d, want AUTO(%d)", *ct2, enum.MsoColorTypeAuto)
	}
}

func TestColorFormat_Type_None(t *testing.T) {
	r := makeR(t, ``)
	cf := newColorFormat(r)

	ct3, err := cf.Type()
	if err != nil {
		t.Fatal(err)
	}
	if ct3 != nil {
		t.Error("expected nil Type for no rPr")
	}
}

// -----------------------------------------------------------------------
// settings.go — Settings
// -----------------------------------------------------------------------

func TestSettings_OddAndEvenPages(t *testing.T) {
	xml := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:evenAndOddHeaders/>
	</w:settings>`
	el := mustParseXml(t, xml)
	s := &oxml.CT_Settings{Element: *el}
	settings := newSettings(s)

	if !settings.OddAndEvenPagesHeaderFooter() {
		t.Error("expected OddAndEvenPagesHeaderFooter() = true")
	}
}

func TestSettings_OddAndEvenPages_NotPresent(t *testing.T) {
	xml := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el := mustParseXml(t, xml)
	s := &oxml.CT_Settings{Element: *el}
	settings := newSettings(s)

	if settings.OddAndEvenPagesHeaderFooter() {
		t.Error("expected OddAndEvenPagesHeaderFooter() = false when not present")
	}
}

// -----------------------------------------------------------------------
// comments.go — Comments / Comment
// -----------------------------------------------------------------------

func TestComment_Properties(t *testing.T) {
	xml := `<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		w:id="42" w:author="John" w:initials="JD" w:date="2024-01-15T10:30:00Z">
		<w:p><w:r><w:t>Test comment</w:t></w:r></w:p>
	</w:comment>`
	el := mustParseXml(t, xml)
	ce := &oxml.CT_Comment{Element: *el}
	c := newComment(ce, testCommentsPart(t))

	gotAuthor, err := c.Author()
	if err != nil {
		t.Fatal(err)
	}
	if gotAuthor != "John" {
		t.Errorf("Author() = %q, want %q", gotAuthor, "John")
	}
	if c.Initials() != "JD" {
		t.Errorf("Initials() = %q, want %q", c.Initials(), "JD")
	}
	if c.Text() != "Test comment" {
		t.Errorf("Text() = %q, want %q", c.Text(), "Test comment")
	}
	ts, err := c.Timestamp()
	if err != nil {
		t.Fatal(err)
	}
	if ts == nil {
		t.Fatal("expected non-nil Timestamp")
	}
	if ts.Year() != 2024 || ts.Month() != time.January || ts.Day() != 15 {
		t.Errorf("Timestamp = %v, want 2024-01-15", ts)
	}
	id, err := c.CommentID()
	if err != nil {
		t.Fatal(err)
	}
	if id != 42 {
		t.Errorf("CommentID() = %d, want 42", id)
	}
}

// -----------------------------------------------------------------------
// hyperlink.go — Hyperlink
// -----------------------------------------------------------------------

func TestHyperlink_Fragment(t *testing.T) {
	xml := `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		w:anchor="_Toc147925734">
		<w:r><w:t>See section 1</w:t></w:r>
	</w:hyperlink>`
	el := mustParseXml(t, xml)
	h := &oxml.CT_Hyperlink{Element: *el}
	hl := newHyperlink(h, nil)

	if hl.Fragment() != "_Toc147925734" {
		t.Errorf("Fragment() = %q, want %q", hl.Fragment(), "_Toc147925734")
	}
	if hl.Text() != "See section 1" {
		t.Errorf("Text() = %q, want %q", hl.Text(), "See section 1")
	}
	// No rId means address is ""
	if hl.Address() != "" {
		t.Errorf("Address() = %q, want empty for internal link", hl.Address())
	}
	// URL should be empty since address is empty
	if hl.URL() != "" {
		t.Errorf("URL() = %q, want empty", hl.URL())
	}
}

func TestHyperlink_Runs(t *testing.T) {
	xml := `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>word1</w:t></w:r>
		<w:r><w:t> word2</w:t></w:r>
	</w:hyperlink>`
	el := mustParseXml(t, xml)
	h := &oxml.CT_Hyperlink{Element: *el}
	hl := newHyperlink(h, nil)

	runs := hl.Runs()
	if len(runs) != 2 {
		t.Fatalf("len(Runs) = %d, want 2", len(runs))
	}
}

// -----------------------------------------------------------------------
// shape.go — InlineShapes
// -----------------------------------------------------------------------

func TestInlineShapes_Len(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
		<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/></wp:inline></w:drawing></w:r></w:p>
		<w:p><w:r><w:drawing><wp:inline><wp:extent cx="457200" cy="457200"/></wp:inline></w:drawing></w:r></w:p>
	</w:body>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	iss := newInlineShapes(el, nil)
	if iss.Len() != 2 {
		t.Errorf("Len() = %d, want 2", iss.Len())
	}
}

func TestInlineShapes_Empty(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>No shapes here</w:t></w:r></w:p>
	</w:body>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	iss := newInlineShapes(el, nil)
	if iss.Len() != 0 {
		t.Errorf("Len() = %d, want 0", iss.Len())
	}
}

// -----------------------------------------------------------------------
// tabstops.go — TabStops
// -----------------------------------------------------------------------

func TestTabStops_Len(t *testing.T) {
	xml := `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tabs>
			<w:tab w:val="left" w:pos="720"/>
			<w:tab w:val="center" w:pos="4680"/>
			<w:tab w:val="right" w:pos="9360"/>
		</w:tabs>
	</w:pPr>`
	el := mustParseXml(t, xml)
	pPr := &oxml.CT_PPr{Element: *el}
	ts := newTabStops(pPr)

	if ts.Len() != 3 {
		t.Errorf("Len() = %d, want 3", ts.Len())
	}
}

// -----------------------------------------------------------------------
// blkcntnr.go — BlockItemContainer
// -----------------------------------------------------------------------

func TestBlockItemContainer_Paragraphs(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>para1</w:t></w:r></w:p>
		<w:p><w:r><w:t>para2</w:t></w:r></w:p>
	</w:body>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	bic := newBlockItemContainer(el, nil)
	paras := bic.Paragraphs()
	if len(paras) != 2 {
		t.Fatalf("len(Paragraphs) = %d, want 2", len(paras))
	}
	if paras[0].Text() != "para1" {
		t.Errorf("para[0] = %q, want %q", paras[0].Text(), "para1")
	}
}

func TestBlockItemContainer_IterInnerContent(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p/>
		<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
		<w:p/>
	</w:body>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	bic := newBlockItemContainer(el, nil)
	items := bic.IterInnerContent()
	if len(items) != 3 {
		t.Fatalf("len(IterInnerContent) = %d, want 3", len(items))
	}
	if !items[0].IsParagraph() {
		t.Error("item[0] should be paragraph")
	}
	if !items[1].IsTable() {
		t.Error("item[1] should be table")
	}
	if !items[2].IsParagraph() {
		t.Error("item[2] should be paragraph")
	}
}

// -----------------------------------------------------------------------
// splitNewlines helper
// -----------------------------------------------------------------------

func TestSplitNewlines(t *testing.T) {
	tests := []struct {
		input    string
		expected []string
	}{
		{"", []string{""}},
		{"hello", []string{"hello"}},
		{"a\nb\nc", []string{"a", "b", "c"}},
		{"line1\nline2\n", []string{"line1", "line2", ""}},
		// Windows \r\n
		{"a\r\nb\r\nc", []string{"a", "b", "c"}},
		{"line1\r\nline2\r\n", []string{"line1", "line2", ""}},
		// Classic Mac \r
		{"a\rb\rc", []string{"a", "b", "c"}},
		{"line1\rline2\r", []string{"line1", "line2", ""}},
		// Mixed
		{"unix\nwin\r\nmac\rend", []string{"unix", "win", "mac", "end"}},
	}
	for _, tt := range tests {
		result := splitNewlines(tt.input)
		if len(result) != len(tt.expected) {
			t.Errorf("splitNewlines(%q) = %v (len %d), want len %d", tt.input, result, len(result), len(tt.expected))
			continue
		}
		for i := range result {
			if result[i] != tt.expected[i] {
				t.Errorf("splitNewlines(%q)[%d] = %q, want %q", tt.input, i, result[i], tt.expected[i])
			}
		}
	}
}

// -----------------------------------------------------------------------
// errIndexOutOfRange
// -----------------------------------------------------------------------

func TestErrIndexOutOfRange(t *testing.T) {
	err := errIndexOutOfRange("Test", 5, 3)
	if err == nil {
		t.Fatal("expected non-nil error")
	}
	if !strings.Contains(err.Error(), "Test") || !strings.Contains(err.Error(), "5") {
		t.Errorf("error message = %q, expected to contain collection name and index", err.Error())
	}
}

// -----------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------

func boolPtr(v bool) *bool    { return &v }
func strPtr(v string) *string { return &v }

func ptrStr(p *string) string {
	if p == nil {
		return "<nil>"
	}
	return *p
}
