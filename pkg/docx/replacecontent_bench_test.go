package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// --------------------------------------------------------------------------
// Benchmarks for ReplaceWithContent (Step 8.7)
// --------------------------------------------------------------------------

func BenchmarkRWC_Simple(b *testing.B) {
	source := benchSource(b)
	for b.Loop() {
		target := benchTarget(b)
		target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	}
}

func BenchmarkRWC_WithImages(b *testing.B) {
	source := benchSourceWithImage(b)
	for b.Loop() {
		target := benchTarget(b)
		target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	}
}

func BenchmarkRWC_100Replacements(b *testing.B) {
	source := benchSource(b)
	for b.Loop() {
		target := benchTargetN(b, 100)
		target.ReplaceWithContent("[<TAG>]", ContentData{Source: source})
	}
}

func BenchmarkRWC_KeepSourceFormatting(b *testing.B) {
	source := benchSourceWithConflictingStyle(b)
	for b.Loop() {
		target := benchTargetWithStyle(b)
		target.ReplaceWithContent("[<TAG>]", ContentData{
			Source: source,
			Format: KeepSourceFormatting,
		})
	}
}

func BenchmarkRWC_Snapshot(b *testing.B) {
	for b.Loop() {
		doc := benchDocWith10Paragraphs(b)
		snap, _ := doc.snapshotBody()
		doc.restoreBody(snap)
	}
}

// --------------------------------------------------------------------------
// Benchmark helpers
// --------------------------------------------------------------------------

func benchTarget(b *testing.B) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	doc.AddParagraph("[<TAG>]")
	return doc
}

func benchTargetN(b *testing.B, n int) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	for i := 0; i < n; i++ {
		doc.AddParagraph("[<TAG>]")
	}
	return doc
}

func benchSource(b *testing.B) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	doc.AddParagraph("Source paragraph text")
	return doc
}

func benchSourceWithImage(b *testing.B) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	imgBlob := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	imgPart := parts.NewImagePartWithMeta(
		"/word/media/image1.png", "image/png", imgBlob,
		100, 100, 72, 72, "bench.png",
	)
	doc.wmlPkg.OpcPackage.AddPart(imgPart)
	rel := doc.part.Rels().GetOrAdd(opc.RTImage, imgPart)

	body, _ := doc.getBody()
	pEl, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<w:r><w:drawing><a:blip r:embed="` + rel.RID + `" ` +
			`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></w:drawing></w:r></w:p>`,
	))
	body.insertBeforeSectPr(pEl)
	return doc
}

func benchSourceWithConflictingStyle(b *testing.B) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	srcStyles, _ := doc.part.Styles()
	srcXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="BenchStyle">` +
		`<w:name w:val="Bench Style"/><w:pPr><w:jc w:val="center"/></w:pPr>` +
		`<w:rPr><w:b/></w:rPr></w:style>`
	srcEl, _ := oxml.ParseXml([]byte(srcXml))
	srcStyles.RawElement().AddChild(srcEl)

	srcBody := doc.element.Body().RawElement()
	p, _ := oxml.ParseXml([]byte(
		`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:pPr><w:pStyle w:val="BenchStyle"/></w:pPr>` +
			`<w:r><w:t>Styled</w:t></w:r></w:p>`,
	))
	children := srcBody.ChildElements()
	for i, child := range children {
		if child.Space == "w" && child.Tag == "sectPr" {
			srcBody.InsertChildAt(i, p)
			break
		}
	}
	return doc
}

func benchTargetWithStyle(b *testing.B) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	doc.AddParagraph("[<TAG>]")
	tgtStyles, _ := doc.part.Styles()
	tgtXml := `<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:type="paragraph" w:styleId="BenchStyle">` +
		`<w:name w:val="Bench Style"/><w:pPr><w:jc w:val="left"/></w:pPr>` +
		`</w:style>`
	tgtEl, _ := oxml.ParseXml([]byte(tgtXml))
	tgtStyles.RawElement().AddChild(tgtEl)
	return doc
}

func benchDocWith10Paragraphs(b *testing.B) *Document {
	b.Helper()
	doc, err := New()
	if err != nil {
		b.Fatal(err)
	}
	for i := 0; i < 10; i++ {
		doc.AddParagraph("Paragraph text for benchmarking")
	}
	return doc
}
