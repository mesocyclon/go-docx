package parts

import (
	"bytes"
	goimage "image"
	"image/png"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
)

// newDocPartWithBody creates a minimal DocumentPart with <w:body/> and
// attached OPC package. Optionally pre-wires extra relationship parts.
func newDocPartWithBody(t *testing.T) (*DocumentPart, *opc.OpcPackage) {
	t.Helper()
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)
	pkg := opc.NewOpcPackage(NewDocxPartFactory())
	xp, err := opc.NewXmlPart("/word/document.xml", opc.CTWmlDocumentMain, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	dp := NewDocumentPart(xp)
	pkg.AddPart(dp)
	pkg.RelateTo(dp, opc.RTOfficeDocument)
	return dp, pkg
}

// wireNumberingPart creates and wires a NumberingPart to dp.
func wireNumberingPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *NumberingPart {
	t.Helper()
	np, err := DefaultNumberingPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(np)
	dp.Rels().GetOrAdd(opc.RTNumbering, np)
	return np
}

// wireFootnotesPart creates and wires a FootnotesPart to dp.
func wireFootnotesPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *FootnotesPart {
	t.Helper()
	fp, err := DefaultFootnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(fp)
	dp.Rels().GetOrAdd(opc.RTFootnotes, fp)
	return fp
}

// wireEndnotesPart creates and wires an EndnotesPart to dp.
func wireEndnotesPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *EndnotesPart {
	t.Helper()
	ep, err := DefaultEndnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(ep)
	dp.Rels().GetOrAdd(opc.RTEndnotes, ep)
	return ep
}

// wireCommentsPart creates and wires a CommentsPart to dp.
func wireCommentsPart(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage) *CommentsPart {
	t.Helper()
	cp, err := DefaultCommentsPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	pkg.AddPart(cp)
	dp.Rels().GetOrAdd(opc.RTComments, cp)
	return cp
}

// stylesXML returns styles XML with a single paragraph style.
func stylesXML(styleID, styleName string, isDefault bool) []byte {
	def := ""
	if isDefault {
		def = ` w:default="1"`
	}
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="` + styleID + `"` + def + `>
    <w:name w:val="` + styleName + `"/>
  </w:style>
</w:styles>`)
}

// wireStylesPartFromXML wires a custom styles part with given XML.
func wireStylesPartFromXML(t *testing.T, dp *DocumentPart, pkg *opc.OpcPackage, xml []byte) *StylesPart {
	t.Helper()
	xp, err := opc.NewXmlPart("/word/styles.xml", opc.CTWmlStyles, xml, pkg)
	if err != nil {
		t.Fatal(err)
	}
	sp := NewStylesPart(xp)
	pkg.AddPart(sp)
	dp.Rels().GetOrAdd(opc.RTStyles, sp)
	return sp
}

// minimumPNG returns a valid 1x1 PNG blob.
func minimumPNG() []byte {
	img := goimage.NewRGBA(goimage.Rect(0, 0, 1, 1))
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		panic(err)
	}
	return buf.Bytes()
}

// mockStyledObject implements styledObject interface for testing.
type mockStyledObject struct {
	styleID   string
	styleType enum.WdStyleType
	typeErr   error
}

func (m *mockStyledObject) StyleID() string                 { return m.styleID }
func (m *mockStyledObject) Type() (enum.WdStyleType, error) { return m.styleType, m.typeErr }
