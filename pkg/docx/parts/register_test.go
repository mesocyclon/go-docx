package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

func TestDocxPartFactory_DocumentPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	// Minimal valid XML for a document part
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>`)

	part, err := f.New(
		opc.PackURI("/word/document.xml"),
		opc.CTWmlDocumentMain,
		opc.RTOfficeDocument,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*DocumentPart); !ok {
		t.Errorf("factory returned %T, want *DocumentPart", part)
	}
}

func TestDocxPartFactory_StylesPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)

	part, err := f.New(
		opc.PackURI("/word/styles.xml"),
		opc.CTWmlStyles,
		opc.RTStyles,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*StylesPart); !ok {
		t.Errorf("factory returned %T, want *StylesPart", part)
	}
}

func TestDocxPartFactory_SettingsPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)

	part, err := f.New(
		opc.PackURI("/word/settings.xml"),
		opc.CTWmlSettings,
		opc.RTSettings,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*SettingsPart); !ok {
		t.Errorf("factory returned %T, want *SettingsPart", part)
	}
}

func TestDocxPartFactory_HeaderPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
</w:hdr>`)

	part, err := f.New(
		opc.PackURI("/word/header1.xml"),
		opc.CTWmlHeader,
		opc.RTHeader,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*HeaderPart); !ok {
		t.Errorf("factory returned %T, want *HeaderPart", part)
	}
}

func TestDocxPartFactory_FooterPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
</w:ftr>`)

	part, err := f.New(
		opc.PackURI("/word/footer1.xml"),
		opc.CTWmlFooter,
		opc.RTFooter,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*FooterPart); !ok {
		t.Errorf("factory returned %T, want *FooterPart", part)
	}
}

func TestDocxPartFactory_CommentsPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)

	part, err := f.New(
		opc.PackURI("/word/comments.xml"),
		opc.CTWmlComments,
		opc.RTComments,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*CommentsPart); !ok {
		t.Errorf("factory returned %T, want *CommentsPart", part)
	}
}

func TestDocxPartFactory_NumberingPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)

	part, err := f.New(
		opc.PackURI("/word/numbering.xml"),
		opc.CTWmlNumbering,
		opc.RTNumbering,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*NumberingPart); !ok {
		t.Errorf("factory returned %T, want *NumberingPart", part)
	}
}

func TestDocxPartFactory_ImagePart_Selector(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	// Image parts are selected by reltype + content-type prefix
	blob := []byte{0x89, 0x50, 0x4E, 0x47} // fake PNG bytes
	part, err := f.New(
		opc.PackURI("/word/media/image1.png"),
		opc.CTPng,
		opc.RTImage,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*ImagePart); !ok {
		t.Errorf("factory returned %T, want *ImagePart", part)
	}
}

func TestDocxPartFactory_ImagePart_JPEG(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte{0xFF, 0xD8} // fake JPEG bytes
	part, err := f.New(
		opc.PackURI("/word/media/image1.jpg"),
		opc.CTJpeg,
		opc.RTImage,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*ImagePart); !ok {
		t.Errorf("factory returned %T, want *ImagePart", part)
	}
}

func TestDocxPartFactory_FootnotesPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
</w:footnotes>`)

	part, err := f.New(
		opc.PackURI("/word/footnotes.xml"),
		opc.CTWmlFootnotes,
		opc.RTFootnotes,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*FootnotesPart); !ok {
		t.Errorf("factory returned %T, want *FootnotesPart", part)
	}
}

func TestDocxPartFactory_EndnotesPart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
</w:endnotes>`)

	part, err := f.New(
		opc.PackURI("/word/endnotes.xml"),
		opc.CTWmlEndnotes,
		opc.RTEndnotes,
		blob,
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*EndnotesPart); !ok {
		t.Errorf("factory returned %T, want *EndnotesPart", part)
	}
}

func TestDocxPartFactory_UnknownPart_FallsBackToBasePart(t *testing.T) {
	f := NewDocxPartFactory()
	pkg := opc.NewOpcPackage(nil)

	part, err := f.New(
		opc.PackURI("/word/theme/theme1.xml"),
		"application/vnd.openxmlformats-officedocument.theme+xml",
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
		[]byte("<a:theme/>"),
		pkg,
	)
	if err != nil {
		t.Fatal(err)
	}
	// Should fall back to BasePart
	if _, ok := part.(*opc.BasePart); !ok {
		t.Errorf("factory returned %T for unknown content type, want *opc.BasePart", part)
	}
}
