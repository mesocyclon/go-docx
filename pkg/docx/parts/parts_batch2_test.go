package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// -----------------------------------------------------------------------
// parts_batch2_test.go — StylesPart, SettingsPart, NumberingPart,
//                         HeaderPart, FooterPart, CommentsPart (Batch 2)
// Mirrors Python: tests/parts/test_styles.py, test_settings.py,
//                 test_numbering.py, test_hdrftr.py, test_comments.py
// -----------------------------------------------------------------------

// -----------------------------------------------------------------------
// StylesPart tests
// -----------------------------------------------------------------------

// Mirrors Python: it_provides_access_to_its_styles
func TestStylesPart_Styles(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/styles.xml", opc.CTWmlStyles, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	sp := NewStylesPart(xp)

	styles, err := sp.Styles()
	if err != nil {
		t.Fatalf("Styles(): %v", err)
	}
	if styles == nil {
		t.Fatal("Styles() returned nil")
	}
	// Verify we can access the underlying CT_Styles
	list := styles.StyleList()
	if len(list) == 0 {
		t.Error("expected at least one style in CT_Styles")
	}
}

// Mirrors Python: it_can_construct_a_default_styles_part_to_help
func TestStylesPart_Default(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	sp, err := DefaultStylesPart(pkg)
	if err != nil {
		t.Fatalf("DefaultStylesPart: %v", err)
	}
	if sp == nil {
		t.Fatal("DefaultStylesPart returned nil")
	}
	// Verify element is valid
	if sp.Element() == nil {
		t.Error("default StylesPart has nil element")
	}
	// Verify Styles() works on default
	styles, err := sp.Styles()
	if err != nil {
		t.Fatalf("default Styles(): %v", err)
	}
	if styles == nil {
		t.Fatal("default Styles() returned nil")
	}
}

// -----------------------------------------------------------------------
// SettingsPart tests
// -----------------------------------------------------------------------

// Mirrors Python: it_provides_access_to_its_settings
func TestSettingsPart_SettingsElement(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/settings.xml", opc.CTWmlSettings, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	sp := NewSettingsPart(xp)

	settings, err := sp.SettingsElement()
	if err != nil {
		t.Fatalf("SettingsElement(): %v", err)
	}
	if settings == nil {
		t.Fatal("SettingsElement() returned nil")
	}
}

// Mirrors Python: it_constructs_a_default_settings_part_to_help
func TestSettingsPart_Default(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	sp, err := DefaultSettingsPart(pkg)
	if err != nil {
		t.Fatalf("DefaultSettingsPart: %v", err)
	}
	if sp == nil {
		t.Fatal("DefaultSettingsPart returned nil")
	}
	if sp.Element() == nil {
		t.Error("default SettingsPart has nil element")
	}
	settings, err := sp.SettingsElement()
	if err != nil {
		t.Fatalf("default SettingsElement(): %v", err)
	}
	if settings == nil {
		t.Fatal("default SettingsElement() returned nil")
	}
}

// -----------------------------------------------------------------------
// NumberingPart tests
// -----------------------------------------------------------------------

// Mirrors Python: it_provides_access_to_the_numbering_definitions
func TestNumberingPart_Element(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
  </w:abstractNum>
</w:numbering>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/numbering.xml", opc.CTWmlNumbering, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	np := NewNumberingPart(xp)

	if np.Element() == nil {
		t.Fatal("NumberingPart Element() is nil")
	}
}

// Test NumberingPart via LoadNumberingPart constructor
func TestNumberingPart_Load(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadNumberingPart("/word/numbering.xml", opc.CTWmlNumbering, "", blob, pkg)
	if err != nil {
		t.Fatalf("LoadNumberingPart: %v", err)
	}
	np, ok := part.(*NumberingPart)
	if !ok {
		t.Fatalf("LoadNumberingPart returned %T, want *NumberingPart", part)
	}
	if np.Element() == nil {
		t.Error("loaded NumberingPart has nil element")
	}
}

// -----------------------------------------------------------------------
// HeaderPart / FooterPart tests
// -----------------------------------------------------------------------

// Mirrors Python: it_can_create_a_new_header_part
func TestHeaderPart_New(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		t.Fatalf("NewHeaderPart: %v", err)
	}
	if hp == nil {
		t.Fatal("NewHeaderPart returned nil")
	}
	if hp.Element() == nil {
		t.Error("new HeaderPart has nil element")
	}
	// Verify the root element is w:hdr
	root := hp.Element()
	if root.Tag != "hdr" {
		t.Errorf("root tag = %q, want %q", root.Tag, "hdr")
	}
}

// Mirrors Python: it_loads_default_header_XML_from_a_template_to_help
func TestHeaderPart_TemplateHasContent(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	// Default header template should contain at least one paragraph
	children := hp.Element().ChildElements()
	hasParagraph := false
	for _, c := range children {
		if c.Tag == "p" {
			hasParagraph = true
			break
		}
	}
	if !hasParagraph {
		t.Error("default header template should contain at least one w:p element")
	}
}

// Mirrors Python: it_can_create_a_new_footer_part
func TestFooterPart_New(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	fp, err := NewFooterPart(pkg)
	if err != nil {
		t.Fatalf("NewFooterPart: %v", err)
	}
	if fp == nil {
		t.Fatal("NewFooterPart returned nil")
	}
	if fp.Element() == nil {
		t.Error("new FooterPart has nil element")
	}
	root := fp.Element()
	if root.Tag != "ftr" {
		t.Errorf("root tag = %q, want %q", root.Tag, "ftr")
	}
}

// Mirrors Python: it_loads_default_footer_XML_from_a_template_to_help
func TestFooterPart_TemplateHasContent(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	fp, err := NewFooterPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	children := fp.Element().ChildElements()
	hasParagraph := false
	for _, c := range children {
		if c.Tag == "p" {
			hasParagraph = true
			break
		}
	}
	if !hasParagraph {
		t.Error("default footer template should contain at least one w:p element")
	}
}

// Test LoadHeaderPart constructor
func TestHeaderPart_Load(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
</w:hdr>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadHeaderPart("/word/header1.xml", opc.CTWmlHeader, "", blob, pkg)
	if err != nil {
		t.Fatalf("LoadHeaderPart: %v", err)
	}
	hp, ok := part.(*HeaderPart)
	if !ok {
		t.Fatalf("LoadHeaderPart returned %T, want *HeaderPart", part)
	}
	if hp.Element() == nil {
		t.Error("loaded HeaderPart has nil element")
	}
}

// Test LoadFooterPart constructor
func TestFooterPart_Load(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
</w:ftr>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadFooterPart("/word/footer1.xml", opc.CTWmlFooter, "", blob, pkg)
	if err != nil {
		t.Fatalf("LoadFooterPart: %v", err)
	}
	fp, ok := part.(*FooterPart)
	if !ok {
		t.Fatalf("LoadFooterPart returned %T, want *FooterPart", part)
	}
	if fp.Element() == nil {
		t.Error("loaded FooterPart has nil element")
	}
}

// -----------------------------------------------------------------------
// CommentsPart tests
// -----------------------------------------------------------------------

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

// -----------------------------------------------------------------------
// Cross-cutting: Load constructors produce correct part types
// -----------------------------------------------------------------------

// Mirrors Python: it_is_used_by_loader_to_construct_*_part (via PartFactory)
func TestLoadStylesPart_ReturnsStylesPart(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadStylesPart("/word/styles.xml", opc.CTWmlStyles, "", blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*StylesPart); !ok {
		t.Errorf("LoadStylesPart returned %T, want *StylesPart", part)
	}
}

func TestLoadSettingsPart_ReturnsSettingsPart(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadSettingsPart("/word/settings.xml", opc.CTWmlSettings, "", blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*SettingsPart); !ok {
		t.Errorf("LoadSettingsPart returned %T, want *SettingsPart", part)
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

// Verify that templates are accessible (they use embedded FS)
func TestTemplates_Embedded(t *testing.T) {
	templateTests := []struct {
		name     string
		file     string
		usedBy   string
		opc      bool // uses package?
		ctor     func(*opc.OpcPackage) error
	}{
		{
			"default_styles", "default-styles.xml", "StylesPart",
			true,
			func(pkg *opc.OpcPackage) error { _, err := DefaultStylesPart(pkg); return err },
		},
		{
			"default_settings", "default-settings.xml", "SettingsPart",
			true,
			func(pkg *opc.OpcPackage) error { _, err := DefaultSettingsPart(pkg); return err },
		},
		{
			"default_comments", "default-comments.xml", "CommentsPart",
			true,
			func(pkg *opc.OpcPackage) error { _, err := DefaultCommentsPart(pkg); return err },
		},
		{
			"default_header", "default-header.xml", "HeaderPart",
			true,
			func(pkg *opc.OpcPackage) error { _, err := NewHeaderPart(pkg); return err },
		},
		{
			"default_footer", "default-footer.xml", "FooterPart",
			true,
			func(pkg *opc.OpcPackage) error { _, err := NewFooterPart(pkg); return err },
		},
	}
	for _, tt := range templateTests {
		t.Run(tt.name, func(t *testing.T) {
			pkg := opc.NewOpcPackage(nil)
			if err := tt.ctor(pkg); err != nil {
				t.Errorf("%s: failed to construct from %s: %v", tt.usedBy, tt.file, err)
			}
		})
	}
}

