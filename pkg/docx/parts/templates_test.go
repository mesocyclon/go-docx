package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// Verify that templates are accessible (they use embedded FS)
func TestTemplates_Embedded(t *testing.T) {
	templateTests := []struct {
		name   string
		file   string
		usedBy string
		opc    bool // uses package?
		ctor   func(*opc.OpcPackage) error
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
