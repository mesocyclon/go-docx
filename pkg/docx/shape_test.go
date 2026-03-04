package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// shape_test.go — InlineShapes / InlineShape (Batch 1)
// Mirrors Python: tests/test_shape.py
// Complements inlineshapes_test.go which has round-trip integration tests.
// -----------------------------------------------------------------------

const wpNS = `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`
const wNS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`
const aNS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`
const picNS = `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`
const rNS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`

func makeInlineShapesBody(t *testing.T, shapeCount int) *InlineShapes {
	t.Helper()
	inner := ""
	for i := 0; i < shapeCount; i++ {
		inner += `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/></wp:inline></w:drawing></w:r></w:p>`
	}
	xml := `<w:body ` + wNS + ` ` + wpNS + `>` + inner + `</w:body>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	return newInlineShapes(el, nil)
}

// Mirrors Python: it_can_iterate_over_InlineShape_instances
func TestInlineShapes_Iter_XML(t *testing.T) {
	iss := makeInlineShapesBody(t, 3)
	items := iss.Iter()
	if len(items) != 3 {
		t.Errorf("Iter() len = %d, want 3", len(items))
	}
}

// Mirrors Python: it_provides_indexed_access
func TestInlineShapes_Get_XML(t *testing.T) {
	iss := makeInlineShapesBody(t, 2)
	shape, err := iss.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	if shape == nil {
		t.Fatal("Get(0) returned nil")
	}
	shape2, err := iss.Get(1)
	if err != nil {
		t.Fatal(err)
	}
	if shape2 == nil {
		t.Fatal("Get(1) returned nil")
	}
}

// Mirrors Python: it_raises_on_indexed_access_out_of_range
func TestInlineShapes_Get_OutOfRange_XML(t *testing.T) {
	iss := makeInlineShapesBody(t, 0)
	_, err := iss.Get(0)
	if err == nil {
		t.Error("expected error for Get(0) on empty")
	}
	_, err = iss.Get(-1)
	if err == nil {
		t.Error("expected error for Get(-1)")
	}
}

// Mirrors Python: it_knows_what_type_of_shape_it_is
func TestInlineShape_Type_XML(t *testing.T) {
	tests := []struct {
		name     string
		xml      string
		expected enum.WdInlineShapeType
	}{
		{
			"picture",
			`<wp:inline ` + wpNS + ` ` + aNS + ` ` + picNS + ` ` + rNS + `>
				<wp:extent cx="914400" cy="914400"/>
				<a:graphic>
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:pic>
							<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
						</pic:pic>
					</a:graphicData>
				</a:graphic>
			</wp:inline>`,
			enum.WdInlineShapeTypePicture,
		},
		{
			"chart",
			`<wp:inline ` + wpNS + ` ` + aNS + `>
				<wp:extent cx="914400" cy="914400"/>
				<a:graphic>
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
				</a:graphic>
			</wp:inline>`,
			enum.WdInlineShapeTypeChart,
		},
		{
			"smart_art",
			`<wp:inline ` + wpNS + ` ` + aNS + `>
				<wp:extent cx="914400" cy="914400"/>
				<a:graphic>
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
				</a:graphic>
			</wp:inline>`,
			enum.WdInlineShapeTypeSmartArt,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			el, err := oxml.ParseXml([]byte(tt.xml))
			if err != nil {
				t.Fatal(err)
			}
			is := newInlineShape(&oxml.CT_Inline{Element: oxml.WrapElement(el)}, nil)
			gotType, err := is.Type()
			if err != nil {
				t.Fatal(err)
			}
			if gotType != tt.expected {
				t.Errorf("Type() = %d, want %d", gotType, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_knows_its_display_dimensions
func TestInlineShape_Dimensions_XML(t *testing.T) {
	xml := `<wp:inline ` + wpNS + `>
		<wp:extent cx="914400" cy="457200"/>
	</wp:inline>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	is := newInlineShape(&oxml.CT_Inline{Element: oxml.WrapElement(el)}, nil)

	w, err := is.Width()
	if err != nil {
		t.Fatal(err)
	}
	if w != Length(914400) {
		t.Errorf("Width() = %d, want 914400", w)
	}
	h, err := is.Height()
	if err != nil {
		t.Fatal(err)
	}
	if h != Length(457200) {
		t.Errorf("Height() = %d, want 457200", h)
	}
}

// Mirrors Python: it_can_change_its_display_dimensions
func TestInlineShape_SetDimensions_XML(t *testing.T) {
	t.Run("picture_sets_extent_and_spPr", func(t *testing.T) {
		xml := `<wp:inline ` + wpNS + ` ` + aNS + ` ` + picNS + ` ` + rNS + `>
			<wp:extent cx="914400" cy="914400"/>
			<wp:docPr id="1" name="Picture 1"/>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic>
						<pic:nvPicPr><pic:cNvPr id="0" name="img.png"/><pic:cNvPicPr/></pic:nvPicPr>
						<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
						<pic:spPr>
							<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
						</pic:spPr>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:inline>`
		el, err := oxml.ParseXml([]byte(xml))
		if err != nil {
			t.Fatal(err)
		}
		is := newInlineShape(&oxml.CT_Inline{Element: oxml.WrapElement(el)}, nil)

		newWidth := Inches(2)
		if err := is.SetWidth(newWidth); err != nil {
			t.Fatal(err)
		}
		w, _ := is.Width()
		if w != newWidth {
			t.Errorf("Width() after set = %d, want %d", w, newWidth)
		}

		newHeight := Inches(3)
		if err := is.SetHeight(newHeight); err != nil {
			t.Fatal(err)
		}
		h, _ := is.Height()
		if h != newHeight {
			t.Errorf("Height() after set = %d, want %d", h, newHeight)
		}
	})

	t.Run("chart_sets_extent_only", func(t *testing.T) {
		xml := `<wp:inline ` + wpNS + ` ` + aNS + `>
			<wp:extent cx="914400" cy="914400"/>
			<wp:docPr id="1" name="Chart 1"/>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
			</a:graphic>
		</wp:inline>`
		el, err := oxml.ParseXml([]byte(xml))
		if err != nil {
			t.Fatal(err)
		}
		is := newInlineShape(&oxml.CT_Inline{Element: oxml.WrapElement(el)}, nil)

		newWidth := Inches(4)
		if err := is.SetWidth(newWidth); err != nil {
			t.Fatal(err)
		}
		w, _ := is.Width()
		if w != newWidth {
			t.Errorf("Width() after set = %d, want %d", w, newWidth)
		}

		newHeight := Inches(5)
		if err := is.SetHeight(newHeight); err != nil {
			t.Fatal(err)
		}
		h, _ := is.Height()
		if h != newHeight {
			t.Errorf("Height() after set = %d, want %d", h, newHeight)
		}
	})
}
