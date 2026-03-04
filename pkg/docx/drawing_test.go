package docx

import (
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// drawing_test.go — Drawing (Batch 1)
// Mirrors Python: tests/test_drawing.py
// -----------------------------------------------------------------------

// Mirrors Python: it_knows_when_it_contains_a_Picture
func TestDrawing_HasPicture(t *testing.T) {
	tests := []struct {
		name     string
		xml      string
		expected bool
	}{
		{
			"has_picture",
			`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
				xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<wp:inline>
					<a:graphic>
						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
							<pic:pic/>
						</a:graphicData>
					</a:graphic>
				</wp:inline>
			</w:drawing>`,
			true,
		},
		{
			"no_picture_chart",
			`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
				xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<wp:inline>
					<a:graphic>
						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
					</a:graphic>
				</wp:inline>
			</w:drawing>`,
			false,
		},
		{
			"empty_drawing",
			`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			el, err := oxml.ParseXml([]byte(tt.xml))
			if err != nil {
				t.Fatal(err)
			}
			d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
			drawing := newDrawing(d, nil)
			if got := drawing.HasPicture(); got != tt.expected {
				t.Errorf("HasPicture() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_provides_access_to_the_image
// Tests both error paths: no picture, and no part.
func TestDrawing_ImagePart_NoPicture(t *testing.T) {
	// Drawing with a chart (no pic:pic) → ImagePart should error "does not contain a picture"
	xml := `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
		<wp:inline>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
			</a:graphic>
		</wp:inline>
	</w:drawing>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
	drawing := newDrawing(d, nil)

	_, err = drawing.ImagePart()
	if err == nil {
		t.Error("expected error for ImagePart on drawing without picture")
	}
}

func TestDrawing_ImagePart_NoPart(t *testing.T) {
	// Drawing WITH a picture but nil part → should error gracefully, not panic
	xml := `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<wp:inline>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic>
						<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:inline>
	</w:drawing>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
	drawing := newDrawing(d, nil) // nil part → should error, not panic

	_, err = drawing.ImagePart()
	if err == nil {
		t.Error("expected error for ImagePart with nil part")
	}
}

// -----------------------------------------------------------------------
// walkPath
// -----------------------------------------------------------------------

func TestWalkPath_SingleTag(t *testing.T) {
	t.Parallel()
	xml := `<root><child/><child/><other/></root>`
	root := parseTestXml(t, xml)

	got := walkPath(root, "child")
	if len(got) != 2 {
		t.Errorf("walkPath(root, child) = %d elements, want 2", len(got))
	}
}

func TestWalkPath_MultiLevel(t *testing.T) {
	t.Parallel()
	xml := `<root><a><b><c/></b></a></root>`
	root := parseTestXml(t, xml)

	got := walkPath(root, "a", "b", "c")
	if len(got) != 1 {
		t.Errorf("walkPath(root, a, b, c) = %d elements, want 1", len(got))
	}
	if got[0].Tag != "c" {
		t.Errorf("got tag %q, want %q", got[0].Tag, "c")
	}
}

func TestWalkPath_NoMatch(t *testing.T) {
	t.Parallel()
	xml := `<root><a><b/></a></root>`
	root := parseTestXml(t, xml)

	got := walkPath(root, "a", "x")
	if len(got) != 0 {
		t.Errorf("walkPath with missing tag = %d elements, want 0", len(got))
	}
}

func TestWalkPath_EmptyTags(t *testing.T) {
	t.Parallel()
	xml := `<root><a/></root>`
	root := parseTestXml(t, xml)

	got := walkPath(root)
	if len(got) != 1 || got[0].Tag != "root" {
		t.Errorf("walkPath with no tags should return root, got %d elements", len(got))
	}
}

func TestWalkPath_BranchingPaths(t *testing.T) {
	t.Parallel()
	// Two independent a → b paths
	xml := `<root><a><b id="1"/></a><a><b id="2"/></a></root>`
	root := parseTestXml(t, xml)

	got := walkPath(root, "a", "b")
	if len(got) != 2 {
		t.Errorf("walkPath through two branches = %d elements, want 2", len(got))
	}
}

func TestWalkPath_DoesNotRecurse(t *testing.T) {
	t.Parallel()
	// "b" exists at depth 3 but not as direct child of "a" at depth 1
	xml := `<root><a><x><b/></x></a></root>`
	root := parseTestXml(t, xml)

	// walkPath(root, "a", "b") should NOT find the b nested under x
	got := walkPath(root, "a", "b")
	if len(got) != 0 {
		t.Errorf("walkPath should not skip levels, got %d elements", len(got))
	}
}

func TestWalkPath_BrokenChainStopsEarly(t *testing.T) {
	t.Parallel()
	xml := `<root><a><b><c/></b></a></root>`
	root := parseTestXml(t, xml)

	// Ask for a path that diverges at the second step
	got := walkPath(root, "a", "MISSING", "c")
	if len(got) != 0 {
		t.Errorf("walkPath with broken middle = %d elements, want 0", len(got))
	}
}

// -----------------------------------------------------------------------
// findBlipRId
// -----------------------------------------------------------------------

func TestFindBlipRId_InlinePicture(t *testing.T) {
	t.Parallel()
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<a:graphic>
			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<pic:pic>
					<pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>
				</pic:pic>
			</a:graphicData>
		</a:graphic>
	</wp:inline>`
	root := parseTestXml(t, xml)

	got := findBlipRId(root)
	if got != "rId7" {
		t.Errorf("findBlipRId = %q, want %q", got, "rId7")
	}
}

func TestFindBlipRId_AnchorPicture(t *testing.T) {
	t.Parallel()
	xml := `<wp:anchor
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<a:graphic>
			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<pic:pic>
					<pic:blipFill><a:blip r:embed="rId3"/></pic:blipFill>
				</pic:pic>
			</a:graphicData>
		</a:graphic>
	</wp:anchor>`
	root := parseTestXml(t, xml)

	got := findBlipRId(root)
	if got != "rId3" {
		t.Errorf("findBlipRId = %q, want %q", got, "rId3")
	}
}

func TestFindBlipRId_NoPicture(t *testing.T) {
	t.Parallel()
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
		<a:graphic>
			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
		</a:graphic>
	</wp:inline>`
	root := parseTestXml(t, xml)

	got := findBlipRId(root)
	if got != "" {
		t.Errorf("findBlipRId on chart drawing = %q, want empty", got)
	}
}

func TestFindBlipRId_EmptyElement(t *testing.T) {
	t.Parallel()
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/>`
	root := parseTestXml(t, xml)

	got := findBlipRId(root)
	if got != "" {
		t.Errorf("findBlipRId on empty = %q, want empty", got)
	}
}

func TestFindBlipRId_BlipFillWithoutEmbed(t *testing.T) {
	t.Parallel()
	// blip element exists but has no r:embed attribute (e.g. linked image with r:link only)
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<a:graphic>
			<a:graphicData>
				<pic:pic>
					<pic:blipFill><a:blip r:link="rId9"/></pic:blipFill>
				</pic:pic>
			</a:graphicData>
		</a:graphic>
	</wp:inline>`
	root := parseTestXml(t, xml)

	got := findBlipRId(root)
	if got != "" {
		t.Errorf("findBlipRId with r:link only = %q, want empty", got)
	}
}

func TestFindBlipRId_BlipFillAtWrongDepth_NotFound(t *testing.T) {
	t.Parallel()
	// blipFill directly under inline — not at the correct path
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<pic:blipFill><a:blip r:embed="rIdBAD"/></pic:blipFill>
	</wp:inline>`
	root := parseTestXml(t, xml)

	got := findBlipRId(root)
	if got != "" {
		t.Errorf("findBlipRId with blipFill at wrong depth = %q, want empty", got)
	}
}

// -----------------------------------------------------------------------
// findPicInGraphicData
// -----------------------------------------------------------------------

func TestFindPicInGraphicData_Found(t *testing.T) {
	t.Parallel()
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
		<a:graphic>
			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<pic:pic/>
			</a:graphicData>
		</a:graphic>
	</wp:inline>`
	root := parseTestXml(t, xml)

	if !findPicInGraphicData(root) {
		t.Error("findPicInGraphicData should be true for inline with pic:pic")
	}
}

func TestFindPicInGraphicData_ChartNotPic(t *testing.T) {
	t.Parallel()
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
		<a:graphic>
			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
				<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
			</a:graphicData>
		</a:graphic>
	</wp:inline>`
	root := parseTestXml(t, xml)

	if findPicInGraphicData(root) {
		t.Error("findPicInGraphicData should be false for chart drawing")
	}
}

func TestFindPicInGraphicData_Empty(t *testing.T) {
	t.Parallel()
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/>`
	root := parseTestXml(t, xml)

	if findPicInGraphicData(root) {
		t.Error("findPicInGraphicData should be false for empty inline")
	}
}

func TestFindPicInGraphicData_PicAtWrongDepth(t *testing.T) {
	t.Parallel()
	// pic directly under inline, not under graphic → graphicData
	xml := `<wp:inline
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
		<pic:pic/>
	</wp:inline>`
	root := parseTestXml(t, xml)

	if findPicInGraphicData(root) {
		t.Error("findPicInGraphicData should be false when pic is at wrong depth")
	}
}

// -----------------------------------------------------------------------
// HasPicture — floating (anchor) variant
// -----------------------------------------------------------------------

func TestDrawing_HasPicture_Anchor(t *testing.T) {
	t.Parallel()
	xml := `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
		<wp:anchor>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic/>
				</a:graphicData>
			</a:graphic>
		</wp:anchor>
	</w:drawing>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
	drawing := newDrawing(d, nil)

	if !drawing.HasPicture() {
		t.Error("HasPicture should be true for anchor picture")
	}
}

// -----------------------------------------------------------------------
// pictureRId via ImagePart (integration-level)
// -----------------------------------------------------------------------

func TestDrawing_PictureRId_Anchor(t *testing.T) {
	t.Parallel()
	// Floating picture → pictureRId should still find the rId
	xml := `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<wp:anchor>
			<a:graphic>
				<a:graphicData>
					<pic:pic>
						<pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:anchor>
	</w:drawing>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
	drawing := newDrawing(d, nil)

	// ImagePart will fail (nil part), but the error message tells us the rId was found
	_, err = drawing.ImagePart()
	if err == nil {
		t.Fatal("expected error with nil part")
	}
	// If rId was NOT found, error would say "does not contain a picture"
	if got := err.Error(); got == "docx: drawing does not contain a picture" {
		t.Error("pictureRId should have found rId5 in anchor drawing")
	}
}

// -----------------------------------------------------------------------
// helpers
// -----------------------------------------------------------------------

func parseTestXml(t *testing.T, xml string) *etree.Element {
	t.Helper()
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatalf("parseTestXml: %v", err)
	}
	return el
}
