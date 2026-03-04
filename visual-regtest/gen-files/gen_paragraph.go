package main

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// 01 — All heading levels 0–9
func genHeadings() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	for level := 0; level <= 9; level++ {
		text := fmt.Sprintf("Heading Level %d", level)
		if level == 0 {
			text = "Document Title (Level 0)"
		}
		if _, err := doc.AddHeading(text, level); err != nil {
			return nil, fmt.Errorf("heading level %d: %w", level, err)
		}
		if _, err := doc.AddParagraph(fmt.Sprintf("Body text after heading level %d. Lorem ipsum dolor sit amet.", level)); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 02 — Paragraphs with built-in styles
func genParagraphStyles() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	styles := []string{
		"Normal", "Title", "Subtitle",
		"Heading 1", "Heading 2", "Heading 3",
		"List Bullet", "List Number",
		"Quote", "Intense Quote",
		"No Spacing",
	}
	for _, s := range styles {
		if _, err := doc.AddParagraph(
			fmt.Sprintf("This paragraph uses the %q style.", s),
			docx.StyleName(s),
		); err != nil {
			return nil, fmt.Errorf("style %q: %w", s, err)
		}
	}
	return doc, nil
}

// 07 — Paragraph alignment
func genParagraphAlignment() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Alignment", 1); err != nil {
		return nil, err
	}

	aligns := []struct {
		name string
		val  enum.WdParagraphAlignment
	}{
		{"Left aligned", enum.WdParagraphAlignmentLeft},
		{"Center aligned", enum.WdParagraphAlignmentCenter},
		{"Right aligned", enum.WdParagraphAlignmentRight},
		{"Justified – Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.", enum.WdParagraphAlignmentJustify},
		{"Distributed – Short text fills full width", enum.WdParagraphAlignmentDistribute},
	}

	for _, a := range aligns {
		para, err := doc.AddParagraph(a.name)
		if err != nil {
			return nil, err
		}
		if err := para.SetAlignment(&a.val); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 08 — Paragraph indentation
func genParagraphIndent() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Indentation", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("Left indent 720 twips (0.5 inch)")
	if err != nil {
		return nil, err
	}
	if err := p1.ParagraphFormat().SetLeftIndent(IntPtr(720)); err != nil {
		return nil, err
	}

	p2, err := doc.AddParagraph("Right indent 720 twips (0.5 inch)")
	if err != nil {
		return nil, err
	}
	if err := p2.ParagraphFormat().SetRightIndent(IntPtr(720)); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("Left + Right indent 1440 twips (1 inch each)")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetLeftIndent(IntPtr(1440)); err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetRightIndent(IntPtr(1440)); err != nil {
		return nil, err
	}

	p4, err := doc.AddParagraph("First-line indent 360 twips (0.25 inch). Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
	if err != nil {
		return nil, err
	}
	if err := p4.ParagraphFormat().SetFirstLineIndent(IntPtr(360)); err != nil {
		return nil, err
	}

	p5, err := doc.AddParagraph("Hanging indent: left=720, firstLine=-360. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt.")
	if err != nil {
		return nil, err
	}
	if err := p5.ParagraphFormat().SetLeftIndent(IntPtr(720)); err != nil {
		return nil, err
	}
	if err := p5.ParagraphFormat().SetFirstLineIndent(IntPtr(-360)); err != nil {
		return nil, err
	}

	return doc, nil
}

// 09 — Space before/after
func genParagraphSpacing() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Spacing", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("Space before = 480 twips (24pt)")
	if err != nil {
		return nil, err
	}
	if err := p1.ParagraphFormat().SetSpaceBefore(IntPtr(480)); err != nil {
		return nil, err
	}

	p2, err := doc.AddParagraph("Space after = 480 twips (24pt)")
	if err != nil {
		return nil, err
	}
	if err := p2.ParagraphFormat().SetSpaceAfter(IntPtr(480)); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("Space before=240 and after=240 (12pt each)")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetSpaceBefore(IntPtr(240)); err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetSpaceAfter(IntPtr(240)); err != nil {
		return nil, err
	}

	if _, err := doc.AddParagraph("Normal spacing after."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 10 — Line spacing
func genLineSpacing() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Line Spacing", 1); err != nil {
		return nil, err
	}

	loremShort := "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation."

	p1, err := doc.AddParagraph("Single spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls1 := docx.LineSpacingMultiple(1.0)
	if err := p1.ParagraphFormat().SetLineSpacing(&ls1); err != nil {
		return nil, err
	}

	p2, err := doc.AddParagraph("1.5 line spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls2 := docx.LineSpacingMultiple(1.5)
	if err := p2.ParagraphFormat().SetLineSpacing(&ls2); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("Double line spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls3 := docx.LineSpacingMultiple(2.0)
	if err := p3.ParagraphFormat().SetLineSpacing(&ls3); err != nil {
		return nil, err
	}

	p4, err := doc.AddParagraph("Exact 18pt spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls4 := docx.LineSpacingTwips(360)
	if err := p4.ParagraphFormat().SetLineSpacing(&ls4); err != nil {
		return nil, err
	}

	p5, err := doc.AddParagraph("SetLineSpacingRule(Double): " + loremShort)
	if err != nil {
		return nil, err
	}
	if err := p5.ParagraphFormat().SetLineSpacingRule(enum.WdLineSpacingDouble); err != nil {
		return nil, err
	}

	return doc, nil
}

// 11 — Tab stops
func genTabStops() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Tab Stops", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts1 := p1.ParagraphFormat().TabStops()
	if _, err := ts1.AddTabStop(2880, enum.WdTabAlignmentLeft, enum.WdTabLeaderSpaces); err != nil {
		return nil, err
	}
	r1, err := p1.AddRun("Before tab")
	if err != nil {
		return nil, err
	}
	r1.AddTab()
	r1.AddText("After left tab at 2\"")

	p2, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts2 := p2.ParagraphFormat().TabStops()
	if _, err := ts2.AddTabStop(4680, enum.WdTabAlignmentCenter, enum.WdTabLeaderDots); err != nil {
		return nil, err
	}
	r2, err := p2.AddRun("Item")
	if err != nil {
		return nil, err
	}
	r2.AddTab()
	r2.AddText("Centered with dots")

	p3, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts3 := p3.ParagraphFormat().TabStops()
	if _, err := ts3.AddTabStop(8640, enum.WdTabAlignmentRight, enum.WdTabLeaderDashes); err != nil {
		return nil, err
	}
	r3, err := p3.AddRun("Left text")
	if err != nil {
		return nil, err
	}
	r3.AddTab()
	r3.AddText("$99.99")

	p4, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts4 := p4.ParagraphFormat().TabStops()
	if _, err := ts4.AddTabStop(4320, enum.WdTabAlignmentDecimal, enum.WdTabLeaderSpaces); err != nil {
		return nil, err
	}
	r4, err := p4.AddRun("")
	if err != nil {
		return nil, err
	}
	r4.AddTab()
	r4.AddText("123.456 (decimal tab)")

	return doc, nil
}

// 27 — Custom styles
func genCustomStyles() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	styles, err := doc.Styles()
	if err != nil {
		return nil, err
	}

	ps, err := styles.AddStyle("CustomParagraph", enum.WdStyleTypeParagraph, false)
	if err != nil {
		return nil, err
	}
	sz := docx.Pt(14)
	if err := ps.Font().SetSize(&sz); err != nil {
		return nil, err
	}
	if err := ps.Font().SetBold(BoolPtr(true)); err != nil {
		return nil, err
	}
	rgb := docx.NewRGBColor(0x00, 0x66, 0xCC)
	if err := ps.Font().Color().SetRGB(&rgb); err != nil {
		return nil, err
	}
	if err := ps.ParagraphFormat().SetSpaceBefore(IntPtr(240)); err != nil {
		return nil, err
	}
	if err := ps.ParagraphFormat().SetSpaceAfter(IntPtr(120)); err != nil {
		return nil, err
	}

	cs, err := styles.AddStyle("CustomChar", enum.WdStyleTypeCharacter, false)
	if err != nil {
		return nil, err
	}
	if err := cs.Font().SetItalic(BoolPtr(true)); err != nil {
		return nil, err
	}
	rgb2 := docx.NewRGBColor(0xCC, 0x00, 0x00)
	if err := cs.Font().Color().SetRGB(&rgb2); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Custom Styles", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Using custom paragraph style:", docx.StyleName("CustomParagraph")); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Another paragraph with custom style.", docx.StyleName("CustomParagraph")); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := para.AddRun("Normal text with "); err != nil {
		return nil, err
	}
	if _, err := para.AddRun("custom character style", docx.StyleName("CustomChar")); err != nil {
		return nil, err
	}
	if _, err := para.AddRun(" applied."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 29 — Paragraph format: widow control, keep together, keep with next
func genParagraphFormatFlow() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Flow Control", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("WidowControl = true: This paragraph has widow/orphan control enabled. Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
	if err != nil {
		return nil, err
	}
	if err := p1.ParagraphFormat().SetWidowControl(BoolPtr(true)); err != nil {
		return nil, err
	}

	p2, err := doc.AddParagraph("KeepTogether = true: This paragraph's lines are kept together on the same page. Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
	if err != nil {
		return nil, err
	}
	if err := p2.ParagraphFormat().SetKeepTogether(BoolPtr(true)); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("KeepWithNext = true: This paragraph stays with the next.")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetKeepWithNext(BoolPtr(true)); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This paragraph follows the keep-with-next paragraph."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 40 — Paragraph.Clear / Paragraph.SetText
func genParagraphClearSetText() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Clear & SetText", 1); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("Original text that will be replaced.")
	if err != nil {
		return nil, err
	}
	para.Clear()
	if err := para.SetText("Replaced text via SetText()."); err != nil {
		return nil, err
	}

	if _, err := doc.AddParagraph(fmt.Sprintf("Paragraph text reads: %q", para.Text())); err != nil {
		return nil, err
	}

	return doc, nil
}

// 46 — Tab and newline characters in text
func genTabAndNewlineInText() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Tab and Newline in Text", 1); err != nil {
		return nil, err
	}

	if _, err := doc.AddParagraph("Column1\tColumn2\tColumn3"); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Line 1\nLine 2\nLine 3"); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Name:\tJohn\nAge:\t30\nCity:\tNew York"); err != nil {
		return nil, err
	}

	return doc, nil
}
