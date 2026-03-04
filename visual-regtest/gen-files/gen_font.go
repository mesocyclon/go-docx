package main

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// 03 — Bold, Italic, Underline on runs
func genFontBasic() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Formatting: Bold / Italic / Underline", 1); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}

	r1, err := para.AddRun("Bold text. ")
	if err != nil {
		return nil, err
	}
	if err := r1.SetBold(BoolPtr(true)); err != nil {
		return nil, err
	}

	r2, err := para.AddRun("Italic text. ")
	if err != nil {
		return nil, err
	}
	if err := r2.SetItalic(BoolPtr(true)); err != nil {
		return nil, err
	}

	r3, err := para.AddRun("Underlined text. ")
	if err != nil {
		return nil, err
	}
	u := docx.UnderlineSingle()
	if err := r3.SetUnderline(&u); err != nil {
		return nil, err
	}

	r4, err := para.AddRun("Bold & Italic. ")
	if err != nil {
		return nil, err
	}
	if err := r4.SetBold(BoolPtr(true)); err != nil {
		return nil, err
	}
	if err := r4.SetItalic(BoolPtr(true)); err != nil {
		return nil, err
	}

	r5, err := para.AddRun("Bold, Italic & Underlined.")
	if err != nil {
		return nil, err
	}
	if err := r5.SetBold(BoolPtr(true)); err != nil {
		return nil, err
	}
	if err := r5.SetItalic(BoolPtr(true)); err != nil {
		return nil, err
	}
	if err := r5.SetUnderline(&u); err != nil {
		return nil, err
	}

	return doc, nil
}

// 04 — Advanced font properties: strikethrough, all-caps, small-caps, shadow, emboss, outline, hidden
func genFontAdvanced() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Advanced Font Properties", 1); err != nil {
		return nil, err
	}

	type fontTest struct {
		label string
		apply func(f *docx.Font) error
	}
	tests := []fontTest{
		{"Strikethrough", func(f *docx.Font) error { return f.SetStrike(BoolPtr(true)) }},
		{"Double Strikethrough", func(f *docx.Font) error { return f.SetDoubleStrike(BoolPtr(true)) }},
		{"ALL CAPS", func(f *docx.Font) error { return f.SetAllCaps(BoolPtr(true)) }},
		{"Small Caps", func(f *docx.Font) error { return f.SetSmallCaps(BoolPtr(true)) }},
		{"Shadow", func(f *docx.Font) error { return f.SetShadow(BoolPtr(true)) }},
		{"Emboss", func(f *docx.Font) error { return f.SetEmboss(BoolPtr(true)) }},
		{"Outline", func(f *docx.Font) error { return f.SetOutline(BoolPtr(true)) }},
		{"Imprint (engrave)", func(f *docx.Font) error { return f.SetImprint(BoolPtr(true)) }},
		{"Hidden text", func(f *docx.Font) error { return f.SetHidden(BoolPtr(true)) }},
	}

	for _, tt := range tests {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		r, err := para.AddRun(tt.label)
		if err != nil {
			return nil, err
		}
		if err := tt.apply(r.Font()); err != nil {
			return nil, fmt.Errorf("%s: %w", tt.label, err)
		}
	}
	return doc, nil
}

// 05 — Font color (RGB)
func genFontColor() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Colors", 1); err != nil {
		return nil, err
	}

	colors := []struct {
		label   string
		r, g, b byte
	}{
		{"Red text", 0xFF, 0x00, 0x00},
		{"Green text", 0x00, 0x80, 0x00},
		{"Blue text", 0x00, 0x00, 0xFF},
		{"Orange text", 0xFF, 0xA5, 0x00},
		{"Purple text", 0x80, 0x00, 0x80},
		{"Dark cyan text", 0x00, 0x80, 0x80},
	}

	for _, c := range colors {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(c.label)
		if err != nil {
			return nil, err
		}
		rgb := docx.NewRGBColor(c.r, c.g, c.b)
		if err := run.Font().Color().SetRGB(&rgb); err != nil {
			return nil, err
		}
	}

	// Theme color
	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	run, err := para.AddRun("Theme color: Accent1")
	if err != nil {
		return nil, err
	}
	tc := enum.MsoThemeColorIndexAccent1
	if err := run.Font().Color().SetThemeColor(&tc); err != nil {
		return nil, err
	}

	return doc, nil
}

// 06 — Font sizes
func genFontSize() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Sizes", 1); err != nil {
		return nil, err
	}

	sizes := []float64{8, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72}
	for _, sz := range sizes {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(fmt.Sprintf("%.0fpt text", sz))
		if err != nil {
			return nil, err
		}
		length := docx.Pt(sz)
		if err := run.Font().SetSize(&length); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 31 — Font highlight colors
func genFontHighlight() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Highlight Colors", 1); err != nil {
		return nil, err
	}

	highlights := []struct {
		label string
		color enum.WdColorIndex
	}{
		{"Yellow highlight", enum.WdColorIndexYellow},
		{"Green highlight", enum.WdColorIndexBrightGreen},
		{"Turquoise highlight", enum.WdColorIndexTurquoise},
		{"Pink highlight", enum.WdColorIndexPink},
		{"Blue highlight", enum.WdColorIndexBlue},
		{"Red highlight", enum.WdColorIndexRed},
		{"Dark Blue highlight", enum.WdColorIndexDarkBlue},
		{"Teal highlight", enum.WdColorIndexTeal},
		{"Violet highlight", enum.WdColorIndexViolet},
		{"Gray 50% highlight", enum.WdColorIndexGray50},
		{"Gray 25% highlight", enum.WdColorIndexGray25},
	}

	for _, h := range highlights {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(h.label)
		if err != nil {
			return nil, err
		}
		if err := run.Font().SetHighlightColor(&h.color); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 32 — Font name
func genFontName() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Names", 1); err != nil {
		return nil, err
	}

	fonts := []string{
		"Arial", "Times New Roman", "Courier New",
		"Verdana", "Georgia", "Trebuchet MS",
		"Comic Sans MS", "Impact", "Calibri",
	}

	for _, f := range fonts {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(fmt.Sprintf("The quick brown fox (%s)", f))
		if err != nil {
			return nil, err
		}
		if err := run.Font().SetName(&f); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 33 — Various underline styles
func genUnderlineStyles() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Underline Styles", 1); err != nil {
		return nil, err
	}

	styles := []struct {
		label string
		val   docx.UnderlineVal
	}{
		{"Single", docx.UnderlineSingle()},
		{"Double", docx.UnderlineStyle(enum.WdUnderlineDouble)},
		{"Thick", docx.UnderlineStyle(enum.WdUnderlineThick)},
		{"Dotted", docx.UnderlineStyle(enum.WdUnderlineDotted)},
		{"Dash", docx.UnderlineStyle(enum.WdUnderlineDash)},
		{"Dot-Dash", docx.UnderlineStyle(enum.WdUnderlineDotDash)},
		{"Dot-Dot-Dash", docx.UnderlineStyle(enum.WdUnderlineDotDotDash)},
		{"Wavy", docx.UnderlineStyle(enum.WdUnderlineWavy)},
		{"Words only", docx.UnderlineStyle(enum.WdUnderlineWords)},
		{"Dash Long", docx.UnderlineStyle(enum.WdUnderlineDashLong)},
		{"Wavy Double", docx.UnderlineStyle(enum.WdUnderlineWavyDouble)},
	}

	for _, s := range styles {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(s.label + " underline")
		if err != nil {
			return nil, err
		}
		val := s.val
		if err := run.SetUnderline(&val); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 45 — Subscript and Superscript
func genFontSubSuperscript() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Subscript & Superscript", 1); err != nil {
		return nil, err
	}

	// Superscript: E=mc²
	p1, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := p1.AddRun("E = mc"); err != nil {
		return nil, err
	}
	r1, err := p1.AddRun("2")
	if err != nil {
		return nil, err
	}
	if err := r1.Font().SetSuperscript(BoolPtr(true)); err != nil {
		return nil, err
	}

	// Subscript: H₂O
	p2, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := p2.AddRun("H"); err != nil {
		return nil, err
	}
	r2, err := p2.AddRun("2")
	if err != nil {
		return nil, err
	}
	if err := r2.Font().SetSubscript(BoolPtr(true)); err != nil {
		return nil, err
	}
	if _, err := p2.AddRun("O"); err != nil {
		return nil, err
	}

	// Mixed: x² + y₁
	p3, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := p3.AddRun("x"); err != nil {
		return nil, err
	}
	rSup, err := p3.AddRun("2")
	if err != nil {
		return nil, err
	}
	if err := rSup.Font().SetSuperscript(BoolPtr(true)); err != nil {
		return nil, err
	}
	if _, err := p3.AddRun(" + y"); err != nil {
		return nil, err
	}
	rSub, err := p3.AddRun("1")
	if err != nil {
		return nil, err
	}
	if err := rSub.Font().SetSubscript(BoolPtr(true)); err != nil {
		return nil, err
	}

	return doc, nil
}
