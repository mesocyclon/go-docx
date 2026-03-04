package main

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"

	"github.com/vortex/go-docx/pkg/docx"
	. "github.com/vortex/go-docx/visual-regtest/internal/docfmt"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

type sourceSpec struct {
	filename string
	builder  func() (*docx.Document, error)
}

func allSources() []sourceSpec {
	return []sourceSpec{
		{"src_paragraph.docx", buildSrcParagraph},
		{"src_paragraph_ru.docx", buildSrcParagraphRu},
		{"src_multi_para.docx", buildSrcMultiPara},
		{"src_table.docx", buildSrcTable},
		{"src_complex.docx", buildSrcComplex},
		{"src_image.docx", buildSrcImage},
		{"src_heading.docx", buildSrcHeading},
		{"src_empty.docx", buildSrcEmpty},
		{"src_cyrillic.docx", buildSrcCyrillic},
		{"src_large.docx", buildSrcLarge},
		{"src_styled.docx", buildSrcStyled},
		{"src_a.docx", buildSrcA},
		{"src_b.docx", buildSrcB},
	}
}

func buildSrcParagraph() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Inserted paragraph."))
	return doc, nil
}

func buildSrcParagraphRu() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Вставленный абзац на русском языке."))
	return doc, nil
}

func buildSrcMultiPara() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Первый абзац источника."))
	Must(doc.AddParagraph("Второй абзац источника."))
	Must(doc.AddParagraph("Третий абзац источника."))
	return doc, nil
}

func buildSrcTable() (*docx.Document, error) {
	doc := Must(docx.New())
	tbl := Must(doc.AddTable(2, 3, docx.StyleName("Table Grid")))
	FillTable(tbl, [][]string{
		{"Заг1", "Заг2", "Заг3"},
		{"A", "Б", "В"},
	})
	return doc, nil
}

func buildSrcComplex() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Абзац перед таблицей."))
	tbl := Must(doc.AddTable(2, 2, docx.StyleName("Table Grid")))
	FillTable(tbl, [][]string{
		{"Яч1", "Яч2"},
		{"Яч3", "Яч4"},
	})
	Must(doc.AddParagraph("Абзац после таблицы."))
	return doc, nil
}

func buildSrcImage() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Абзац с изображением ниже:"))
	img := image.NewRGBA(image.Rect(0, 0, 120, 60))
	for y := 0; y < 60; y++ {
		for x := 0; x < 120; x++ {
			img.Set(x, y, color.RGBA{
				R: uint8(x * 255 / 120),
				G: uint8(y * 255 / 60),
				B: 200, A: 255,
			})
		}
	}
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		return nil, fmt.Errorf("encoding PNG: %w", err)
	}
	w := int64(docx.Inches(1.5).Emu())
	h := int64(docx.Inches(0.75).Emu())
	Must(doc.AddPicture(bytes.NewReader(buf.Bytes()), &w, &h))
	Must(doc.AddParagraph("Абзац после изображения."))
	return doc, nil
}

func buildSrcHeading() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Заголовок первого уровня", docx.StyleName("Heading 1")))
	return doc, nil
}

func buildSrcEmpty() (*docx.Document, error) {
	doc := Must(docx.New())
	return doc, nil
}

func buildSrcCyrillic() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("Кириллический текст: Привет, мир!"))
	tbl := Must(doc.AddTable(2, 2, docx.StyleName("Table Grid")))
	FillTable(tbl, [][]string{
		{"Колонка1", "Колонка2"},
		{"Данные1", "Данные2"},
	})
	return doc, nil
}

func buildSrcLarge() (*docx.Document, error) {
	doc := Must(docx.New())
	for i := 1; i <= 5; i++ {
		Must(doc.AddParagraph(fmt.Sprintf(
			"Большой абзац %d: Съешь же ещё этих мягких французских булок, да выпей чаю.", i)))
	}
	tbl := Must(doc.AddTable(4, 3, docx.StyleName("Table Grid")))
	for r := 0; r < 4; r++ {
		for c := 0; c < 3; c++ {
			SetCell(tbl, r, c, fmt.Sprintf("Р%dК%d", r, c))
		}
	}
	return doc, nil
}

func buildSrcStyled() (*docx.Document, error) {
	doc := Must(docx.New())
	p := Must(doc.AddParagraph(""))
	rBold, _ := p.AddRun("Жирный текст")
	_ = rBold.SetBold(BoolPtr(true))
	_, _ = p.AddRun(", обычный, ")
	rItalic, _ := p.AddRun("курсив")
	_ = rItalic.SetItalic(BoolPtr(true))
	_, _ = p.AddRun(".")
	return doc, nil
}

func buildSrcA() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("[Источник A]"))
	return doc, nil
}

func buildSrcB() (*docx.Document, error) {
	doc := Must(docx.New())
	Must(doc.AddParagraph("[Источник Б]"))
	return doc, nil
}
