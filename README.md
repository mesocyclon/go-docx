# go-docx

A pure Go library for creating, reading, and modifying Microsoft Word (.docx) documents. The API is modeled after [python-docx](https://python-docx.readthedocs.io/), bringing the same intuitive document manipulation patterns to Go with full type safety.

[English](#features) | [Русский](#русский)

## Features

- **Create** new empty documents from a built-in default template
- **Open** existing `.docx` files from a path, byte slice, or `io.ReaderAt`
- **Save** to a writer or file path
- **Paragraphs** — add, iterate, clear; set alignment, style, and full paragraph formatting (indentation, spacing, tab stops, keep-together, widow/orphan control)
- **Runs** — text runs with character formatting: bold, italic, underline (17 styles), font name, font size, color, highlighting, subscript/superscript, all-caps, small-caps, strikethrough, shadow, emboss, and more
- **Tables** — create and modify tables with rows, columns, cell merging, cell width, vertical alignment, and table style
- **Sections** — page size, orientation, margins, header/footer references, section break types
- **Headers & Footers** — default, first-page, and even-page variants; odd/even page toggle via settings
- **Images** — inline pictures from `io.ReadSeeker` or pre-built `ImagePart`; supports JPEG, PNG, GIF, BMP, and TIFF with automatic dimension detection
- **Hyperlinks** — read addresses, fragments, and contained runs
- **Styles** — access, create, and modify paragraph/character/table styles; BabelFish name translation (UI ↔ internal); latent style management
- **Comments** — add comments anchored to run ranges with author and initials; iterate and modify
- **Core Properties** — author, title, subject, keywords, category, comments, revision, created/modified dates
- **Document Settings** — odd/even page headers
- **Numbering** — list definitions imported during content replacement
- **Shapes & Drawings** — inline shapes with width/height control
- **Text Replacement** — `ReplaceText` across body, headers, footers, and comments with deduplication
- **Table Replacement** — `ReplaceWithTable` replaces text placeholders with structured tables described by `TableData`
- **Content Replacement** — `ReplaceWithContent` replaces placeholders with the entire body of another `.docx` document, including images, hyperlinks, styles, numbering, footnotes, and endnotes with full relationship remapping and SHA-256 image deduplication

## Installation

```bash
go get github.com/vortex/go-docx
```

Requires **Go 1.25** or later.

## Quick Start

```go
package main

import (
	"log"

	"github.com/vortex/go-docx/pkg/docx"
)

func main() {
	doc, err := docx.New()
	if err != nil {
		log.Fatal(err)
	}

	doc.AddHeading("Hello, World!", 1)
	doc.AddParagraph("Created with go-docx.")

	if err := doc.SaveFile("hello.docx"); err != nil {
		log.Fatal(err)
	}
}
```

## Opening and Saving Documents

```go
// Create a new empty document
doc, err := docx.New()

// Open from file
doc, err := docx.OpenFile("input.docx")

// Open from byte slice
data, _ := os.ReadFile("input.docx")
doc, err := docx.OpenBytes(data)

// Open from io.ReaderAt (e.g. http body, S3 object)
doc, err := docx.Open(reader, size)

// Save to file
err = doc.SaveFile("output.docx")

// Save to any io.Writer
var buf bytes.Buffer
err = doc.Save(&buf)
```

## Paragraphs and Headings

```go
// Add a heading (level 0 = Title, 1-9 = Heading N)
doc.AddHeading("Document Title", 0)
doc.AddHeading("Chapter 1", 1)
doc.AddHeading("Section 1.1", 2)

// Add a plain paragraph
p, _ := doc.AddParagraph("Hello, World!")

// Add paragraph with a named style
p, _ = doc.AddParagraph("Styled text", docx.StyleName("Heading 1"))

// Get all paragraphs
paras, _ := doc.Paragraphs()
for _, p := range paras {
	fmt.Println(p.Text())
}

// Replace paragraph content
p.SetText("New text")

// Clear paragraph (remove content, keep formatting)
p.Clear()

// Add page break
doc.AddPageBreak()
```

## Runs and Character Formatting

A paragraph consists of one or more *runs* — contiguous pieces of text sharing the same formatting.

```go
p, _ := doc.AddParagraph("")

// Add runs with different formatting
r1, _ := p.AddRun("Bold text. ")
r1.SetBold(boolPtr(true))

r2, _ := p.AddRun("Italic text. ")
r2.SetItalic(boolPtr(true))

r3, _ := p.AddRun("Underlined. ")
u := docx.UnderlineSingle()
r3.SetUnderline(&u)

// Font color (RGB)
r4, _ := p.AddRun("Red text. ")
red := docx.NewRGBColor(0xFF, 0x00, 0x00)
r4.Font().Color().SetRGB(&red)

// Font name
r5, _ := p.AddRun("Different font. ")
name := "Arial"
r5.Font().SetName(&name)

// Font size (uses Length type)
r6, _ := p.AddRun("Big text.")
size := docx.Pt(24)
r6.Font().SetSize(&size)

// Subscript / superscript
rSup, _ := p.AddRun("2")
rSup.Font().SetSuperscript(boolPtr(true))

// Highlight color
rH, _ := p.AddRun("Highlighted")
yellow := enum.WdColorIndexYellow
rH.Font().SetHighlightColor(&yellow)

// Helper for *bool
func boolPtr(v bool) *bool { return &v }
```

### Advanced Font Properties

All properties are available via `run.Font()`:

| Method | Description |
|--------|-------------|
| `SetBold` / `SetItalic` | Bold / Italic |
| `SetUnderline` | 17 underline styles |
| `SetStrike` / `SetDoubleStrike` | Strikethrough |
| `SetAllCaps` / `SetSmallCaps` | Capitalization |
| `SetShadow` / `SetEmboss` / `SetOutline` / `SetImprint` | Text effects |
| `SetSuperscript` / `SetSubscript` | Vertical position |
| `SetHidden` | Hidden text |
| `SetName` | Font family |
| `SetSize` | Font size (`*docx.Length`) |
| `SetHighlightColor` | Highlight (`*enum.WdColorIndex`) |
| `Color().SetRGB` | Text color (`*docx.RGBColor`) |
| `Color().SetThemeColor` | Theme color (`*enum.MsoThemeColorIndex`) |

All boolean properties use tri-state `*bool`: `nil` = inherit, `&true` = on, `&false` = off.

## Paragraph Formatting

```go
pf := p.ParagraphFormat()

// Alignment
center := enum.WdParagraphAlignmentCenter
pf.SetAlignment(&center)

// Indentation (in twips; 1 inch = 1440 twips)
indent := 720 // 0.5 inch
pf.SetLeftIndent(&indent)
pf.SetFirstLineIndent(&indent)

// Spacing before/after (twips)
space := 240
pf.SetSpaceBefore(&space)
pf.SetSpaceAfter(&space)

// Line spacing
ls := docx.LineSpacingMultiple(1.5) // 1.5x
pf.SetLineSpacing(&ls)

// Keep together / keep with next
pf.SetKeepTogether(boolPtr(true))
pf.SetKeepWithNext(boolPtr(true))

// Widow/orphan control
pf.SetWidowControl(boolPtr(true))
```

### Tab Stops

```go
tabs := p.ParagraphFormat().TabStops()
tabs.AddTabStop(
	docx.Inches(3).Twips(),     // position
	enum.WdTabAlignmentCenter,  // alignment
	enum.WdTabLeaderDots,       // leader
)
tabs.ClearAll()
```

## Tables

```go
// Create a 3x4 table (width computed automatically from page margins)
tbl, _ := doc.AddTable(3, 4)

// Access and fill cells
cell, _ := tbl.CellAt(0, 0)
cell.SetText("Header")

// Fill table data
headers := []string{"Name", "Age", "City", "Score"}
for i, h := range headers {
	c, _ := tbl.CellAt(0, i)
	c.SetText(h)
}

// Add a new row
row, _ := tbl.AddRow()

// Add a new column
col, _ := tbl.AddColumn(2000) // width in twips

// Merge cells (horizontal)
c1, _ := tbl.CellAt(0, 0)
c2, _ := tbl.CellAt(0, 2)
merged, _ := c1.Merge(c2) // merges columns 0-2 in row 0

// Vertical alignment
top := enum.WdCellVerticalAlignmentTop
cell.SetVerticalAlignment(&top)

// Cell width
cell.SetWidth(3000) // twips

// Row height
rows := tbl.Rows()
r, _ := rows.Get(0)
height := 400
r.SetHeight(&height)

// Table style
tbl.SetStyle(docx.StyleName("Table Grid"))

// Iterate rows and cells
for _, row := range tbl.Rows().Iter() {
	for _, cell := range row.Cells() {
		fmt.Print(cell.Text(), "\t")
	}
	fmt.Println()
}

// Nested table inside a cell
cell.AddTable(2, 2)
```

## Sections and Page Setup

```go
sections := doc.Sections()
sect, _ := sections.Get(0)

// Page size (twips)
w := docx.Inches(8.5).Twips()
h := docx.Inches(11).Twips()
sect.SetPageWidth(&w)
sect.SetPageHeight(&h)

// Orientation
sect.SetOrientation(enum.WdOrientationLandscape)

// Margins (twips; 1440 twips = 1 inch)
margin := 1440
sect.SetTopMargin(&margin)
sect.SetBottomMargin(&margin)
sect.SetLeftMargin(&margin)
sect.SetRightMargin(&margin)

// Add a new section break
doc.AddSection(enum.WdSectionStartNewPage)
```

## Headers and Footers

```go
sect, _ := doc.Sections().Get(0)

// Primary header/footer
header := sect.Header()
header.AddParagraph("Page Header")

footer := sect.Footer()
footer.AddParagraph("Page Footer")

// First-page header/footer
sect.SetDifferentFirstPageHeaderFooter(true)
firstHdr := sect.FirstPageHeader()
firstHdr.AddParagraph("Title Page Header")

// Even-page header/footer (requires settings toggle)
settings, _ := doc.Settings()
settings.SetOddAndEvenPagesHeaderFooter(true)
evenHdr := sect.EvenPageHeader()
evenHdr.AddParagraph("Even Page Header")

// Check if linked to previous section
isLinked := header.IsLinkedToPrevious()
```

## Images

```go
// Add inline picture from file
f, _ := os.Open("photo.png")
defer f.Close()

w := int64(docx.Inches(2).Emu())
h := int64(docx.Inches(1.5).Emu())
shape, _ := doc.AddPicture(f, &w, &h)

// Add picture to a specific run
run, _ := p.AddRun("")
shape, _ = run.AddPicture(f, &w, &h)

// Pass nil for native image dimensions
shape, _ = doc.AddPicture(f, nil, nil)
```

Supported formats: JPEG, PNG, GIF, BMP, TIFF.

## Styles

```go
styles, _ := doc.Styles()

// Check if style exists
if styles.Contains("Custom Style") {
	st, _ := styles.Get("Custom Style")
	fmt.Println(st.Name())
}

// Create a new paragraph style
st, _ := styles.AddStyle("My Style", enum.WdStyleTypeParagraph, false)
st.Font().SetBold(boolPtr(true))
center := enum.WdParagraphAlignmentCenter
st.ParagraphFormat().SetAlignment(&center)

// Set base (parent) style
normal, _ := styles.Get("Normal")
st.SetBaseStyle(normal)

// Get default style for a type
defaultPara, _ := styles.Default(enum.WdStyleTypeParagraph)

// Iterate all styles
for _, s := range styles.Iter() {
	name, _ := s.Name()
	fmt.Println(name, s.StyleID())
}

// Latent styles
ls := styles.LatentStyles()
fmt.Println(ls.Len(), "latent style exceptions")
```

## Comments

```go
p, _ := doc.AddParagraph("")
r1, _ := p.AddRun("This text has a comment.")

initials := "R"
comment, _ := doc.AddComment(
	[]*docx.Run{r1},           // runs to annotate
	"Review this section.",     // comment text
	"Reviewer",                 // author
	&initials,                  // initials (optional)
)

// Read comments
comments, _ := doc.Comments()
for _, c := range comments.Iter() {
	author, _ := c.Author()
	fmt.Printf("%s: %s\n", author, c.Text())
}
```

## Core Properties (Document Metadata)

```go
cp, _ := doc.CoreProperties()

cp.SetTitle("Annual Report 2025")
cp.SetAuthor("Jane Doe")
cp.SetSubject("Finance")
cp.SetKeywords("report, annual, 2025")
cp.SetCategory("Reports")
cp.SetComments("Auto-generated document")
cp.SetLanguage("en-US")
cp.SetRevision(1)
cp.SetCreated(time.Now())
cp.SetModified(time.Now())

// Read properties
fmt.Println(cp.Title(), "by", cp.Author())
```

## Text Replacement

Replaces text across the entire document: body, headers, footers, and comments. Works across run boundaries while preserving formatting.

```go
count, err := doc.ReplaceText("{{NAME}}", "John Smith")
// count = number of replacements made
```

## Table Replacement

Replaces a text placeholder with a structured table.

```go
td := docx.TableData{
	Rows: [][]string{
		{"Product", "Qty", "Price"},
		{"Widget A", "100", "$5.00"},
		{"Widget B", "200", "$3.50"},
	},
	Style: docx.StyleName("Table Grid"),
}

count, err := doc.ReplaceWithTable("{{PRICE_TABLE}}", td)
```

Table width is computed automatically from context (page margins, cell width, etc.).

## Content Replacement (Document Merging)

Replaces a text placeholder with the entire body of another `.docx` document. Handles images, hyperlinks, styles, numbering, footnotes, endnotes, and bookmarks with full relationship remapping.

```go
// Open source document
source, _ := docx.OpenFile("chapter1.docx")

// Basic usage (UseDestinationStyles — target styles take precedence)
count, err := doc.ReplaceWithContent("{{CHAPTER_1}}", docx.ContentData{
	Source: source,
})

// Keep source visual formatting
count, err = doc.ReplaceWithContent("{{CHAPTER_2}}", docx.ContentData{
	Source: source,
	Format: docx.KeepSourceFormatting,
})

// Copy conflicting styles with a unique suffix
count, err = doc.ReplaceWithContent("{{CHAPTER_3}}", docx.ContentData{
	Source: source,
	Format: docx.KeepSourceFormatting,
	Options: docx.ImportFormatOptions{
		ForceCopyStyles: true,
	},
})

// Copy only styles that differ
count, err = doc.ReplaceWithContent("{{CHAPTER_4}}", docx.ContentData{
	Source: source,
	Format: docx.KeepDifferentStyles,
})
```

### ImportFormatMode

| Mode | Behavior |
|------|----------|
| `UseDestinationStyles` | Target style wins on conflict. Source-only styles are copied. *(default)* |
| `KeepSourceFormatting` | Source formatting is expanded to direct attributes. With `ForceCopyStyles`, the source style is copied with a suffix (`_0`, `_1`, ...). |
| `KeepDifferentStyles` | Identical styles use the target; different styles are copied with suffix. |

### ImportFormatOptions

| Option | Description |
|--------|-------------|
| `ForceCopyStyles` | Copy conflicting styles with suffix instead of expanding. Only with `KeepSourceFormatting`. |
| `KeepSourceNumbering` | Keep source list numbering as separate definitions. |
| `IgnoreHeaderFooter` | Headers/footers always use destination styles regardless of mode. |
| `MergePastedLists` | Merge adjacent lists with compatible numbering into one continuous list. |

On error, the document body is automatically rolled back to its pre-call state.

## Units

The library provides a `Length` type for unit conversions (internally stored as EMU — English Metric Units):

```go
docx.Inches(1.5)   // 1.5 inches
docx.Cm(2.54)      // 2.54 centimeters
docx.Mm(25.4)      // 25.4 millimeters
docx.Pt(12)        // 12 points
docx.Twips(1440)   // 1440 twips = 1 inch
docx.Emu(914400)   // raw EMU

// Convert between units
l := docx.Inches(1)
l.Twips()   // 1440
l.Pt()      // 72.0
l.Cm()      // 2.54
l.Emu()     // 914400
```

Note: section margins, indents, and spacing use **twips** (1/20 of a point). Image dimensions use **EMU** (1/914400 of an inch).

## Colors

```go
// RGB from components
red := docx.NewRGBColor(0xFF, 0x00, 0x00)

// RGB from hex string
color, err := docx.RGBColorFromString("3C2F80")

// Get components
color.R() // red byte
color.G() // green byte
color.B() // blue byte

// String representation
color.String() // "3C2F80"
```

## Complete Example

```go
package main

import (
	"log"
	"os"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

func main() {
	doc, err := docx.New()
	if err != nil {
		log.Fatal(err)
	}

	// Document properties
	cp, _ := doc.CoreProperties()
	cp.SetTitle("Sales Report Q1")
	cp.SetAuthor("Finance Team")
	cp.SetCreated(time.Now())

	// Title
	doc.AddHeading("Sales Report Q1 2025", 0)

	// Introduction
	p, _ := doc.AddParagraph("")
	r1, _ := p.AddRun("Revenue target: ")
	r2, _ := p.AddRun("$1,000,000")
	r2.SetBold(boolPtr(true))
	green := docx.NewRGBColor(0x00, 0x80, 0x00)
	r2.Font().Color().SetRGB(&green)

	// Summary table
	doc.AddHeading("By Region", 1)
	tbl, _ := doc.AddTable(4, 3)
	data := [][]string{
		{"Region", "Revenue", "Target"},
		{"North", "$350,000", "$300,000"},
		{"South", "$280,000", "$250,000"},
		{"West", "$420,000", "$400,000"},
	}
	for r, row := range data {
		for c, val := range row {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(val)
		}
	}

	// Page break + second page
	doc.AddPageBreak()
	doc.AddHeading("Appendix", 1)
	doc.AddParagraph("Detailed breakdown available on request.")

	// Add comment
	lastP, _ := doc.AddParagraph("")
	run, _ := lastP.AddRun("Pending review.")
	initials := "FT"
	doc.AddComment([]*docx.Run{run}, "Needs CFO approval", "Finance Team", &initials)

	// Image (if available)
	if f, err := os.Open("chart.png"); err == nil {
		defer f.Close()
		w := int64(docx.Inches(4).Emu())
		h := int64(docx.Inches(3).Emu())
		doc.AddPicture(f, &w, &h)
	}

	// Save
	if err := doc.SaveFile("report.docx"); err != nil {
		log.Fatal(err)
	}
}

func boolPtr(v bool) *bool { return &v }
```

## Iterating Document Content

```go
// Paragraphs and tables in order
items, _ := doc.IterInnerContent()
for _, item := range items {
	if item.IsParagraph() {
		fmt.Println("P:", item.Paragraph().Text())
	}
	if item.IsTable() {
		fmt.Println("T:", item.Table().Rows().Len(), "rows")
	}
}

// Runs and hyperlinks within a paragraph
for _, item := range p.IterInnerContent() {
	if item.IsRun() {
		fmt.Print(item.Run().Text())
	}
	if item.IsHyperlink() {
		hl := item.Hyperlink()
		fmt.Printf("[%s](%s)", hl.Text(), hl.URL())
	}
}
```

## Dependencies

- [github.com/beevik/etree](https://github.com/beevik/etree) — XML processing
- No CGO required

## License

See [LICENSE](LICENSE).

---

# Русский

## go-docx

Библиотека на чистом Go для создания, чтения и редактирования документов Microsoft Word (.docx). API спроектирован по образцу [python-docx](https://python-docx.readthedocs.io/), обеспечивая те же удобные паттерны работы с документами с полной типобезопасностью Go.

## Возможности

- **Создание** новых пустых документов из встроенного шаблона
- **Открытие** существующих `.docx` файлов из пути, байтового среза или `io.ReaderAt`
- **Сохранение** в `io.Writer` или файл
- **Абзацы** — добавление, итерация, очистка; выравнивание, стиль, полное форматирование (отступы, интервалы, табуляция, контроль висячих строк)
- **Runs (текстовые фрагменты)** — жирный, курсив, подчёркивание (17 стилей), шрифт, размер, цвет, выделение, индексы, капитель, зачёркивание, тень, тиснение и др.
- **Таблицы** — создание и редактирование: строки, столбцы, объединение ячеек, ширина, вертикальное выравнивание, стиль таблицы
- **Разделы** — размер страницы, ориентация, поля, типы разрывов секций
- **Колонтитулы** — основные, первой страницы и чётных страниц; переключение чёт/нечёт через настройки
- **Изображения** — inline-картинки из `io.ReadSeeker`; поддержка JPEG, PNG, GIF, BMP, TIFF с автоопределением размеров
- **Гиперссылки** — чтение адресов, фрагментов и вложенных runs
- **Стили** — доступ, создание и редактирование стилей абзацев/символов/таблиц; перевод имён BabelFish (UI ↔ внутреннее); управление латентными стилями
- **Примечания** — добавление примечаний к диапазонам runs с автором и инициалами
- **Свойства документа** — автор, заголовок, тема, ключевые слова, категория, ревизия, даты создания/изменения
- **Нумерация** — определения списков импортируются при вставке контента
- **Замена текста** — `ReplaceText` по всему документу: тело, колонтитулы, примечания
- **Замена таблицей** — `ReplaceWithTable` заменяет текстовые метки структурированными таблицами
- **Замена контентом** — `ReplaceWithContent` заменяет метки содержимым другого `.docx` документа с полным ремаппингом связей, изображений, стилей, нумерации, сносок и дедупликацией изображений по SHA-256

## Установка

```bash
go get github.com/vortex/go-docx
```

Требуется **Go 1.25** или новее.

## Быстрый старт

```go
package main

import (
	"log"

	"github.com/vortex/go-docx/pkg/docx"
)

func main() {
	doc, err := docx.New()
	if err != nil {
		log.Fatal(err)
	}

	doc.AddHeading("Привет, мир!", 1)
	doc.AddParagraph("Создано с помощью go-docx.")

	if err := doc.SaveFile("hello.docx"); err != nil {
		log.Fatal(err)
	}
}
```

## Открытие и сохранение документов

```go
// Создать новый пустой документ
doc, err := docx.New()

// Открыть из файла
doc, err := docx.OpenFile("input.docx")

// Открыть из байтового среза
data, _ := os.ReadFile("input.docx")
doc, err := docx.OpenBytes(data)

// Открыть из io.ReaderAt
doc, err := docx.Open(reader, size)

// Сохранить в файл
err = doc.SaveFile("output.docx")

// Сохранить в io.Writer
var buf bytes.Buffer
err = doc.Save(&buf)
```

## Абзацы и заголовки

```go
// Заголовки (уровень 0 = Title, 1-9 = Heading N)
doc.AddHeading("Заголовок документа", 0)
doc.AddHeading("Глава 1", 1)
doc.AddHeading("Раздел 1.1", 2)

// Простой абзац
p, _ := doc.AddParagraph("Текст абзаца")

// Абзац со стилем
p, _ = doc.AddParagraph("Стилизованный текст", docx.StyleName("Heading 1"))

// Получить все абзацы
paras, _ := doc.Paragraphs()
for _, p := range paras {
	fmt.Println(p.Text())
}

// Разрыв страницы
doc.AddPageBreak()
```

## Текстовые фрагменты (Runs) и форматирование символов

Абзац состоит из одного или нескольких *runs* — последовательных фрагментов текста с одинаковым форматированием.

```go
p, _ := doc.AddParagraph("")

// Жирный
r1, _ := p.AddRun("Жирный текст. ")
r1.SetBold(boolPtr(true))

// Курсив
r2, _ := p.AddRun("Курсив. ")
r2.SetItalic(boolPtr(true))

// Подчёркивание
r3, _ := p.AddRun("Подчёркнутый. ")
u := docx.UnderlineSingle()
r3.SetUnderline(&u)

// Цвет шрифта (RGB)
r4, _ := p.AddRun("Красный текст. ")
red := docx.NewRGBColor(0xFF, 0x00, 0x00)
r4.Font().Color().SetRGB(&red)

// Имя шрифта
r5, _ := p.AddRun("Другой шрифт. ")
name := "Arial"
r5.Font().SetName(&name)

// Размер шрифта
r6, _ := p.AddRun("Крупный текст.")
size := docx.Pt(24)
r6.Font().SetSize(&size)

// Верхний/нижний индексы: H₂O
pChem, _ := doc.AddParagraph("")
pChem.AddRun("H")
rSub, _ := pChem.AddRun("2")
rSub.Font().SetSubscript(boolPtr(true))
pChem.AddRun("O")

func boolPtr(v bool) *bool { return &v }
```

### Все свойства шрифта

| Метод | Описание |
|-------|----------|
| `SetBold` / `SetItalic` | Жирный / Курсив |
| `SetUnderline` | 17 стилей подчёркивания |
| `SetStrike` / `SetDoubleStrike` | Зачёркивание |
| `SetAllCaps` / `SetSmallCaps` | Регистр |
| `SetShadow` / `SetEmboss` / `SetOutline` / `SetImprint` | Эффекты |
| `SetSuperscript` / `SetSubscript` | Верхний/нижний индекс |
| `SetHidden` | Скрытый текст |
| `SetName` | Семейство шрифта |
| `SetSize` | Размер (`*docx.Length`) |
| `SetHighlightColor` | Выделение (`*enum.WdColorIndex`) |
| `Color().SetRGB` | Цвет текста (`*docx.RGBColor`) |
| `Color().SetThemeColor` | Цвет темы (`*enum.MsoThemeColorIndex`) |

Все булевы свойства используют `*bool`: `nil` = наследовать, `&true` = вкл., `&false` = выкл.

## Форматирование абзаца

```go
pf := p.ParagraphFormat()

// Выравнивание
center := enum.WdParagraphAlignmentCenter
pf.SetAlignment(&center)

// Отступы (в twips; 1 дюйм = 1440 twips)
indent := 720 // 0.5 дюйма
pf.SetLeftIndent(&indent)
pf.SetFirstLineIndent(&indent)

// Интервал до/после (twips)
space := 240
pf.SetSpaceBefore(&space)
pf.SetSpaceAfter(&space)

// Межстрочный интервал
ls := docx.LineSpacingMultiple(1.5) // 1.5x
pf.SetLineSpacing(&ls)
```

### Табуляция

```go
tabs := p.ParagraphFormat().TabStops()
tabs.AddTabStop(
	docx.Inches(3).Twips(),     // позиция
	enum.WdTabAlignmentCenter,  // выравнивание
	enum.WdTabLeaderDots,       // заполнитель
)
```

## Таблицы

```go
// Создать таблицу 3×4 (ширина вычисляется автоматически из полей страницы)
tbl, _ := doc.AddTable(3, 4)

// Заполнить ячейки
cell, _ := tbl.CellAt(0, 0)
cell.SetText("Заголовок")

// Добавить строку / столбец
tbl.AddRow()
tbl.AddColumn(2000) // ширина в twips

// Объединить ячейки
c1, _ := tbl.CellAt(0, 0)
c2, _ := tbl.CellAt(0, 2)
merged, _ := c1.Merge(c2)

// Стиль таблицы
tbl.SetStyle(docx.StyleName("Table Grid"))

// Итерация
for _, row := range tbl.Rows().Iter() {
	for _, cell := range row.Cells() {
		fmt.Print(cell.Text(), "\t")
	}
	fmt.Println()
}
```

## Разделы и настройка страницы

```go
sect, _ := doc.Sections().Get(0)

// Размер страницы (twips)
w := docx.Inches(8.5).Twips()
h := docx.Inches(11).Twips()
sect.SetPageWidth(&w)
sect.SetPageHeight(&h)

// Ориентация
sect.SetOrientation(enum.WdOrientationLandscape)

// Поля (twips)
margin := 1440 // 1 дюйм
sect.SetTopMargin(&margin)
sect.SetBottomMargin(&margin)
sect.SetLeftMargin(&margin)
sect.SetRightMargin(&margin)

// Новый разрыв раздела
doc.AddSection(enum.WdSectionStartNewPage)
```

## Колонтитулы

```go
sect, _ := doc.Sections().Get(0)

// Основной колонтитул
header := sect.Header()
header.AddParagraph("Верхний колонтитул")

footer := sect.Footer()
footer.AddParagraph("Нижний колонтитул")

// Колонтитул первой страницы
sect.SetDifferentFirstPageHeaderFooter(true)
firstHdr := sect.FirstPageHeader()
firstHdr.AddParagraph("Титульная страница")

// Колонтитулы чётных страниц
settings, _ := doc.Settings()
settings.SetOddAndEvenPagesHeaderFooter(true)
evenHdr := sect.EvenPageHeader()
evenHdr.AddParagraph("Чётная страница")
```

## Изображения

```go
f, _ := os.Open("photo.png")
defer f.Close()

w := int64(docx.Inches(2).Emu())
h := int64(docx.Inches(1.5).Emu())
doc.AddPicture(f, &w, &h)

// Нативный размер изображения
doc.AddPicture(f, nil, nil)
```

Поддерживаемые форматы: JPEG, PNG, GIF, BMP, TIFF.

## Стили

```go
styles, _ := doc.Styles()

// Создать стиль
st, _ := styles.AddStyle("Мой стиль", enum.WdStyleTypeParagraph, false)
st.Font().SetBold(boolPtr(true))

// Базовый (родительский) стиль
normal, _ := styles.Get("Normal")
st.SetBaseStyle(normal)

// Получить стиль по имени
heading, _ := styles.Get("Heading 1")
```

## Примечания

```go
p, _ := doc.AddParagraph("")
r, _ := p.AddRun("Текст с примечанием.")

initials := "Р"
doc.AddComment([]*docx.Run{r}, "Проверить этот раздел.", "Рецензент", &initials)

// Чтение примечаний
comments, _ := doc.Comments()
for _, c := range comments.Iter() {
	author, _ := c.Author()
	fmt.Printf("%s: %s\n", author, c.Text())
}
```

## Свойства документа

```go
cp, _ := doc.CoreProperties()

cp.SetTitle("Годовой отчёт 2025")
cp.SetAuthor("Иванов И.И.")
cp.SetSubject("Финансы")
cp.SetKeywords("отчёт, годовой, 2025")
cp.SetCreated(time.Now())
cp.SetModified(time.Now())
```

## Замена текста

Заменяет текст по всему документу: тело, колонтитулы, примечания. Работает через границы runs, сохраняя форматирование.

```go
count, _ := doc.ReplaceText("{{ИМЯ}}", "Иванов Иван Иванович")
```

## Замена таблицей

```go
td := docx.TableData{
	Rows: [][]string{
		{"Товар", "Кол-во", "Цена"},
		{"Виджет А", "100", "500 руб."},
		{"Виджет Б", "200", "350 руб."},
	},
	Style: docx.StyleName("Table Grid"),
}

count, _ := doc.ReplaceWithTable("{{ТАБЛИЦА_ЦЕН}}", td)
```

## Замена контентом (слияние документов)

Заменяет текстовую метку содержимым другого `.docx` документа. Обрабатывает изображения, гиперссылки, стили, нумерацию, сноски и закладки с полным ремаппингом связей.

```go
source, _ := docx.OpenFile("chapter1.docx")

// Базовое использование (стили назначения имеют приоритет)
count, _ := doc.ReplaceWithContent("{{ГЛАВА_1}}", docx.ContentData{
	Source: source,
})

// Сохранить визуальное форматирование источника
count, _ = doc.ReplaceWithContent("{{ГЛАВА_2}}", docx.ContentData{
	Source: source,
	Format: docx.KeepSourceFormatting,
})

// Копировать конфликтующие стили с уникальным суффиксом
count, _ = doc.ReplaceWithContent("{{ГЛАВА_3}}", docx.ContentData{
	Source: source,
	Format: docx.KeepSourceFormatting,
	Options: docx.ImportFormatOptions{
		ForceCopyStyles: true,
	},
})
```

### Режимы ImportFormatMode

| Режим | Поведение |
|-------|-----------|
| `UseDestinationStyles` | Стиль назначения побеждает при конфликте. Стили только из источника копируются. *(по умолчанию)* |
| `KeepSourceFormatting` | Форматирование источника раскрывается в прямые атрибуты. С `ForceCopyStyles` — стиль копируется с суффиксом (`_0`, `_1`, ...). |
| `KeepDifferentStyles` | Идентичные стили используют назначение; различающиеся копируются с суффиксом. |

### Опции ImportFormatOptions

| Опция | Описание |
|-------|----------|
| `ForceCopyStyles` | Копировать конфликтующие стили с суффиксом. Только для `KeepSourceFormatting`. |
| `KeepSourceNumbering` | Сохранить нумерацию списков источника как отдельные определения. |
| `IgnoreHeaderFooter` | Колонтитулы всегда используют стили назначения. |
| `MergePastedLists` | Объединить смежные списки с совместимой нумерацией в один. |

При ошибке тело документа автоматически откатывается к состоянию до вызова.

## Единицы измерения

```go
docx.Inches(1.5)   // 1.5 дюйма
docx.Cm(2.54)      // 2.54 сантиметра
docx.Mm(25.4)      // 25.4 миллиметра
docx.Pt(12)        // 12 пунктов
docx.Twips(1440)   // 1440 twips = 1 дюйм
docx.Emu(914400)   // сырые EMU

// Конвертация
l := docx.Inches(1)
l.Twips()   // 1440
l.Pt()      // 72.0
l.Cm()      // 2.54
```

Поля разделов, отступы и интервалы — в **twips** (1/20 пункта). Размеры изображений — в **EMU** (1/914400 дюйма).

## Зависимости

- [github.com/beevik/etree](https://github.com/beevik/etree) — обработка XML
- CGO не требуется

## Лицензия

См. [LICENSE](LICENSE).
