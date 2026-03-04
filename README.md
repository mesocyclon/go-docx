# go-docx

A pure Go library for creating, reading, and modifying Microsoft Word (.docx) documents. The API is modeled after [python-docx](https://python-docx.readthedocs.io/), bringing the same intuitive document manipulation patterns to Go with full type safety.

## Features

- **Create** documents from scratch or from a built-in default template
- **Open** existing `.docx` files from a path, byte slice, or `io.ReaderAt`
- **Save** to a writer or file path
- **Paragraphs** — add, iterate, clear; set alignment, style, and full paragraph formatting (indentation, spacing, tab stops, keep-together, widow/orphan control)
- **Runs** — text runs with character formatting: bold, italic, underline (16 styles), font name, font size, color, highlighting, subscript/superscript, all-caps, small-caps, strikethrough, shadow, emboss, and 20+ more properties
- **Tables** — create and modify tables with rows, columns, cell merging, cell width, vertical alignment, table style, and shading
- **Sections** — page size, orientation, margins, header/footer references, section break types, columns
- **Headers & Footers** — default, first-page, and even-page variants; odd/even page toggle via settings
- **Images** — inline pictures from `io.ReadSeeker` or pre-built `ImagePart`; supports JPEG, PNG, GIF, BMP, and TIFF with automatic dimension detection
- **Hyperlinks** — read addresses, fragments, and contained runs
- **Styles** — access, create, and modify paragraph/character/table styles; BabelFish name translation (UI ↔ internal); latent style management
- **Comments** — add comments anchored to run ranges with author and initials; iterate and modify
- **Core Properties** — author, title, subject, keywords, category, comments, revision, created/modified dates
- **Document Settings** — odd/even page headers, element layout control
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

### Create a New Document

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
    doc.AddParagraph("This is a paragraph created with go-docx.")

    if err := doc.SaveFile("hello.docx"); err != nil {
        log.Fatal(err)
    }
}
```

### Open and Modify an Existing Document

```go
doc, err := docx.OpenFile("existing.docx")
if err != nil {
    log.Fatal(err)
}

// Replace text everywhere: body, headers, footers, comments
count, err := doc.ReplaceText("{{NAME}}", "Alice")
if err != nil {
    log.Fatal(err)
}
log.Printf("Replaced %d occurrences", count)

if err := doc.SaveFile("modified.docx"); err != nil {
    log.Fatal(err)
}
```

### Add a Table

```go
doc, _ := docx.New()

table, err := doc.AddTable(3, 2, docx.StyleName("Table Grid"))
if err != nil {
    log.Fatal(err)
}

// Fill cell text
rows := table.Rows()
for i, row := range rows {
    cells := row.Cells()
    cells[0].SetText(fmt.Sprintf("Row %d", i+1))
    cells[1].SetText(fmt.Sprintf("Value %d", (i+1)*10))
}

doc.SaveFile("table.docx")
```

### Add an Image

```go
doc, _ := docx.New()

f, _ := os.Open("photo.jpg")
defer f.Close()

width := docx.Inches(4).Emu()
_, err := doc.AddPicture(f, &width, nil) // proportional height
if err != nil {
    log.Fatal(err)
}

doc.SaveFile("with_image.docx")
```

### Character Formatting

```go
doc, _ := docx.New()

para, _ := doc.AddParagraph("")
run, _ := para.AddRun("Bold and red text")

font := run.Font()
bold := true
font.SetBold(&bold)
font.SetColor(docx.NewRGBColor(0xFF, 0x00, 0x00))
font.SetSize(docx.Pt(14))

doc.SaveFile("formatted.docx")
```

### Replace Placeholder with Another Document

```go
template, _ := docx.OpenFile("template.docx")
source, _ := docx.OpenFile("content.docx")

count, err := template.ReplaceWithContent("(<insert_here>)", docx.ContentData{
    Source: source,
})
if err != nil {
    log.Fatal(err)
}

template.SaveFile("filled.docx")
```

### Replace Placeholder with a Table

```go
doc, _ := docx.OpenFile("template.docx")

td := docx.TableData{
    Rows: [][]string{
        {"Name", "Score"},
        {"Alice", "95"},
        {"Bob", "87"},
    },
    Style: docx.StyleName("Table Grid"),
}

count, err := doc.ReplaceWithTable("{{RESULTS}}", td)
if err != nil {
    log.Fatal(err)
}

doc.SaveFile("with_table.docx")
```

## Units

The library uses English Metric Units (EMU) internally and provides convenience constructors:

```go
docx.Inches(1.5)   // 1.5 inches → Length
docx.Cm(2.54)      // centimeters → Length
docx.Mm(25.4)      // millimeters → Length
docx.Pt(12)        // points → Length
docx.Twips(240)    // twips (1/20 pt) → Length
docx.Emu(914400)   // raw EMU → Length
```

Each `Length` value can be converted back:

```go
l := docx.Inches(1)
l.Emu()    // 914400
l.Twips()  // 1440
l.Pt()     // 72.0
l.Cm()     // 2.54
```

## Architecture

The library is organized into four layers, matching the OOXML specification structure:

```
pkg/docx/              ← Public API: Document, Paragraph, Run, Table, ...
pkg/docx/oxml/         ← XML element types (CT_Document, CT_P, CT_R, ...)
pkg/docx/parts/        ← Document parts (DocumentPart, StylesPart, ImagePart, ...)
pkg/docx/opc/          ← OPC package layer (ZIP I/O, relationships, content types)
pkg/docx/enum/         ← OOXML enumerations (WdParagraphAlignment, WdBreakType, ...)
pkg/docx/image/        ← Image format readers (JPEG, PNG, GIF, BMP, TIFF)
```

The `oxml` types are partially generated from YAML schemas via a code generator:

```bash
go run ./cmd/codegen -schema ./schema/ -out ./pkg/docx/oxml/
```

Generated files follow the `zz_gen_*.go` naming convention. Hand-written `*_custom.go` files extend generated types with business logic.

## Concurrency

A single `Document` and all objects derived from it (paragraphs, runs, tables, sections, etc.) must be accessed from one goroutine at a time. Independent `Document` instances may be used concurrently without synchronization.

## Visual Regression Testing

The `visual-regtest/` directory contains a Docker-based visual regression test suite that verifies roundtrip fidelity. It opens a `.docx` file and saves it back, then compares page images rendered by LibreOffice. See [`visual-regtest/README.md`](visual-regtest/README.md) for usage.

## Dependencies

- [github.com/beevik/etree](https://github.com/beevik/etree) — XML tree manipulation
- [gopkg.in/yaml.v3](https://gopkg.in/yaml.v3) — YAML parsing (code generator only)

## License

MIT — see [LICENSE](LICENSE) for details.