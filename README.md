# go-docx

A pure Go library for creating, reading, and modifying Microsoft Word (.docx) documents. The API is modeled after [python-docx](https://python-docx.readthedocs.io/), bringing the same intuitive document manipulation patterns to Go with full type safety.

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