package docfmt

import (
	"log"

	"github.com/vortex/go-docx/pkg/docx"
)

// Heading adds a heading paragraph to the document.
func Heading(doc *docx.Document, text string, level int) {
	if _, err := doc.AddHeading(text, level); err != nil {
		log.Fatalf("AddHeading(%q): %v", text, err)
	}
}

// Para adds a plain text paragraph to the document.
func Para(doc *docx.Document, text string) {
	if _, err := doc.AddParagraph(text); err != nil {
		log.Fatalf("AddParagraph: %v", err)
	}
}

// Note adds a gray-colored description/instruction line.
func Note(doc *docx.Document, text string) {
	p, _ := doc.AddParagraph("")
	r, _ := p.AddRun(text)
	_ = r.Font().Color().SetRGB(&ColorGray)
}

// TagPara builds a paragraph: prefix + highlighted(tag) + suffix.
// The tag run gets yellow highlight — visually shows what will be replaced.
func TagPara(doc *docx.Document, prefix, tag, suffix string) {
	p, _ := doc.AddParagraph("")
	if prefix != "" {
		AddPlain(p, prefix)
	}
	AddHighlighted(p, tag)
	if suffix != "" {
		AddPlain(p, suffix)
	}
}

// ParagraphAdder is satisfied by *docx.Document, headers, footers, etc.
type ParagraphAdder interface {
	AddParagraph(text string, style ...docx.StyleRef) (*docx.Paragraph, error)
}

// BuildHighlightedParagraph adds a paragraph with yellow-highlighted text
// to a header, footer, or any ParagraphAdder.
func BuildHighlightedParagraph(hf ParagraphAdder, text string) {
	p, err := hf.AddParagraph("")
	if err != nil {
		log.Fatalf("AddParagraph in header/footer: %v", err)
	}
	r, _ := p.AddRun(text)
	SetHighlightYellow(r)
}
