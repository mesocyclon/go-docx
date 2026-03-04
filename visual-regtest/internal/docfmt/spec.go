package docfmt

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx"
)

// SpecReplace adds a specification line for text replacement:
//
//	▸ ~~old~~ (yellow, strikethrough, dark red)  →  new (green, bold, dark green)
func SpecReplace(doc *docx.Document, old, new string) {
	p, _ := doc.AddParagraph("")

	pfx, _ := p.AddRun("    ▸ ")
	_ = pfx.Font().Color().SetRGB(&ColorLightGray)

	rOld, _ := p.AddRun(old)
	SetHighlightYellow(rOld)
	_ = rOld.Font().SetStrike(boolPtr(true))
	_ = rOld.Font().Color().SetRGB(&ColorDarkRed)

	arr, _ := p.AddRun("   →   ")
	_ = arr.Font().Color().SetRGB(&ColorLightGray)

	rNew, _ := p.AddRun(new)
	SetHighlightGreen(rNew)
	_ = rNew.SetBold(boolPtr(true))
	_ = rNew.Font().Color().SetRGB(&ColorDarkGreen)
}

// SpecTable adds a specification line for table replacement:
//
//	▸ |<TAG>|  →  [TABLE: expected data] (green, bold)
func SpecTable(doc *docx.Document, tag, expectedDesc string) {
	p, _ := doc.AddParagraph("")

	// Prefix arrow
	pfx, _ := p.AddRun("    ▸ ")
	_ = pfx.Font().Color().SetRGB(&ColorLightGray)

	// Tag: yellow highlight
	rTag, _ := p.AddRun(tag)
	SetHighlightYellow(rTag)

	// Arrow
	arr, _ := p.AddRun("   →   ")
	_ = arr.Font().Color().SetRGB(&ColorLightGray)

	// Expected: green highlight + bold
	rExp, _ := p.AddRun("[TABLE: " + expectedDesc + "]")
	SetHighlightGreen(rExp)
	_ = rExp.SetBold(boolPtr(true))
	_ = rExp.Font().Color().SetRGB(&ColorDarkGreen)
}

// SpecContent adds a specification line for content replacement:
//
//	▸ (<TAG>)  →  [СОДЕРЖИМОЕ: desc] (green, bold)
func SpecContent(doc *docx.Document, tag, desc string) {
	specWithPrefix(doc, tag, "СОДЕРЖИМОЕ", desc)
}

func specWithPrefix(doc *docx.Document, tag, prefix, desc string) {
	p, _ := doc.AddParagraph("")

	pfxRun, _ := p.AddRun(fmt.Sprintf("    ▸ %s → ", tag))
	_ = pfxRun.Font().Color().SetRGB(&ColorGray)

	rDesc, _ := p.AddRun(fmt.Sprintf("[%s: %s]", prefix, desc))
	SetHighlightGreen(rDesc)
	_ = rDesc.SetBold(boolPtr(true))
	_ = rDesc.Font().Color().SetRGB(&ColorDarkGreen)
}
