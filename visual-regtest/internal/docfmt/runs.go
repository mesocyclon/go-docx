package docfmt

import (
	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// SetHighlightYellow applies yellow highlight to a run.
func SetHighlightYellow(r *docx.Run) {
	hl := enum.WdColorIndexYellow
	_ = r.Font().SetHighlightColor(&hl)
}

// SetHighlightGreen applies bright-green highlight to a run.
func SetHighlightGreen(r *docx.Run) {
	hl := enum.WdColorIndexBrightGreen
	_ = r.Font().SetHighlightColor(&hl)
}

// AddPlain adds a plain (unformatted) run to a paragraph.
func AddPlain(p *docx.Paragraph, text string) {
	_, _ = p.AddRun(text)
}

// AddBold adds a bold run to a paragraph.
func AddBold(p *docx.Paragraph, text string) {
	r, _ := p.AddRun(text)
	_ = r.SetBold(boolPtr(true))
}

// AddHighlighted adds a yellow-highlighted run to a paragraph.
func AddHighlighted(p *docx.Paragraph, text string) {
	r, _ := p.AddRun(text)
	SetHighlightYellow(r)
}

// AddGreen adds a green-highlighted run to a paragraph.
func AddGreen(p *docx.Paragraph, text string) {
	r, _ := p.AddRun(text)
	SetHighlightGreen(r)
}

func boolPtr(v bool) *bool { return &v }
