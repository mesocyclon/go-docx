package docfmt

import (
	"log"

	"github.com/vortex/go-docx/pkg/docx"
)

// FillTable fills a table with the given cell values.
func FillTable(tbl *docx.Table, data [][]string) {
	for r, row := range data {
		for c, val := range row {
			cell, err := tbl.CellAt(r, c)
			if err != nil {
				log.Fatalf("CellAt(%d,%d): %v", r, c, err)
			}
			cell.SetText(val)
		}
	}
}

// SetCell sets the text of a single table cell.
func SetCell(tbl *docx.Table, row, col int, text string) {
	cell, err := tbl.CellAt(row, col)
	if err != nil {
		log.Fatalf("CellAt(%d,%d): %v", row, col, err)
	}
	cell.SetText(text)
}
