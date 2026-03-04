package main

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// 14 — Basic table
func genTableBasic() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Basic Table (3x3)", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(3, 3)
	if err != nil {
		return nil, err
	}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			cell, err := tbl.CellAt(r, c)
			if err != nil {
				return nil, err
			}
			cell.SetText(fmt.Sprintf("Row %d, Col %d", r+1, c+1))
		}
	}

	if _, err := doc.AddParagraph("Paragraph after the table."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 15 — Table with merged cells
func genTableMergedCells() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table with Merged Cells", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(4, 4)
	if err != nil {
		return nil, err
	}
	for r := 0; r < 4; r++ {
		for c := 0; c < 4; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("(%d,%d)", r, c))
		}
	}

	c00, _ := tbl.CellAt(0, 0)
	c01, _ := tbl.CellAt(0, 1)
	merged, err := c00.Merge(c01)
	if err != nil {
		return nil, fmt.Errorf("h-merge: %w", err)
	}
	merged.SetText("Merged (0,0)-(0,1)")

	c13, _ := tbl.CellAt(1, 3)
	c33, _ := tbl.CellAt(3, 3)
	merged2, err := c13.Merge(c33)
	if err != nil {
		return nil, fmt.Errorf("v-merge: %w", err)
	}
	merged2.SetText("Merged (1,3)-(3,3)")

	c20, _ := tbl.CellAt(2, 0)
	c31, _ := tbl.CellAt(3, 1)
	merged3, err := c20.Merge(c31)
	if err != nil {
		return nil, fmt.Errorf("block-merge: %w", err)
	}
	merged3.SetText("Block merge (2,0)-(3,1)")

	return doc, nil
}

// 16 — Table alignment (left, center, right)
func genTableAlignment() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Alignment", 1); err != nil {
		return nil, err
	}

	aligns := []struct {
		label string
		val   enum.WdTableAlignment
	}{
		{"Left", enum.WdTableAlignmentLeft},
		{"Center", enum.WdTableAlignmentCenter},
		{"Right", enum.WdTableAlignmentRight},
	}

	for _, a := range aligns {
		if _, err := doc.AddParagraph(a.label + " aligned table:"); err != nil {
			return nil, err
		}
		tbl, err := doc.AddTable(2, 2)
		if err != nil {
			return nil, err
		}
		if err := tbl.SetAlignment(&a.val); err != nil {
			return nil, err
		}
		if err := tbl.SetAutofit(false); err != nil {
			return nil, err
		}
		for r := 0; r < 2; r++ {
			for c := 0; c < 2; c++ {
				cell, _ := tbl.CellAt(r, c)
				cell.SetText(fmt.Sprintf("%s R%dC%d", a.label, r, c))
			}
		}
	}
	return doc, nil
}

// 17 — Nested table (table inside a cell)
func genTableNested() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Nested Table", 1); err != nil {
		return nil, err
	}

	outer, err := doc.AddTable(2, 2)
	if err != nil {
		return nil, err
	}
	c00, _ := outer.CellAt(0, 0)
	c00.SetText("Outer (0,0) — has nested table below")

	inner, err := c00.AddTable(2, 2)
	if err != nil {
		return nil, fmt.Errorf("nested table: %w", err)
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := inner.CellAt(r, c)
			cell.SetText(fmt.Sprintf("Inner %d,%d", r, c))
		}
	}

	c01, _ := outer.CellAt(0, 1)
	c01.SetText("Outer (0,1)")
	c10, _ := outer.CellAt(1, 0)
	c10.SetText("Outer (1,0)")
	c11, _ := outer.CellAt(1, 1)
	c11.SetText("Outer (1,1)")

	return doc, nil
}

// 18 — Table cell vertical alignment
func genTableCellVerticalAlign() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Cell Vertical Alignment", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(1, 3)
	if err != nil {
		return nil, err
	}

	rows := tbl.Rows()
	row, _ := rows.Get(0)
	if err := row.SetHeight(IntPtr(1440)); err != nil {
		return nil, err
	}
	rule := enum.WdRowHeightRuleExactly
	if err := row.SetHeightRule(&rule); err != nil {
		return nil, err
	}

	valigns := []struct {
		label string
		val   enum.WdCellVerticalAlignment
	}{
		{"Top", enum.WdCellVerticalAlignmentTop},
		{"Center", enum.WdCellVerticalAlignmentCenter},
		{"Bottom", enum.WdCellVerticalAlignmentBottom},
	}

	for i, va := range valigns {
		cell, _ := tbl.CellAt(0, i)
		cell.SetText(va.label)
		if err := cell.SetVerticalAlignment(&va.val); err != nil {
			return nil, err
		}
	}

	return doc, nil
}

// 19 — Table: AddRow / AddColumn
func genTableAddRowCol() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table — AddRow / AddColumn", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 2)
	if err != nil {
		return nil, err
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("Original R%dC%d", r, c))
		}
	}

	newRow, err := tbl.AddRow()
	if err != nil {
		return nil, err
	}
	cells := newRow.Cells()
	for i, c := range cells {
		c.SetText(fmt.Sprintf("New row C%d", i))
	}

	if _, err := tbl.AddColumn(1440); err != nil {
		return nil, err
	}
	for r := 0; r < 3; r++ {
		cell, err := tbl.CellAt(r, 2)
		if err != nil {
			continue
		}
		cell.SetText(fmt.Sprintf("New col R%d", r))
	}

	return doc, nil
}

// 34 — Table row height
func genTableRowHeight() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Row Heights", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(3, 2)
	if err != nil {
		return nil, err
	}

	heights := []struct {
		twips int
		rule  enum.WdRowHeightRule
		label string
	}{
		{360, enum.WdRowHeightRuleExactly, "Exact 0.25\""},
		{720, enum.WdRowHeightRuleAtLeast, "AtLeast 0.5\""},
		{1440, enum.WdRowHeightRuleExactly, "Exact 1.0\""},
	}

	rows := tbl.Rows()
	for i, h := range heights {
		row, _ := rows.Get(i)
		if err := row.SetHeight(IntPtr(h.twips)); err != nil {
			return nil, err
		}
		if err := row.SetHeightRule(&h.rule); err != nil {
			return nil, err
		}
		cell, _ := tbl.CellAt(i, 0)
		cell.SetText(h.label)
		cell2, _ := tbl.CellAt(i, 1)
		cell2.SetText(fmt.Sprintf("%d twips", h.twips))
	}

	return doc, nil
}

// 35 — Table column width
func genTableColumnWidth() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Column Widths", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 3)
	if err != nil {
		return nil, err
	}
	if err := tbl.SetAutofit(false); err != nil {
		return nil, err
	}

	cols, _ := tbl.Columns()
	widths := []int{1440, 2880, 4320}
	for i, w := range widths {
		col, _ := cols.Get(i)
		if err := col.SetWidth(IntPtr(w)); err != nil {
			return nil, err
		}
	}

	cell00, _ := tbl.CellAt(0, 0)
	cell00.SetText("1 inch")
	cell01, _ := tbl.CellAt(0, 1)
	cell01.SetText("2 inches")
	cell02, _ := tbl.CellAt(0, 2)
	cell02.SetText("3 inches")
	cell10, _ := tbl.CellAt(1, 0)
	cell10.SetText("narrow")
	cell11, _ := tbl.CellAt(1, 1)
	cell11.SetText("medium")
	cell12, _ := tbl.CellAt(1, 2)
	cell12.SetText("wide")

	return doc, nil
}

// 41 — Table cell SetText
func genTableCellSetText() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Cell SetText", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 2)
	if err != nil {
		return nil, err
	}

	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("First (%d,%d)", r, c))
		}
	}

	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("Replaced (%d,%d)", r, c))
		}
	}

	return doc, nil
}

// 42 — Table with style
func genTableStyle() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table with Style", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(4, 3, docx.StyleName("Table Grid"))
	if err != nil {
		return nil, err
	}
	headers := []string{"Product", "Quantity", "Price"}
	for i, h := range headers {
		cell, _ := tbl.CellAt(0, i)
		cell.SetText(h)
	}
	data := [][]string{
		{"Widget A", "100", "$5.99"},
		{"Widget B", "250", "$3.49"},
		{"Widget C", "50", "$12.00"},
	}
	for r, row := range data {
		for c, val := range row {
			cell, _ := tbl.CellAt(r+1, c)
			cell.SetText(val)
		}
	}

	return doc, nil
}

// 43 — Table bidi (RTL direction)
func genTableBidi() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table BiDi (RTL)", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 3)
	if err != nil {
		return nil, err
	}
	if err := tbl.SetTableDirection(BoolPtr(true)); err != nil {
		return nil, err
	}

	for r := 0; r < 2; r++ {
		for c := 0; c < 3; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("R%dC%d", r, c))
		}
	}

	if _, err := doc.AddParagraph("Table above has RTL (bidiVisual) direction."); err != nil {
		return nil, err
	}
	return doc, nil
}
