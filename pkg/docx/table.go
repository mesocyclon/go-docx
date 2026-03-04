package docx

import (
	"fmt"
	"strings"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Table is a proxy object wrapping a <w:tbl> element.
//
// Mirrors Python Table(StoryChild).
type Table struct {
	tbl  *oxml.CT_Tbl
	part *parts.StoryPart
}

// newTable creates a new Table proxy.
func newTable(tbl *oxml.CT_Tbl, part *parts.StoryPart) *Table {
	return &Table{tbl: tbl, part: part}
}

// AddColumn adds a new column with the given width (twips) rightmost.
//
// Mirrors Python Table.add_column.
func (t *Table) AddColumn(widthTwips int) (*Column, error) {
	grid, err := t.tbl.TblGrid()
	if err != nil {
		return nil, fmt.Errorf("docx: getting table grid: %w", err)
	}
	gridCol := grid.AddGridCol()
	w := widthTwips
	if err := gridCol.SetW(&w); err != nil {
		return nil, err
	}
	for _, tr := range t.tbl.TrList() {
		tc := tr.AddTc()
		if err := tc.SetWidthTwips(widthTwips); err != nil {
			return nil, err
		}
	}
	return &Column{gridCol: gridCol, table: t}, nil
}

// AddRow adds a new row at the bottom of the table.
//
// Mirrors Python Table.add_row.
func (t *Table) AddRow() (*Row, error) {
	tr := t.tbl.AddTr()
	grid, err := t.tbl.TblGrid()
	if err != nil {
		return nil, fmt.Errorf("docx: getting table grid: %w", err)
	}
	for _, gc := range grid.GridColList() {
		tc := tr.AddTc()
		w, err := gc.W()
		if err == nil && w != nil {
			if err := tc.SetWidthTwips(*w); err != nil {
				return nil, fmt.Errorf("docx: setting cell width in new row: %w", err)
			}
		}
	}
	return &Row{tr: tr, table: t}, nil
}

// Alignment returns the table alignment, or nil if inherited.
func (t *Table) Alignment() (*enum.WdTableAlignment, error) {
	return t.tbl.AlignmentVal()
}

// SetAlignment sets the table alignment. Passing nil removes it.
func (t *Table) SetAlignment(v *enum.WdTableAlignment) error {
	return t.tbl.SetAlignmentVal(v)
}

// Autofit returns true if column widths can be automatically adjusted.
func (t *Table) Autofit() (bool, error) {
	return t.tbl.Autofit()
}

// SetAutofit sets the autofit property.
func (t *Table) SetAutofit(v bool) error {
	return t.tbl.SetAutofit(v)
}

// CellAt returns the cell at (row_idx, col_idx). (0, 0) is top-left.
//
// Mirrors Python Table.cell.
func (t *Table) CellAt(rowIdx, colIdx int) (*Cell, error) {
	cells, err := t.cells()
	if err != nil {
		return nil, err
	}
	colCount, err := t.columnCount()
	if err != nil {
		return nil, err
	}
	idx := colIdx + (rowIdx * colCount)
	if idx < 0 || idx >= len(cells) {
		return nil, fmt.Errorf("docx: cell index (%d, %d) out of range", rowIdx, colIdx)
	}
	return cells[idx], nil
}

// ColumnCells returns a slice of cells in the column at colIdx.
func (t *Table) ColumnCells(colIdx int) ([]*Cell, error) {
	cells, err := t.cells()
	if err != nil {
		return nil, err
	}
	colCount, err := t.columnCount()
	if err != nil {
		return nil, err
	}
	var result []*Cell
	for i := colIdx; i < len(cells); i += colCount {
		result = append(result, cells[i])
	}
	return result, nil
}

// Columns returns the Columns sequence for this table.
func (t *Table) Columns() (*Columns, error) {
	grid, err := t.tbl.TblGrid()
	if err != nil {
		return nil, err
	}
	return &Columns{grid: grid, table: t}, nil
}

// RowCells returns a slice of cells in the row at rowIdx.
func (t *Table) RowCells(rowIdx int) ([]*Cell, error) {
	cells, err := t.cells()
	if err != nil {
		return nil, err
	}
	colCount, err := t.columnCount()
	if err != nil {
		return nil, err
	}
	start := rowIdx * colCount
	end := start + colCount
	if start < 0 || end > len(cells) {
		return nil, fmt.Errorf("docx: row index [%d] out of range", rowIdx)
	}
	return cells[start:end], nil
}

// Rows returns the Rows sequence for this table.
func (t *Table) Rows() *Rows {
	return &Rows{tbl: t.tbl, table: t}
}

// Style returns the table style.
func (t *Table) Style() (*oxml.CT_Style, error) {
	styleVal, err := t.tbl.TblStyleVal()
	if err != nil {
		return nil, err
	}
	if styleVal == "" {
		return t.part.GetStyle(nil, enum.WdStyleTypeTable)
	}
	return t.part.GetStyle(&styleVal, enum.WdStyleTypeTable)
}

// SetStyle sets the table style.
func (t *Table) SetStyle(style StyleRef) error {
	return t.setStyleRaw(resolveStyleRef([]StyleRef{style}))
}

// setStyleRaw passes the raw style value to the parts layer.
func (t *Table) setStyleRaw(raw any) error {
	styleID, err := t.part.GetStyleID(raw, enum.WdStyleTypeTable)
	if err != nil {
		return err
	}
	if styleID == nil {
		return t.tbl.SetTblStyleVal("")
	}
	return t.tbl.SetTblStyleVal(*styleID)
}

// TableDirection returns the cell-ordering direction, or nil if inherited.
func (t *Table) TableDirection() (*bool, error) {
	return t.tbl.BidiVisualVal()
}

// SetTableDirection sets the cell-ordering direction.
func (t *Table) SetTableDirection(v *bool) error {
	return t.tbl.SetBidiVisualVal(v)
}

// ReplaceText replaces all occurrences of old with new in every cell of
// this table, recursively (including nested tables). Each physical cell
// is processed exactly once via IterTcs() (no duplicates from merged cells).
// Returns the total number of replacements performed.
func (t *Table) ReplaceText(old, new string) int {
	count := 0
	for _, tc := range t.tbl.IterTcs() {
		bic := newBlockItemContainer(tc.RawElement(), t.part)
		count += bic.ReplaceText(old, new)
	}
	return count
}

// CT_Tbl returns the underlying oxml element.
func (t *Table) CT_Tbl() *oxml.CT_Tbl { return t.tbl }

// cells builds the flat list of cells, handling gridSpan and vMerge.
// EXACT copy of Python Table._cells algorithm.
func (t *Table) cells() ([]*Cell, error) {
	colCount, err := t.columnCount()
	if err != nil {
		return nil, err
	}
	var cells []*Cell
	for _, tc := range t.tbl.IterTcs() {
		gridSpan, err := tc.GridSpanVal()
		if err != nil {
			gridSpan = 1
		}
		vMerge := tc.VMergeVal()
		for gsi := 0; gsi < gridSpan; gsi++ {
			if vMerge != nil && *vMerge == "continue" {
				// Reference cell from row above
				aboveIdx := len(cells) - colCount
				if aboveIdx >= 0 && aboveIdx < len(cells) {
					cells = append(cells, cells[aboveIdx])
				} else {
					cells = append(cells, newCell(tc, t))
				}
			} else if gsi > 0 {
				// Same cell for horizontal span
				cells = append(cells, cells[len(cells)-1])
			} else {
				cells = append(cells, newCell(tc, t))
			}
		}
	}
	return cells, nil
}

func (t *Table) columnCount() (int, error) {
	return t.tbl.ColCount()
}

// --------------------------------------------------------------------------
// Cell
// --------------------------------------------------------------------------

// Cell is a proxy for a <w:tc> table cell element.
//
// Mirrors Python _Cell(BlockItemContainer).
type Cell struct {
	BlockItemContainer
	tc    *oxml.CT_Tc
	table *Table
}

// newCell creates a new Cell proxy.
func newCell(tc *oxml.CT_Tc, table *Table) *Cell {
	return &Cell{
		BlockItemContainer: newBlockItemContainer(tc.RawElement(), table.part),
		tc:                 tc,
		table:              table,
	}
}

// AddTable adds a table to this cell and adds a trailing empty paragraph.
//
// Mirrors Python _Cell.add_table.
func (c *Cell) AddTable(rows, cols int) (*Table, error) {
	width := 1440 // default Inches(1) = 1440 twips
	w, err := c.tc.WidthTwips()
	if err == nil && w != nil {
		width = *w
	}
	tbl, err := c.BlockItemContainer.AddTable(rows, cols, width)
	if err != nil {
		return nil, err
	}
	// OOXML requires at least one paragraph in a cell after the table.
	if _, err := c.BlockItemContainer.AddParagraph(""); err != nil {
		return nil, fmt.Errorf("docx: adding trailing paragraph in cell: %w", err)
	}
	return tbl, nil
}

// GridSpan returns the number of grid columns this cell spans.
func (c *Cell) GridSpan() int {
	v, err := c.tc.GridSpanVal()
	if err != nil {
		return 1
	}
	return v
}

// Merge merges the rectangular region from this cell to other and returns the merged cell.
//
// Mirrors Python _Cell.merge.
func (c *Cell) Merge(other *Cell) (*Cell, error) {
	merged, err := c.tc.Merge(other.tc)
	if err != nil {
		return nil, fmt.Errorf("docx: merging cells: %w", err)
	}
	return newCell(merged, c.table), nil
}

// Text returns the text content of this cell, paragraphs joined by newlines.
func (c *Cell) Text() string {
	paras := c.Paragraphs()
	texts := make([]string, len(paras))
	for i, p := range paras {
		texts[i] = p.Text()
	}
	return strings.Join(texts, "\n")
}

// SetText replaces all cell content with a single paragraph containing text.
//
// Mirrors Python _Cell.text setter:
//
//	tc.clear_content()
//	p = tc.add_p()
//	r = p.add_r()
//	r.text = text
func (c *Cell) SetText(text string) {
	c.tc.ClearContent()
	p := c.tc.AddP()
	r := p.AddR()
	r.SetRunText(text)
}

// VerticalAlignment returns the cell vertical alignment, or nil if inherited.
func (c *Cell) VerticalAlignment() (*enum.WdCellVerticalAlignment, error) {
	return c.tc.VAlignVal()
}

// SetVerticalAlignment sets the cell vertical alignment.
func (c *Cell) SetVerticalAlignment(v *enum.WdCellVerticalAlignment) error {
	return c.tc.SetVAlignVal(v)
}

// Width returns the cell width in twips, or nil if not set.
func (c *Cell) Width() (*int, error) {
	return c.tc.WidthTwips()
}

// SetWidth sets the cell width in twips.
func (c *Cell) SetWidth(twips int) error {
	return c.tc.SetWidthTwips(twips)
}

// --------------------------------------------------------------------------
// Row
// --------------------------------------------------------------------------

// Row is a proxy for a <w:tr> table row element.
//
// Mirrors Python _Row(Parented).
type Row struct {
	tr    *oxml.CT_Row
	table *Table
}

// Cells returns the cells in this row, expanding horizontal and vertical spans.
//
// Mirrors Python _Row.cells.
func (r *Row) Cells() []*Cell {
	var cells []*Cell
	for _, tc := range r.tr.TcList() {
		gridSpan, err := tc.GridSpanVal()
		if err != nil {
			gridSpan = 1
		}
		vMerge := tc.VMergeVal()
		if vMerge != nil && *vMerge == "continue" {
			// Delegate to the tc above (recursively)
			above := r.tcAbove(tc)
			if above != nil {
				cell := newCell(above, r.table)
				for i := 0; i < gridSpan; i++ {
					cells = append(cells, cell)
				}
				continue
			}
		}
		cell := newCell(tc, r.table)
		for i := 0; i < gridSpan; i++ {
			cells = append(cells, cell)
		}
	}
	return cells
}

// tcAbove finds the top cell of a vertical merge region by walking upward
// iteratively. Previous recursive implementation could overflow the stack
// on tall merge regions.
func (r *Row) tcAbove(tc *oxml.CT_Tc) *oxml.CT_Tc {
	trList := r.table.tbl.TrList()
	idx := r.tr.TrIdx()
	if idx <= 0 || idx > len(trList) {
		return nil
	}

	curTr := r.tr
	curTc := tc

	for idx > 0 {
		// Find the grid offset of curTc in curTr
		offset := 0
		for _, c := range curTr.TcList() {
			if c.RawElement() == curTc.RawElement() {
				break
			}
			gs, err := c.GridSpanVal()
			if err != nil {
				gs = 1
			}
			offset += gs
		}

		prevTr := trList[idx-1]
		above, err := prevTr.TcAtGridOffset(offset)
		if err != nil {
			return nil
		}

		// If the cell above is not a vMerge continue, it's the top of the region.
		vm := above.VMergeVal()
		if vm == nil || *vm != "continue" {
			return above
		}

		// Otherwise keep walking up.
		idx--
		curTr = prevTr
		curTc = above
	}
	// Reached row 0 while still seeing vMerge=continue — malformed document.
	return nil
}

// GridColsBefore returns the count of unpopulated grid-columns before the first cell.
func (r *Row) GridColsBefore() (int, error) {
	return r.tr.GridBeforeVal()
}

// GridColsAfter returns the count of unpopulated grid-columns after the last cell.
func (r *Row) GridColsAfter() (int, error) {
	return r.tr.GridAfterVal()
}

// Height returns the row height in twips, or nil if not set.
func (r *Row) Height() (*int, error) {
	return r.tr.TrHeightVal()
}

// SetHeight sets the row height in twips. Passing nil removes it.
func (r *Row) SetHeight(twips *int) error {
	return r.tr.SetTrHeightVal(twips)
}

// HeightRule returns the height rule, or nil if not set.
func (r *Row) HeightRule() (*enum.WdRowHeightRule, error) {
	return r.tr.TrHeightHRule()
}

// SetHeightRule sets the height rule.
func (r *Row) SetHeightRule(v *enum.WdRowHeightRule) error {
	return r.tr.SetTrHeightHRule(v)
}

// Table returns the Table this row belongs to.
func (r *Row) Table() *Table { return r.table }

// --------------------------------------------------------------------------
// Column
// --------------------------------------------------------------------------

// Column is a proxy for a <w:gridCol> element.
//
// Mirrors Python _Column(Parented).
type Column struct {
	gridCol *oxml.CT_TblGridCol
	table   *Table
}

// Width returns the column width in twips, or nil if not set.
func (c *Column) Width() (*int, error) {
	return c.gridCol.W()
}

// SetWidth sets the column width in twips.
func (c *Column) SetWidth(twips *int) error {
	return c.gridCol.SetW(twips)
}

// Cells returns the cells in this column.
func (c *Column) Cells() ([]*Cell, error) {
	idx, err := c.index()
	if err != nil {
		return nil, err
	}
	return c.table.ColumnCells(idx)
}

// Table returns the Table this column belongs to.
func (c *Column) Table() *Table { return c.table }

func (c *Column) index() (int, error) {
	grid, err := c.table.tbl.TblGrid()
	if err != nil {
		return 0, fmt.Errorf("docx: getting table grid: %w", err)
	}
	for i, gc := range grid.GridColList() {
		if gc.RawElement() == c.gridCol.RawElement() {
			return i, nil
		}
	}
	return 0, fmt.Errorf("docx: column not found in grid")
}

// --------------------------------------------------------------------------
// Rows
// --------------------------------------------------------------------------

// Rows is a sequence of Row objects.
//
// Mirrors Python _Rows(Parented).
type Rows struct {
	tbl   *oxml.CT_Tbl
	table *Table
}

// Len returns the number of rows.
func (rs *Rows) Len() int { return len(rs.tbl.TrList()) }

// Get returns the row at the given index.
func (rs *Rows) Get(idx int) (*Row, error) {
	lst := rs.tbl.TrList()
	if idx < 0 || idx >= len(lst) {
		return nil, fmt.Errorf("docx: row index [%d] out of range", idx)
	}
	return &Row{tr: lst[idx], table: rs.table}, nil
}

// Iter returns all rows in document order.
func (rs *Rows) Iter() []*Row {
	lst := rs.tbl.TrList()
	result := make([]*Row, len(lst))
	for i, tr := range lst {
		result[i] = &Row{tr: tr, table: rs.table}
	}
	return result
}

// Table returns the Table this Rows belongs to.
func (rs *Rows) Table() *Table { return rs.table }

// --------------------------------------------------------------------------
// Columns
// --------------------------------------------------------------------------

// Columns is a sequence of Column objects.
//
// Mirrors Python _Columns(Parented).
type Columns struct {
	grid  *oxml.CT_TblGrid
	table *Table
}

// Len returns the number of columns.
func (cs *Columns) Len() int { return len(cs.grid.GridColList()) }

// Get returns the column at the given index.
func (cs *Columns) Get(idx int) (*Column, error) {
	lst := cs.grid.GridColList()
	if idx < 0 || idx >= len(lst) {
		return nil, fmt.Errorf("docx: column index [%d] out of range", idx)
	}
	return &Column{gridCol: lst[idx], table: cs.table}, nil
}

// Iter returns all columns in document order.
func (cs *Columns) Iter() []*Column {
	lst := cs.grid.GridColList()
	result := make([]*Column, len(lst))
	for i, gc := range lst {
		result[i] = &Column{gridCol: gc, table: cs.table}
	}
	return result
}

// Table returns the Table this Columns belongs to.
func (cs *Columns) Table() *Table { return cs.table }
