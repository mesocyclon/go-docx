package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// -----------------------------------------------------------------------
// table_test.go — Table / Cell (Batch 1)
// Mirrors Python: tests/test_table.py
// -----------------------------------------------------------------------

func twoByTwoGrid() string {
	return `<w:tblPr/>` +
		`<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>` +
		`<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>` +
		`<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>`
}

// Mirrors Python: it_can_add_a_row
func TestTable_AddRow(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	initialRows := table.Rows().Len()

	_, err := table.AddRow()
	if err != nil {
		t.Fatalf("AddRow: %v", err)
	}
	if table.Rows().Len() != initialRows+1 {
		t.Errorf("Rows.Len() = %d, want %d", table.Rows().Len(), initialRows+1)
	}
}

// Mirrors Python: it_can_add_a_column
func TestTable_AddColumn(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	cols, err := table.Columns()
	if err != nil {
		t.Fatal(err)
	}
	initialCols := cols.Len()

	_, err = table.AddColumn(3000)
	if err != nil {
		t.Fatalf("AddColumn: %v", err)
	}
	cols2, err := table.Columns()
	if err != nil {
		t.Fatal(err)
	}
	if cols2.Len() != initialCols+1 {
		t.Errorf("Columns.Len() = %d, want %d", cols2.Len(), initialCols+1)
	}
}

// Mirrors Python: it_knows_its_alignment_setting (getter, 4 cases)
func TestTable_Alignment_Getter(t *testing.T) {
	wdTblAlignPtr := func(v enum.WdTableAlignment) *enum.WdTableAlignment { return &v }
	tests := []struct {
		name     string
		tblPr    string
		expected *enum.WdTableAlignment
	}{
		{"nil_when_absent", `<w:tblPr/>`, nil},
		{"center", `<w:tblPr><w:jc w:val="center"/></w:tblPr>`, wdTblAlignPtr(enum.WdTableAlignmentCenter)},
		{"right", `<w:tblPr><w:jc w:val="right"/></w:tblPr>`, wdTblAlignPtr(enum.WdTableAlignmentRight)},
		{"left", `<w:tblPr><w:jc w:val="left"/></w:tblPr>`, wdTblAlignPtr(enum.WdTableAlignmentLeft)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tbl := makeTbl(t, tt.tblPr)
			table := newTable(tbl, nil)
			got, err := table.Alignment()
			if err != nil {
				t.Fatal(err)
			}
			if got == nil && tt.expected == nil {
				return
			}
			if got == nil || tt.expected == nil || *got != *tt.expected {
				t.Errorf("Alignment() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_can_change_its_alignment (setter, 3 cases)
func TestTable_Alignment_Setter(t *testing.T) {
	wdTblAlignPtr := func(v enum.WdTableAlignment) *enum.WdTableAlignment { return &v }
	tests := []struct {
		name  string
		tblPr string
		value *enum.WdTableAlignment
	}{
		{"none_to_left", `<w:tblPr/>`, wdTblAlignPtr(enum.WdTableAlignmentLeft)},
		{"left_to_right", `<w:tblPr><w:jc w:val="left"/></w:tblPr>`, wdTblAlignPtr(enum.WdTableAlignmentRight)},
		{"right_to_nil", `<w:tblPr><w:jc w:val="right"/></w:tblPr>`, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tbl := makeTbl(t, tt.tblPr)
			table := newTable(tbl, nil)
			if err := table.SetAlignment(tt.value); err != nil {
				t.Fatal(err)
			}
			got, err := table.Alignment()
			if err != nil {
				t.Fatal(err)
			}
			if got == nil && tt.value == nil {
				return
			}
			if got == nil || tt.value == nil || *got != *tt.value {
				t.Errorf("Alignment() after set = %v, want %v", got, tt.value)
			}
		})
	}
}

// Mirrors Python: it_knows_whether_it_should_autofit
func TestTable_Autofit(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)

	// Set autofit true
	if err := table.SetAutofit(true); err != nil {
		t.Fatal(err)
	}
	got, err := table.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if !got {
		t.Error("Autofit() = false after SetAutofit(true)")
	}
	// Set autofit false
	if err := table.SetAutofit(false); err != nil {
		t.Fatal(err)
	}
	got, err = table.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if got {
		t.Error("Autofit() = true after SetAutofit(false)")
	}
}

// Mirrors Python: Cell.it_knows_its_grid_span
func TestCell_GridSpan(t *testing.T) {
	// gridSpan=2 on first cell
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr>
	`)
	table := newTable(tbl, nil)
	cell, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	if cell.GridSpan() != 2 {
		t.Errorf("GridSpan() = %d, want 2", cell.GridSpan())
	}
}

// Mirrors Python: Cell.it_knows_what_text_it_contains
func TestCell_Text(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	cell, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	if cell.Text() != "A1" {
		t.Errorf("Cell.Text() = %q, want %q", cell.Text(), "A1")
	}
}

// Mirrors Python: Cell.it_can_replace_its_content
func TestCell_SetText(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	cell, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	cell.SetText("new text")
	if cell.Text() != "new text" {
		t.Errorf("Cell.Text() after SetText = %q, want %q", cell.Text(), "new text")
	}
}

// Mirrors Python: Cell.it_can_add_a_paragraph
func TestCell_AddParagraph(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	cell, err := table.CellAt(0, 0)
	if err != nil {
		t.Fatal(err)
	}
	para, err := cell.AddParagraph("added para")
	if err != nil {
		t.Fatalf("AddParagraph: %v", err)
	}
	if para.Text() != "added para" {
		t.Errorf("added paragraph text = %q, want %q", para.Text(), "added para")
	}
}

// Mirrors Python: it_provides_access_to_the_cells_in_a_column
func TestTable_ColumnCells(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	cells, err := table.ColumnCells(0)
	if err != nil {
		t.Fatal(err)
	}
	if len(cells) != 2 {
		t.Fatalf("len(ColumnCells(0)) = %d, want 2", len(cells))
	}
	if cells[0].Text() != "A1" || cells[1].Text() != "A2" {
		t.Errorf("ColumnCells(0) = [%q, %q], want [A1, A2]", cells[0].Text(), cells[1].Text())
	}
}

// Mirrors Python: it_provides_access_to_the_cells_in_a_row
func TestTable_RowCells(t *testing.T) {
	tbl := makeTbl(t, twoByTwoGrid())
	table := newTable(tbl, nil)
	cells, err := table.RowCells(0)
	if err != nil {
		t.Fatal(err)
	}
	if len(cells) != 2 {
		t.Fatalf("len(RowCells(0)) = %d, want 2", len(cells))
	}
	if cells[0].Text() != "A1" || cells[1].Text() != "B1" {
		t.Errorf("RowCells(0) = [%q, %q], want [A1, B1]", cells[0].Text(), cells[1].Text())
	}
}

// Mirrors Python: it_knows_its_direction
func TestTable_TableDirection(t *testing.T) {
	tbl := makeTbl(t, `<w:tblPr/>`)
	table := newTable(tbl, nil)

	// Initially nil
	got, err := table.TableDirection()
	if err != nil {
		t.Fatal(err)
	}
	if got != nil {
		t.Errorf("TableDirection() = %v, want nil", got)
	}

	// Set to true (RTL)
	if err := table.SetTableDirection(boolPtr(true)); err != nil {
		t.Fatal(err)
	}
	got, err = table.TableDirection()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || !*got {
		t.Error("TableDirection() should be true after set")
	}
}

// -----------------------------------------------------------------------
// Row.tcAbove — iterative vMerge resolution
// -----------------------------------------------------------------------

// Basic 2-row vertical merge: row 1 continue → resolves to row 0.
func TestTcAbove_Basic(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc>
			<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
			<w:p><w:r><w:t>top</w:t></w:r></w:p>
		</w:tc></w:tr>
		<w:tr><w:tc>
			<w:tcPr><w:vMerge/></w:tcPr>
			<w:p/>
		</w:tc></w:tr>
	`)
	table := newTable(tbl, nil)
	row, err := table.Rows().Get(1)
	if err != nil {
		t.Fatal(err)
	}
	cells := row.Cells()
	if len(cells) != 1 {
		t.Fatalf("len(Cells) = %d, want 1", len(cells))
	}
	if cells[0].Text() != "top" {
		t.Errorf("Cells()[0].Text() = %q, want %q", cells[0].Text(), "top")
	}
}

// Deep chain: 6 rows, row 0 = restart, rows 1–5 = continue.
// Exercises the iterative loop that replaced the old recursion.
func TestTcAbove_DeepChain(t *testing.T) {
	const depth = 6
	xml := `<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>`
	// Row 0: restart
	xml += `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>` +
		`<w:p><w:r><w:t>origin</w:t></w:r></w:p></w:tc></w:tr>`
	// Rows 1..depth-1: continue
	for i := 1; i < depth; i++ {
		xml += `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>`
	}
	tbl := makeTbl(t, xml)
	table := newTable(tbl, nil)

	// Every row should resolve to "origin"
	for i := 0; i < depth; i++ {
		cell, err := table.CellAt(i, 0)
		if err != nil {
			t.Fatalf("CellAt(%d,0): %v", i, err)
		}
		if cell.Text() != "origin" {
			t.Errorf("CellAt(%d,0) = %q, want %q", i, cell.Text(), "origin")
		}
	}
}

// Row 0 has vMerge=continue (malformed document).
// Should fall back to its own cell — no panic, no infinite loop.
func TestTcAbove_ContinueAtRow0(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc>
			<w:tcPr><w:vMerge/></w:tcPr>
			<w:p><w:r><w:t>orphan</w:t></w:r></w:p>
		</w:tc></w:tr>
	`)
	table := newTable(tbl, nil)
	row, err := table.Rows().Get(0)
	if err != nil {
		t.Fatal(err)
	}
	cells := row.Cells()
	if len(cells) != 1 {
		t.Fatalf("len(Cells) = %d, want 1", len(cells))
	}
	// tcAbove returns nil for row 0 → Cells() falls back to the cell itself
	if cells[0].Text() != "orphan" {
		t.Errorf("Cells()[0].Text() = %q, want %q", cells[0].Text(), "orphan")
	}
}

// Multi-column table: only column 1 is vertically merged.
// Verifies grid-offset calculation inside the iterative loop.
func TestTcAbove_MultiColumn_PartialMerge(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid>
			<w:gridCol w:w="3000"/>
			<w:gridCol w:w="3000"/>
			<w:gridCol w:w="3000"/>
		</w:tblGrid>
		<w:tr>
			<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
			<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
				<w:p><w:r><w:t>B-top</w:t></w:r></w:p></w:tc>
			<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
		</w:tr>
		<w:tr>
			<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
			<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
			<w:tc><w:p><w:r><w:t>F</w:t></w:r></w:p></w:tc>
		</w:tr>
		<w:tr>
			<w:tc><w:p><w:r><w:t>G</w:t></w:r></w:p></w:tc>
			<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
			<w:tc><w:p><w:r><w:t>I</w:t></w:r></w:p></w:tc>
		</w:tr>
	`)
	table := newTable(tbl, nil)

	// Row 2, col 1 should resolve through row 1 continue → row 0 restart
	cell, err := table.CellAt(2, 1)
	if err != nil {
		t.Fatal(err)
	}
	if cell.Text() != "B-top" {
		t.Errorf("CellAt(2,1) = %q, want %q", cell.Text(), "B-top")
	}
	// Non-merged columns are unaffected
	for _, tc := range []struct {
		r, c int
		want string
	}{
		{0, 0, "A"}, {1, 0, "D"}, {2, 0, "G"},
		{0, 2, "C"}, {1, 2, "F"}, {2, 2, "I"},
	} {
		got, err := table.CellAt(tc.r, tc.c)
		if err != nil {
			t.Fatalf("CellAt(%d,%d): %v", tc.r, tc.c, err)
		}
		if got.Text() != tc.want {
			t.Errorf("CellAt(%d,%d) = %q, want %q", tc.r, tc.c, got.Text(), tc.want)
		}
	}
}

// Two separate vMerge regions in the same column (restart in the middle).
func TestTcAbove_TwoRegions(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
		<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
			<w:p><w:r><w:t>first</w:t></w:r></w:p></w:tc></w:tr>
		<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>
		<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
			<w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc></w:tr>
		<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>
	`)
	table := newTable(tbl, nil)

	tests := []struct {
		row  int
		want string
	}{
		{0, "first"},
		{1, "first"},  // continue → row 0
		{2, "second"}, // restart — new region
		{3, "second"}, // continue → row 2
	}
	for _, tc := range tests {
		cell, err := table.CellAt(tc.row, 0)
		if err != nil {
			t.Fatalf("CellAt(%d,0): %v", tc.row, err)
		}
		if cell.Text() != tc.want {
			t.Errorf("CellAt(%d,0) = %q, want %q", tc.row, cell.Text(), tc.want)
		}
	}
}

// gridSpan + vMerge: cell spans 2 columns horizontally AND is part of a
// vertical merge.  Row.Cells() should expand to 2 identical cells.
func TestTcAbove_GridSpanPlusVMerge(t *testing.T) {
	tbl := makeTbl(t, `
		<w:tblGrid>
			<w:gridCol w:w="3000"/>
			<w:gridCol w:w="3000"/>
		</w:tblGrid>
		<w:tr><w:tc>
			<w:tcPr>
				<w:gridSpan w:val="2"/>
				<w:vMerge w:val="restart"/>
			</w:tcPr>
			<w:p><w:r><w:t>wide-top</w:t></w:r></w:p>
		</w:tc></w:tr>
		<w:tr><w:tc>
			<w:tcPr>
				<w:gridSpan w:val="2"/>
				<w:vMerge/>
			</w:tcPr>
			<w:p/>
		</w:tc></w:tr>
	`)
	table := newTable(tbl, nil)
	row, err := table.Rows().Get(1)
	if err != nil {
		t.Fatal(err)
	}
	cells := row.Cells()
	if len(cells) != 2 {
		t.Fatalf("len(Cells) = %d, want 2 (gridSpan expansion)", len(cells))
	}
	for i, c := range cells {
		if c.Text() != "wide-top" {
			t.Errorf("Cells()[%d].Text() = %q, want %q", i, c.Text(), "wide-top")
		}
	}
}
