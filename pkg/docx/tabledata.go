package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// tabledata.go — TableData type and table element builder
//
// Contains the public TableData structure for describing a table to insert
// in place of a text placeholder, and the private buildTableElement function
// that converts TableData + width into a raw <w:tbl> element.
//
// This file is part of the ReplaceWithTable feature (Phase 2).
// --------------------------------------------------------------------------

// TableData describes a table to insert in place of a text placeholder.
//
// After passing to [Document.ReplaceWithTable], the contents of Rows are
// copied (defensive copy); the caller may safely modify TableData after
// the call returns.
type TableData struct {
	// Rows contains the table data. The first row may serve as a header.
	// Each row is a slice of cell text values.
	// All rows must have the same length. An empty slice is an error.
	Rows [][]string

	// Style is an optional table style reference (e.g. StyleName("Table Grid")).
	// If nil or zero-value, no style is applied.
	Style StyleRef
}

// defensiveCopy returns a deep copy of td. The caller may safely mutate the
// original after the copy is made. Called once at the Document.ReplaceWithTable
// level — sub-calls (body, headers, comments) use the already-copied value.
func (td TableData) defensiveCopy() TableData {
	rows := make([][]string, len(td.Rows))
	for i, row := range td.Rows {
		rows[i] = make([]string, len(row))
		copy(rows[i], row)
	}
	return TableData{Rows: rows, Style: td.Style}
}

// buildTableElement creates a <w:tbl> element from td with the given
// content width. Returns an error if Rows is empty, a row has zero cells,
// or rows have different lengths.
//
// The returned element is detached (not yet inserted into any container).
// The caller is responsible for splicing it into the document tree.
func buildTableElement(td TableData, widthTwips int) (*etree.Element, error) {
	if len(td.Rows) == 0 {
		return nil, fmt.Errorf("docx: TableData has no rows")
	}
	cols := len(td.Rows[0])
	if cols == 0 {
		return nil, fmt.Errorf("docx: TableData row 0 has no cells")
	}
	for i, row := range td.Rows {
		if len(row) != cols {
			return nil, fmt.Errorf("docx: row %d has %d cells, expected %d", i, len(row), cols)
		}
	}

	tbl := oxml.NewTbl(len(td.Rows), cols, widthTwips)

	// Fill cells with text.
	for r, row := range td.Rows {
		for c, text := range row {
			tc := tbl.TrList()[r].TcList()[c]
			// The first <w:p> in each cell already exists (created by NewTbl).
			p := tc.PList()[0]
			if text != "" {
				run := p.AddR()
				run.AddTWithText(text)
			}
		}
	}

	// Add default single-line borders to tblPr. The "Table Grid" style is
	// listed as an lsdException in the default template but has no actual
	// <w:style> definition, so tblStyle alone produces no visible borders.
	// Explicit tblBorders guarantee borders regardless of style availability.
	addDefaultTableBorders(tbl.RawElement())

	// Apply table style if provided.
	if raw := resolveStyleRef([]StyleRef{td.Style}); raw != nil {
		switch v := raw.(type) {
		case string:
			if v != "" {
				if err := tbl.SetTblStyleVal(v); err != nil {
					return nil, fmt.Errorf("docx: setting table style: %w", err)
				}
			}
		case *BaseStyle:
			return nil, fmt.Errorf("docx: *BaseStyle not supported in ReplaceWithTable; use StyleName")
		}
	}

	return tbl.RawElement(), nil
}

// addDefaultTableBorders adds <w:tblBorders> with single-line borders to the
// table's <w:tblPr>. This produces the standard "Table Grid" look: thin black
// borders on all sides and between cells.
//
//	<w:tblBorders>
//	  <w:top    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
//	  <w:left   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
//	  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
//	  <w:right  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
//	  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
//	  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
//	</w:tblBorders>
func addDefaultTableBorders(tblEl *etree.Element) {
	// Find tblPr (always exists — created by NewTbl).
	var tblPr *etree.Element
	for _, child := range tblEl.ChildElements() {
		if child.Space == "w" && child.Tag == "tblPr" {
			tblPr = child
			break
		}
	}
	if tblPr == nil {
		return
	}

	borders := tblPr.CreateElement("tblBorders")
	borders.Space = "w"
	for _, side := range []string{"top", "left", "bottom", "right", "insideH", "insideV"} {
		b := borders.CreateElement(side)
		b.Space = "w"
		b.CreateAttr("w:val", "single")
		b.CreateAttr("w:sz", "4")
		b.CreateAttr("w:space", "0")
		b.CreateAttr("w:color", "auto")
	}
}
