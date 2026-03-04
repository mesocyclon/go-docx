package oxml

import (
	"fmt"
	"strconv"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ===========================================================================
// CT_Tbl — custom methods
// ===========================================================================

// NewTbl creates a new <w:tbl> element with the given number of rows and columns.
// Width (in twips) is distributed evenly among columns.
func NewTbl(rows, cols int, widthTwips int) *CT_Tbl {
	tblE := OxmlElement("w:tbl")
	tbl := &CT_Tbl{Element{e: tblE}}

	// tblPr
	tblPrE := tblE.CreateElement("tblPr")
	tblPrE.Space = "w"
	tblW := tblPrE.CreateElement("tblW")
	tblW.Space = "w"
	tblW.CreateAttr("w:type", "auto")
	tblW.CreateAttr("w:w", "0")
	tblLook := tblPrE.CreateElement("tblLook")
	tblLook.Space = "w"
	tblLook.CreateAttr("w:firstColumn", "1")
	tblLook.CreateAttr("w:firstRow", "1")
	tblLook.CreateAttr("w:lastColumn", "0")
	tblLook.CreateAttr("w:lastRow", "0")
	tblLook.CreateAttr("w:noHBand", "0")
	tblLook.CreateAttr("w:noVBand", "1")
	tblLook.CreateAttr("w:val", "04A0")

	// tblGrid
	colWidth := 0
	if cols > 0 {
		colWidth = widthTwips / cols
	}
	tblGridE := tblE.CreateElement("tblGrid")
	tblGridE.Space = "w"
	for i := 0; i < cols; i++ {
		gc := tblGridE.CreateElement("gridCol")
		gc.Space = "w"
		gc.CreateAttr("w:w", strconv.Itoa(colWidth))
	}

	// rows
	for r := 0; r < rows; r++ {
		trE := tblE.CreateElement("tr")
		trE.Space = "w"
		for c := 0; c < cols; c++ {
			tcE := trE.CreateElement("tc")
			tcE.Space = "w"
			tcPrE := tcE.CreateElement("tcPr")
			tcPrE.Space = "w"
			tcW := tcPrE.CreateElement("tcW")
			tcW.Space = "w"
			tcW.CreateAttr("w:type", "dxa")
			tcW.CreateAttr("w:w", strconv.Itoa(colWidth))
			pE := tcE.CreateElement("p")
			pE.Space = "w"
		}
	}

	return tbl
}

// TblStyleVal returns the value of tblPr/tblStyle/@w:val or "" if not present.
func (t *CT_Tbl) TblStyleVal() (string, error) {
	tblPr, err := t.TblPr()
	if err != nil {
		return "", fmt.Errorf("TblStyleVal: %w", err)
	}
	ts := tblPr.TblStyle()
	if ts == nil {
		return "", nil
	}
	return ts.Val()
}

// SetTblStyleVal sets tblPr/tblStyle/@w:val. Passing "" removes tblStyle.
func (t *CT_Tbl) SetTblStyleVal(styleID string) error {
	tblPr, err := t.TblPr()
	if err != nil {
		return fmt.Errorf("SetTblStyleVal: %w", err)
	}
	tblPr.RemoveTblStyle()
	if styleID == "" {
		return nil
	}
	if err := tblPr.GetOrAddTblStyle().SetVal(styleID); err != nil {
		return err
	}
	return nil
}

// AlignmentVal returns the table alignment from tblPr/jc, or nil if not set.
func (t *CT_Tbl) AlignmentVal() (*enum.WdTableAlignment, error) {
	tblPr, err := t.TblPr()
	if err != nil {
		return nil, fmt.Errorf("AlignmentVal: %w", err)
	}
	jc := tblPr.Jc()
	if jc == nil {
		return nil, nil
	}
	val, ok := jc.GetAttr("w:val")
	if !ok {
		return nil, nil
	}
	v, err := enum.WdTableAlignmentFromXml(val)
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetAlignmentVal sets the table alignment. Passing nil removes jc.
func (t *CT_Tbl) SetAlignmentVal(v *enum.WdTableAlignment) error {
	tblPr, err := t.TblPr()
	if err != nil {
		return fmt.Errorf("SetAlignmentVal: %w", err)
	}
	tblPr.RemoveJc()
	if v == nil {
		return nil
	}
	xmlVal, err := v.ToXml()
	if err != nil {
		return fmt.Errorf("oxml: invalid table alignment: %w", err)
	}
	jc := tblPr.GetOrAddJc()
	jc.SetAttr("w:val", xmlVal)
	return nil
}

// BidiVisualVal returns the value of tblPr/bidiVisual, or nil if not present.
func (t *CT_Tbl) BidiVisualVal() (*bool, error) {
	tblPr, err := t.TblPr()
	if err != nil {
		return nil, fmt.Errorf("BidiVisualVal: %w", err)
	}
	bidi := tblPr.BidiVisual()
	if bidi == nil {
		return nil, nil
	}
	v := bidi.Val()
	return &v, nil
}

// SetBidiVisualVal sets tblPr/bidiVisual. Passing nil removes it.
func (t *CT_Tbl) SetBidiVisualVal(v *bool) error {
	tblPr, err := t.TblPr()
	if err != nil {
		return fmt.Errorf("SetBidiVisualVal: %w", err)
	}
	if v == nil {
		tblPr.RemoveBidiVisual()
		return nil
	}
	if err := tblPr.GetOrAddBidiVisual().SetVal(*v); err != nil {
		return err
	}
	return nil
}

// Autofit returns false when there is a tblLayout with type="fixed", true otherwise.
func (t *CT_Tbl) Autofit() (bool, error) {
	tblPr, err := t.TblPr()
	if err != nil {
		return false, fmt.Errorf("Autofit: %w", err)
	}
	layout := tblPr.TblLayout()
	if layout == nil {
		return true, nil
	}
	return layout.Type() != "fixed", nil
}

// SetAutofit sets the table layout to "autofit" or "fixed".
func (t *CT_Tbl) SetAutofit(v bool) error {
	tblPr, err := t.TblPr()
	if err != nil {
		return fmt.Errorf("SetAutofit: %w", err)
	}
	layout := tblPr.GetOrAddTblLayout()
	if v {
		if err := layout.SetType("autofit"); err != nil {
			return err
		}
	} else {
		if err := layout.SetType("fixed"); err != nil {
			return err
		}
	}
	return nil
}

// ColCount returns the number of grid columns defined in tblGrid.
func (t *CT_Tbl) ColCount() (int, error) {
	grid, err := t.TblGrid()
	if err != nil {
		return 0, fmt.Errorf("ColCount: %w", err)
	}
	return len(grid.GridColList()), nil
}

// IterTcs generates each w:tc element in this table, left to right, top to bottom.
func (t *CT_Tbl) IterTcs() []*CT_Tc {
	var result []*CT_Tc
	for _, tr := range t.TrList() {
		result = append(result, tr.TcList()...)
	}
	return result
}

// ColWidths returns the widths (in twips) of each grid column.
func (t *CT_Tbl) ColWidths() ([]int, error) {
	grid, err := t.TblGrid()
	if err != nil {
		return nil, fmt.Errorf("ColWidths: %w", err)
	}
	cols := grid.GridColList()
	result := make([]int, len(cols))
	for i, col := range cols {
		w, err := col.W()
		if err != nil {
			return nil, fmt.Errorf("ColWidths: grid col %d: %w", i, err)
		}
		if w != nil {
			result[i] = *w
		}
	}
	return result, nil
}

// ===========================================================================
// CT_TblPr — custom methods
// ===========================================================================

// AlignmentVal returns the table alignment, or nil.
func (pr *CT_TblPr) AlignmentVal() (*enum.WdTableAlignment, error) {
	jc := pr.Jc()
	if jc == nil {
		return nil, nil
	}
	val, ok := jc.GetAttr("w:val")
	if !ok {
		return nil, nil
	}
	v, err := enum.WdTableAlignmentFromXml(val)
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetAlignmentVal sets the table alignment. Passing nil removes jc.
func (pr *CT_TblPr) SetAlignmentVal(v *enum.WdTableAlignment) error {
	pr.RemoveJc()
	if v == nil {
		return nil
	}
	xmlVal, err := v.ToXml()
	if err != nil {
		return fmt.Errorf("oxml: invalid table alignment: %w", err)
	}
	jc := pr.GetOrAddJc()
	jc.SetAttr("w:val", xmlVal)
	return nil
}

// AutofitVal returns false when tblLayout type="fixed", true otherwise.
func (pr *CT_TblPr) AutofitVal() bool {
	layout := pr.TblLayout()
	if layout == nil {
		return true
	}
	return layout.Type() != "fixed"
}

// SetAutofitVal sets the autofit property.
func (pr *CT_TblPr) SetAutofitVal(v bool) error {
	layout := pr.GetOrAddTblLayout()
	if v {
		if err := layout.SetType("autofit"); err != nil {
			return err
		}
	} else {
		if err := layout.SetType("fixed"); err != nil {
			return err
		}
	}
	return nil
}

// StyleVal returns the value of tblStyle/@w:val or "" if absent.
func (pr *CT_TblPr) StyleVal() (string, error) {
	ts := pr.TblStyle()
	if ts == nil {
		return "", nil
	}
	return ts.Val()
}

// SetStyleVal sets the table style. Passing "" removes tblStyle.
func (pr *CT_TblPr) SetStyleVal(v string) error {
	pr.RemoveTblStyle()
	if v == "" {
		return nil
	}
	if err := pr.GetOrAddTblStyle().SetVal(v); err != nil {
		return err
	}
	return nil
}

// ===========================================================================
// CT_Row — custom methods
// ===========================================================================

// TrIdx returns the index of this w:tr within its parent w:tbl.
// Returns -1 if parent is not found.
func (r *CT_Row) TrIdx() int {
	parent := r.e.Parent()
	if parent == nil {
		return -1
	}
	idx := 0
	for _, child := range parent.ChildElements() {
		if child.Space == "w" && child.Tag == "tr" {
			if child == r.e {
				return idx
			}
			idx++
		}
	}
	return -1
}

// GridBeforeVal returns the number of unpopulated grid cells at the start of this row.
func (r *CT_Row) GridBeforeVal() (int, error) {
	trPr := r.TrPr()
	if trPr == nil {
		return 0, nil
	}
	return trPr.GridBeforeVal()
}

// GridAfterVal returns the number of unpopulated grid cells at the end of this row.
func (r *CT_Row) GridAfterVal() (int, error) {
	trPr := r.TrPr()
	if trPr == nil {
		return 0, nil
	}
	return trPr.GridAfterVal()
}

// TcAtGridOffset returns the w:tc at the given grid column offset.
// Returns error if no tc at that exact offset.
func (r *CT_Row) TcAtGridOffset(gridOffset int) (*CT_Tc, error) {
	gb, err := r.GridBeforeVal()
	if err != nil {
		return nil, err
	}
	remaining := gridOffset - gb
	for _, tc := range r.TcList() {
		if remaining < 0 {
			break
		}
		if remaining == 0 {
			return tc, nil
		}
		span, err := tc.GridSpanVal()
		if err != nil {
			return nil, err
		}
		remaining -= span
	}
	return nil, fmt.Errorf("no tc element at grid_offset=%d", gridOffset)
}

// TrHeightVal returns the value of trPr/trHeight/@w:val (twips), or nil.
func (r *CT_Row) TrHeightVal() (*int, error) {
	trPr := r.TrPr()
	if trPr == nil {
		return nil, nil
	}
	return trPr.TrHeightValTwips()
}

// SetTrHeightVal sets the row height. Passing nil removes it.
func (r *CT_Row) SetTrHeightVal(twips *int) error {
	if twips == nil {
		trPr := r.TrPr()
		if trPr != nil {
			trPr.RemoveTrHeight()
		}
		return nil
	}
	trPr := r.GetOrAddTrPr()
	h := trPr.GetOrAddTrHeight()
	if err := h.SetVal(twips); err != nil {
		return err
	}
	return nil
}

// TrHeightHRule returns the row height rule, or nil.
func (r *CT_Row) TrHeightHRule() (*enum.WdRowHeightRule, error) {
	trPr := r.TrPr()
	if trPr == nil {
		return nil, nil
	}
	return trPr.TrHeightHRuleVal()
}

// SetTrHeightHRule sets the row height rule. Passing nil removes it.
func (r *CT_Row) SetTrHeightHRule(rule *enum.WdRowHeightRule) error {
	if rule == nil {
		trPr := r.TrPr()
		if trPr != nil {
			h := trPr.TrHeight()
			if h != nil {
				return h.SetHRule(enum.WdRowHeightRule(0))
			}
		}
		return nil
	}
	trPr := r.GetOrAddTrPr()
	h := trPr.GetOrAddTrHeight()
	return h.SetHRule(*rule)
}

// ===========================================================================
// CT_TrPr — custom methods
// ===========================================================================

// GridBeforeVal returns the value of gridBefore/@w:val or 0 if not present.
func (pr *CT_TrPr) GridBeforeVal() (int, error) {
	gb := pr.GridBefore()
	if gb == nil {
		return 0, nil
	}
	return gb.Val()
}

// GridAfterVal returns the value of gridAfter/@w:val or 0 if not present.
func (pr *CT_TrPr) GridAfterVal() (int, error) {
	ga := pr.GridAfter()
	if ga == nil {
		return 0, nil
	}
	return ga.Val()
}

// TrHeightValTwips returns the value of trHeight/@w:val in twips, or nil.
func (pr *CT_TrPr) TrHeightValTwips() (*int, error) {
	h := pr.TrHeight()
	if h == nil {
		return nil, nil
	}
	return h.Val()
}

// SetTrHeightValTwips sets the trHeight value. Passing nil removes trHeight.
func (pr *CT_TrPr) SetTrHeightValTwips(twips *int) error {
	if twips == nil {
		pr.RemoveTrHeight()
		return nil
	}
	h := pr.GetOrAddTrHeight()
	if err := h.SetVal(twips); err != nil {
		return err
	}
	return nil
}

// TrHeightHRuleVal returns the height rule, or nil.
func (pr *CT_TrPr) TrHeightHRuleVal() (*enum.WdRowHeightRule, error) {
	h := pr.TrHeight()
	if h == nil {
		return nil, nil
	}
	v, err := h.HRule()
	if err != nil {
		return nil, err
	}
	if v == enum.WdRowHeightRule(0) {
		return nil, nil
	}
	return &v, nil
}

// SetTrHeightHRuleVal sets the height rule. Passing nil removes it.
func (pr *CT_TrPr) SetTrHeightHRuleVal(rule *enum.WdRowHeightRule) error {
	if rule == nil {
		h := pr.TrHeight()
		if h != nil {
			return h.SetHRule(enum.WdRowHeightRule(0))
		}
		return nil
	}
	h := pr.GetOrAddTrHeight()
	return h.SetHRule(*rule)
}

// ===========================================================================
// CT_Tc — custom methods
// ===========================================================================

// NewTc creates a new <w:tc> element with a single empty <w:p>.
func NewTc() *CT_Tc {
	tcE := OxmlElement("w:tc")
	tc := &CT_Tc{Element{e: tcE}}
	tc.AddP()
	return tc
}

// GridSpanVal returns the number of grid columns this cell spans (default 1).
func (tc *CT_Tc) GridSpanVal() (int, error) {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return 1, nil
	}
	return tcPr.GridSpanVal()
}

// SetGridSpanVal sets the grid span. Values ≤ 1 remove the gridSpan element.
func (tc *CT_Tc) SetGridSpanVal(v int) error {
	tcPr := tc.GetOrAddTcPr()
	if err := tcPr.SetGridSpanVal(v); err != nil {
		return err
	}
	return nil
}

// VMergeVal returns the value of tcPr/vMerge/@w:val, or nil if vMerge is not present.
// When vMerge is present without @val, returns "continue".
func (tc *CT_Tc) VMergeVal() *string {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return nil
	}
	return tcPr.VMergeValStr()
}

// SetVMergeVal sets the vMerge value. Passing nil removes vMerge.
func (tc *CT_Tc) SetVMergeVal(v *string) error {
	tcPr := tc.GetOrAddTcPr()
	if err := tcPr.SetVMergeValStr(v); err != nil {
		return err
	}
	return nil
}

// WidthTwips returns the cell width in twips from tcPr/tcW, or nil if not present
// or not dxa type.
func (tc *CT_Tc) WidthTwips() (*int, error) {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return nil, nil
	}
	return tcPr.WidthTwips()
}

// SetWidthTwips sets the cell width in twips.
func (tc *CT_Tc) SetWidthTwips(twips int) error {
	tcPr := tc.GetOrAddTcPr()
	if err := tcPr.SetWidthTwips(twips); err != nil {
		return err
	}
	return nil
}

// VAlignVal returns the vertical alignment of this cell, or nil.
func (tc *CT_Tc) VAlignVal() (*enum.WdCellVerticalAlignment, error) {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return nil, nil
	}
	return tcPr.VAlignValEnum()
}

// SetVAlignVal sets the vertical alignment. Passing nil removes vAlign.
func (tc *CT_Tc) SetVAlignVal(v *enum.WdCellVerticalAlignment) error {
	tcPr := tc.GetOrAddTcPr()
	return tcPr.SetVAlignValEnum(v)
}

// InnerContentElements returns all w:p and w:tbl direct children in document order.
func (tc *CT_Tc) InnerContentElements() []BlockItem {
	var result []BlockItem
	for _, child := range tc.e.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			result = append(result, &CT_P{Element{e: child}})
		} else if child.Space == "w" && child.Tag == "tbl" {
			result = append(result, &CT_Tbl{Element{e: child}})
		}
	}
	return result
}

// IterBlockItems generates all block-level content elements: w:p, w:tbl, w:sdt.
func (tc *CT_Tc) IterBlockItems() []*etree.Element {
	var result []*etree.Element
	for _, child := range tc.e.ChildElements() {
		if child.Space == "w" {
			switch child.Tag {
			case "p", "tbl", "sdt":
				result = append(result, child)
			}
		}
	}
	return result
}

// ClearContent removes all children except w:tcPr.
// NOTE: This leaves the cell in an invalid state (missing required w:p).
// Caller must add a w:p afterwards.
func (tc *CT_Tc) ClearContent() {
	var toRemove []*etree.Element
	for _, child := range tc.e.ChildElements() {
		if !(child.Space == "w" && child.Tag == "tcPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		tc.e.RemoveChild(child)
	}
}

// GridOffset returns the starting offset of this cell in the grid columns.
func (tc *CT_Tc) GridOffset() (int, error) {
	tr := tc.parentTr()
	if tr == nil {
		return 0, nil
	}
	offset, err := tr.GridBeforeVal()
	if err != nil {
		return 0, err
	}
	for _, child := range tr.e.ChildElements() {
		if child.Space == "w" && child.Tag == "tc" {
			if child == tc.e {
				return offset, nil
			}
			sibling := &CT_Tc{Element{e: child}}
			span, err := sibling.GridSpanVal()
			if err != nil {
				return 0, err
			}
			offset += span
		}
	}
	return offset, nil
}

// Left is an alias for GridOffset.
func (tc *CT_Tc) Left() (int, error) {
	return tc.GridOffset()
}

// Right returns the grid column just past the right edge of this cell.
func (tc *CT_Tc) Right() (int, error) {
	off, err := tc.GridOffset()
	if err != nil {
		return 0, err
	}
	span, err := tc.GridSpanVal()
	if err != nil {
		return 0, err
	}
	return off + span, nil
}

// Top returns the top-most row index in the vertical span of this cell.
func (tc *CT_Tc) Top() (int, error) {
	vm := tc.VMergeVal()
	if vm == nil || *vm == "restart" {
		return tc.trIdx(), nil
	}
	above, err := tc.tcAbove()
	if err != nil {
		return 0, fmt.Errorf("Top: %w", err)
	}
	if above != nil {
		return above.Top()
	}
	return tc.trIdx(), nil
}

// Bottom returns the row index just past the bottom of the vertical span.
func (tc *CT_Tc) Bottom() (int, error) {
	vm := tc.VMergeVal()
	if vm != nil {
		below, err := tc.tcBelow()
		if err != nil {
			return 0, fmt.Errorf("Bottom: %w", err)
		}
		if below != nil {
			bvm := below.VMergeVal()
			if bvm != nil && *bvm == "continue" {
				return below.Bottom()
			}
		}
	}
	return tc.trIdx() + 1, nil
}

// IsEmpty returns true if this cell contains only a single empty w:p.
func (tc *CT_Tc) IsEmpty() bool {
	blocks := tc.IterBlockItems()
	if len(blocks) != 1 {
		return false
	}
	b := blocks[0]
	if !(b.Space == "w" && b.Tag == "p") {
		return false
	}
	p := &CT_P{Element{e: b}}
	return len(p.RList()) == 0
}

// NextTc returns the w:tc element immediately following this one in the row, or nil.
func (tc *CT_Tc) NextTc() *CT_Tc {
	found := false
	tr := tc.parentTr()
	if tr == nil {
		return nil
	}
	for _, child := range tr.e.ChildElements() {
		if found && child.Space == "w" && child.Tag == "tc" {
			return &CT_Tc{Element{e: child}}
		}
		if child == tc.e {
			found = true
		}
	}
	return nil
}

// AddWidthOf adds the width of other to this cell. Does nothing if either has no width.
func (tc *CT_Tc) AddWidthOf(other *CT_Tc) error {
	w1, err := tc.WidthTwips()
	if err != nil {
		return err
	}
	w2, err := other.WidthTwips()
	if err != nil {
		return err
	}
	if w1 != nil && w2 != nil {
		sum := *w1 + *w2
		if err := tc.SetWidthTwips(sum); err != nil {
			return err
		}
	}
	return nil
}

// MoveContentTo appends the block-level content of this cell to other.
// Leaves this cell with a single empty w:p.
func (tc *CT_Tc) MoveContentTo(other *CT_Tc) {
	if other.e == tc.e {
		return
	}
	if tc.IsEmpty() {
		return
	}
	// Remove trailing empty p from other
	other.removeTrailingEmptyP()
	// Move all block items
	for _, block := range tc.IterBlockItems() {
		tc.e.RemoveChild(block)
		other.e.AddChild(block)
	}
	// Restore minimum required p
	tc.AddP()
}

// RemoveElement removes this tc from its parent row.
func (tc *CT_Tc) RemoveElement() {
	parent := tc.e.Parent()
	if parent != nil {
		parent.RemoveChild(tc.e)
	}
}

// Merge merges the rectangular region defined by this tc and other as diagonal corners.
// Returns the top-left tc element of the new span.
func (tc *CT_Tc) Merge(other *CT_Tc) (*CT_Tc, error) {
	top, left, height, width, err := tc.spanDimensions(other)
	if err != nil {
		return nil, err
	}
	tbl := tc.parentTbl()
	if tbl == nil {
		return nil, fmt.Errorf("tc has no parent tbl")
	}
	trs := tbl.TrList()
	if top >= len(trs) {
		return nil, fmt.Errorf("top row %d out of range", top)
	}
	topTc, err := trs[top].TcAtGridOffset(left)
	if err != nil {
		return nil, err
	}
	err = topTc.growTo(width, height, topTc)
	if err != nil {
		return nil, err
	}
	return topTc, nil
}

// --- private helpers ---

func (tc *CT_Tc) parentTr() *CT_Row {
	p := tc.e.Parent()
	if p == nil || !(p.Space == "w" && p.Tag == "tr") {
		return nil
	}
	return &CT_Row{Element{e: p}}
}

func (tc *CT_Tc) parentTbl() *CT_Tbl {
	tr := tc.parentTr()
	if tr == nil {
		return nil
	}
	p := tr.e.Parent()
	if p == nil || !(p.Space == "w" && p.Tag == "tbl") {
		return nil
	}
	return &CT_Tbl{Element{e: p}}
}

func (tc *CT_Tc) trIdx() int {
	tr := tc.parentTr()
	if tr == nil {
		return 0
	}
	return tr.TrIdx()
}

func (tc *CT_Tc) tcAbove() (*CT_Tc, error) {
	tbl := tc.parentTbl()
	if tbl == nil {
		return nil, nil
	}
	trs := tbl.TrList()
	idx := tc.trIdx()
	if idx <= 0 {
		return nil, nil
	}
	off, err := tc.GridOffset()
	if err != nil {
		return nil, fmt.Errorf("tcAbove: %w", err)
	}
	above, err := trs[idx-1].TcAtGridOffset(off)
	if err != nil {
		return nil, fmt.Errorf("tcAbove: %w", err)
	}
	return above, nil
}

func (tc *CT_Tc) tcBelow() (*CT_Tc, error) {
	tbl := tc.parentTbl()
	if tbl == nil {
		return nil, nil
	}
	trs := tbl.TrList()
	idx := tc.trIdx()
	if idx >= len(trs)-1 {
		return nil, nil
	}
	off, err := tc.GridOffset()
	if err != nil {
		return nil, fmt.Errorf("tcBelow: %w", err)
	}
	below, err := trs[idx+1].TcAtGridOffset(off)
	if err != nil {
		return nil, fmt.Errorf("tcBelow: %w", err)
	}
	return below, nil
}

func (tc *CT_Tc) removeTrailingEmptyP() {
	blocks := tc.IterBlockItems()
	if len(blocks) == 0 {
		return
	}
	last := blocks[len(blocks)-1]
	if !(last.Space == "w" && last.Tag == "p") {
		return
	}
	p := &CT_P{Element{e: last}}
	if len(p.RList()) > 0 {
		return
	}
	tc.e.RemoveChild(last)
}

func (tc *CT_Tc) spanDimensions(other *CT_Tc) (top, left, height, width int, err error) {
	aTop, err := tc.Top()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	aLeft, err := tc.Left()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	aBottom, err := tc.Bottom()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	aRight, err := tc.Right()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	bTop, err := other.Top()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	bLeft, err := other.Left()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	bBottom, err := other.Bottom()
	if err != nil {
		return 0, 0, 0, 0, err
	}
	bRight, err := other.Right()
	if err != nil {
		return 0, 0, 0, 0, err
	}

	// Check inverted-L
	if aTop == bTop && aBottom != bBottom {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (inverted-L)")
	}
	if aLeft == bLeft && aRight != bRight {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (inverted-L)")
	}
	// Check tee-shaped (using already-extracted values to avoid re-calling failable methods)
	topMostTop, otherTcTop := aTop, bTop
	topMostBottom, otherTcBottom := aBottom, bBottom
	if otherTcTop < topMostTop {
		topMostTop, otherTcTop = otherTcTop, topMostTop
		topMostBottom, otherTcBottom = otherTcBottom, topMostBottom
	}
	if topMostTop < otherTcTop && topMostBottom > otherTcBottom {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (tee)")
	}
	leftMostLeft, otherTc2Left := aLeft, bLeft
	leftMostRight, otherTc2Right := aRight, bRight
	if otherTc2Left < leftMostLeft {
		leftMostLeft, otherTc2Left = otherTc2Left, leftMostLeft
		leftMostRight, otherTc2Right = otherTc2Right, leftMostRight
	}
	if leftMostLeft < otherTc2Left && leftMostRight > otherTc2Right {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (tee)")
	}

	if aTop < bTop {
		top = aTop
	} else {
		top = bTop
	}
	if aLeft < bLeft {
		left = aLeft
	} else {
		left = bLeft
	}
	bottom := aBottom
	if bBottom > bottom {
		bottom = bBottom
	}
	right := aRight
	if bRight > right {
		right = bRight
	}
	return top, left, bottom - top, right - left, nil
}

func (tc *CT_Tc) growTo(width, height int, topTc *CT_Tc) error {
	vMerge := ""
	if topTc.e != tc.e {
		vMerge = "continue"
	} else if height > 1 {
		vMerge = "restart"
	}

	tc.MoveContentTo(topTc)
	// Span to width
	for {
		gsv, err := tc.GridSpanVal()
		if err != nil {
			return err
		}
		if gsv >= width {
			break
		}
		next := tc.NextTc()
		if next == nil {
			return fmt.Errorf("not enough grid columns")
		}
		nextSpan, err := next.GridSpanVal()
		if err != nil {
			return err
		}
		if gsv+nextSpan > width {
			return fmt.Errorf("span is not rectangular")
		}
		next.MoveContentTo(topTc)
		if err := tc.AddWidthOf(next); err != nil {
			return err
		}
		newSpan, err := tc.GridSpanVal()
		if err != nil {
			return err
		}
		if err := tc.SetGridSpanVal(newSpan + nextSpan); err != nil {
			return err
		}
		next.RemoveElement()
	}

	if vMerge == "" {
		// Remove vMerge entirely
		if err := tc.SetVMergeVal(nil); err != nil {
			return err
		}
	} else {
		if err := tc.SetVMergeVal(&vMerge); err != nil {
			return err
		}
	}

	if height > 1 {
		below, err := tc.tcBelow()
		if err != nil {
			return fmt.Errorf("growTo: %w", err)
		}
		if below == nil {
			return fmt.Errorf("not enough rows for vertical span")
		}
		return below.growTo(width, height-1, topTc)
	}
	return nil
}

// ===========================================================================
// CT_TcPr — custom methods
// ===========================================================================

// GridSpanVal returns the grid span value (default 1).
func (pr *CT_TcPr) GridSpanVal() (int, error) {
	gs := pr.GridSpan()
	if gs == nil {
		return 1, nil
	}
	v, err := gs.Val()
	if err != nil {
		return 0, err
	}
	return v, nil
}

// SetGridSpanVal sets the grid span. Values ≤ 1 remove the gridSpan element.
func (pr *CT_TcPr) SetGridSpanVal(v int) error {
	pr.RemoveGridSpan()
	if v > 1 {
		if err := pr.GetOrAddGridSpan().SetVal(v); err != nil {
			return err
		}
	}
	return nil
}

// VMergeValStr returns the vMerge value as a string pointer.
// nil means vMerge is not present. "continue" or "restart".
func (pr *CT_TcPr) VMergeValStr() *string {
	vm := pr.VMerge()
	if vm == nil {
		return nil
	}
	v := vm.Val()
	// Per OOXML spec §17.4.85: when w:vMerge is present without a val
	// attribute, the default is "continue" (CT_VMerge val is optional,
	// default=continue). Python handles this via OptionalAttribute(default=ST_Merge.CONTINUE).
	if v == "" {
		s := "continue"
		return &s
	}
	return &v
}

// SetVMergeValStr sets the vMerge value. nil removes vMerge.
func (pr *CT_TcPr) SetVMergeValStr(v *string) error {
	pr.RemoveVMerge()
	if v == nil {
		return nil
	}
	vm := pr.GetOrAddVMerge()
	if err := vm.SetVal(*v); err != nil {
		return err
	}
	return nil
}

// VAlignValEnum returns the vertical alignment enum, or nil.
func (pr *CT_TcPr) VAlignValEnum() (*enum.WdCellVerticalAlignment, error) {
	va := pr.VAlign()
	if va == nil {
		return nil, nil
	}
	v, err := va.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetVAlignValEnum sets the vertical alignment. nil removes vAlign.
func (pr *CT_TcPr) SetVAlignValEnum(v *enum.WdCellVerticalAlignment) error {
	if v == nil {
		pr.RemoveVAlign()
		return nil
	}
	return pr.GetOrAddVAlign().SetVal(*v)
}

// WidthTwips returns the cell width in twips from tcW, or nil if not dxa or absent.
func (pr *CT_TcPr) WidthTwips() (*int, error) {
	tcW := pr.TcW()
	if tcW == nil {
		return nil, nil
	}
	return tcW.WidthTwips()
}

// SetWidthTwips sets the cell width to dxa type with the given twips value.
func (pr *CT_TcPr) SetWidthTwips(twips int) error {
	tcW := pr.GetOrAddTcW()
	if err := tcW.SetType("dxa"); err != nil {
		return err
	}
	if err := tcW.SetW(twips); err != nil {
		return err
	}
	return nil
}

// ===========================================================================
// CT_TblWidth — custom methods
// ===========================================================================

// WidthTwips returns the width in twips if type is "dxa", otherwise nil.
func (tw *CT_TblWidth) WidthTwips() (*int, error) {
	t, err := tw.Type()
	if err != nil {
		return nil, err
	}
	if t != "dxa" {
		return nil, nil
	}
	w, err := tw.W()
	if err != nil {
		return nil, err
	}
	return &w, nil
}

// SetWidthDxa sets the width in dxa (twips) and type to "dxa".
func (tw *CT_TblWidth) SetWidthDxa(twips int) error {
	if err := tw.SetType("dxa"); err != nil {
		return err
	}
	if err := tw.SetW(twips); err != nil {
		return err
	}
	return nil
}

// ===========================================================================
// CT_TblGridCol — custom methods
// ===========================================================================

// GridColIdx returns the index of this gridCol among its siblings.
func (gc *CT_TblGridCol) GridColIdx() int {
	parent := gc.e.Parent()
	if parent == nil {
		return -1
	}
	idx := 0
	for _, child := range parent.ChildElements() {
		if child.Space == "w" && child.Tag == "gridCol" {
			if child == gc.e {
				return idx
			}
			idx++
		}
	}
	return -1
}
