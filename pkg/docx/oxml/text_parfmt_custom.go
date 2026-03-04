package oxml

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// --- CT_PPr custom methods ---

// JcVal returns the paragraph justification value, or nil if not set.
func (pPr *CT_PPr) JcVal() (*enum.WdParagraphAlignment, error) {
	jc := pPr.Jc()
	if jc == nil {
		return nil, nil
	}
	v, err := jc.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetJcVal sets the justification value. Passing nil removes the jc element.
func (pPr *CT_PPr) SetJcVal(v *enum.WdParagraphAlignment) error {
	if v == nil {
		pPr.RemoveJc()
		return nil
	}
	return pPr.GetOrAddJc().SetVal(*v)
}

// --- Style ---

// StyleVal returns the paragraph style string, or nil if not set.
func (pPr *CT_PPr) StyleVal() (*string, error) {
	pStyle := pPr.PStyle()
	if pStyle == nil {
		return nil, nil
	}
	v, err := pStyle.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetStyleVal sets the paragraph style. Passing nil removes pStyle.
func (pPr *CT_PPr) SetStyleVal(v *string) error {
	if v == nil {
		pPr.RemovePStyle()
		return nil
	}
	if err := pPr.GetOrAddPStyle().SetVal(*v); err != nil {
		return err
	}
	return nil
}

// --- Spacing properties ---

// SpacingBefore returns the value of w:spacing/@w:before in twips, or nil if not present.
func (pPr *CT_PPr) SpacingBefore() (*int, error) {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil, nil
	}
	v, ok := spacing.GetAttr("w:before")
	if !ok {
		return nil, nil
	}
	i, err := parseIntAttr(v)
	if err != nil {
		return nil, err
	}
	return &i, nil
}

// SetSpacingBefore sets the w:spacing/@w:before value in twips.
// Passing nil removes the attribute (creates spacing element if needed for other attrs).
func (pPr *CT_PPr) SetSpacingBefore(v *int) error {
	if v == nil && pPr.Spacing() == nil {
		return nil
	}
	spacing := pPr.GetOrAddSpacing()
	return spacing.SetBefore(v)
}

// SpacingAfter returns the value of w:spacing/@w:after in twips, or nil if not present.
func (pPr *CT_PPr) SpacingAfter() (*int, error) {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil, nil
	}
	v, ok := spacing.GetAttr("w:after")
	if !ok {
		return nil, nil
	}
	i, err := parseIntAttr(v)
	if err != nil {
		return nil, err
	}
	return &i, nil
}

// SetSpacingAfter sets the w:spacing/@w:after value in twips.
func (pPr *CT_PPr) SetSpacingAfter(v *int) error {
	if v == nil && pPr.Spacing() == nil {
		return nil
	}
	spacing := pPr.GetOrAddSpacing()
	return spacing.SetAfter(v)
}

// SpacingLine returns the value of w:spacing/@w:line in twips, or nil if not present.
func (pPr *CT_PPr) SpacingLine() (*int, error) {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil, nil
	}
	v, ok := spacing.GetAttr("w:line")
	if !ok {
		return nil, nil
	}
	i, err := parseIntAttr(v)
	if err != nil {
		return nil, err
	}
	return &i, nil
}

// SetSpacingLine sets the w:spacing/@w:line value.
func (pPr *CT_PPr) SetSpacingLine(v *int) error {
	if v == nil && pPr.Spacing() == nil {
		return nil
	}
	spacing := pPr.GetOrAddSpacing()
	return spacing.SetLine(v)
}

// SpacingLineRule returns the line spacing rule, or nil if not present.
// Defaults to WdLineSpacingMultiple if spacing/@w:line is present but lineRule is absent.
func (pPr *CT_PPr) SpacingLineRule() (*enum.WdLineSpacing, error) {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil, nil
	}
	lr := spacing.LineRule()
	if lr == "" {
		// Check if line is present; if so, default to MULTIPLE
		_, hasLine := spacing.GetAttr("w:line")
		if hasLine {
			v := enum.WdLineSpacingMultiple
			return &v, nil
		}
		return nil, nil
	}
	v, err := enum.WdLineSpacingFromXml(lr)
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetSpacingLineRule sets the line spacing rule.
func (pPr *CT_PPr) SetSpacingLineRule(v *enum.WdLineSpacing) error {
	if v == nil && pPr.Spacing() == nil {
		return nil
	}
	spacing := pPr.GetOrAddSpacing()
	if v == nil {
		if err := spacing.SetLineRule(""); err != nil {
			return err
		}
	} else {
		xml, err := v.ToXml()
		if err == nil {
			if err := spacing.SetLineRule(xml); err != nil {
				return err
			}
		}
	}
	return nil
}

// --- Indentation properties ---

// IndLeft returns the value of w:ind/@w:left in twips, or nil if not present.
func (pPr *CT_PPr) IndLeft() (*int, error) {
	ind := pPr.Ind()
	if ind == nil {
		return nil, nil
	}
	return ind.Left()
}

// SetIndLeft sets the w:ind/@w:left in twips.
func (pPr *CT_PPr) SetIndLeft(v *int) error {
	if v == nil && pPr.Ind() == nil {
		return nil
	}
	ind := pPr.GetOrAddInd()
	return ind.SetLeft(v)
}

// IndRight returns the value of w:ind/@w:right in twips, or nil if not present.
func (pPr *CT_PPr) IndRight() (*int, error) {
	ind := pPr.Ind()
	if ind == nil {
		return nil, nil
	}
	return ind.Right()
}

// SetIndRight sets the w:ind/@w:right in twips.
func (pPr *CT_PPr) SetIndRight(v *int) error {
	if v == nil && pPr.Ind() == nil {
		return nil
	}
	ind := pPr.GetOrAddInd()
	return ind.SetRight(v)
}

// FirstLineIndent returns a calculated indentation from w:ind/@w:firstLine and
// w:ind/@w:hanging. A hanging indent is returned as negative.
// Returns nil if no w:ind element.
func (pPr *CT_PPr) FirstLineIndent() (*int, error) {
	ind := pPr.Ind()
	if ind == nil {
		return nil, nil
	}
	_, hasHanging := ind.GetAttr("w:hanging")
	if hasHanging {
		h, err := ind.Hanging()
		if err != nil {
			return nil, err
		}
		if h == nil {
			return nil, nil
		}
		v := -(*h)
		return &v, nil
	}
	_, hasFirstLine := ind.GetAttr("w:firstLine")
	if !hasFirstLine {
		return nil, nil
	}
	return ind.FirstLine()
}

// SetFirstLineIndent sets the first-line indent. Negative values become hanging indents.
// nil clears both firstLine and hanging.
func (pPr *CT_PPr) SetFirstLineIndent(v *int) error {
	if pPr.Ind() == nil && v == nil {
		return nil
	}
	ind := pPr.GetOrAddInd()
	if err := ind.SetFirstLine(nil); err != nil {
		return err
	}
	if err := ind.SetHanging(nil); err != nil {
		return err
	}
	if v == nil {
		return nil
	}
	if *v < 0 {
		neg := -(*v)
		if err := ind.SetHanging(&neg); err != nil {
			return err
		}
	} else {
		if err := ind.SetFirstLine(v); err != nil {
			return err
		}
	}
	return nil
}

// --- Paragraph formatting booleans (keepLines, keepNext, pageBreakBefore, widowControl) ---

// pPrBoolVal reads a tri-state from a CT_OnOff child by tag.
func (pPr *CT_PPr) pPrBoolVal(tag string) *bool {
	child := pPr.FindChild(tag)
	if child == nil {
		return nil
	}
	onOff := &CT_OnOff{Element{e: child}}
	v := onOff.Val()
	return &v
}

// KeepLinesVal returns the tri-state keepLines value.
func (pPr *CT_PPr) KeepLinesVal() *bool {
	return pPr.pPrBoolVal("w:keepLines")
}

// SetKeepLinesVal sets keepLines. nil removes the element.
func (pPr *CT_PPr) SetKeepLinesVal(v *bool) error {
	if v == nil {
		pPr.RemoveKeepLines()
	} else {
		if err := pPr.GetOrAddKeepLines().SetVal(*v); err != nil {
			return err
		}
	}
	return nil
}

// KeepNextVal returns the tri-state keepNext value.
func (pPr *CT_PPr) KeepNextVal() *bool {
	return pPr.pPrBoolVal("w:keepNext")
}

// SetKeepNextVal sets keepNext. nil removes the element.
func (pPr *CT_PPr) SetKeepNextVal(v *bool) error {
	if v == nil {
		pPr.RemoveKeepNext()
	} else {
		if err := pPr.GetOrAddKeepNext().SetVal(*v); err != nil {
			return err
		}
	}
	return nil
}

// PageBreakBeforeVal returns the tri-state pageBreakBefore value.
func (pPr *CT_PPr) PageBreakBeforeVal() *bool {
	return pPr.pPrBoolVal("w:pageBreakBefore")
}

// SetPageBreakBeforeVal sets pageBreakBefore. nil removes the element.
func (pPr *CT_PPr) SetPageBreakBeforeVal(v *bool) error {
	if v == nil {
		pPr.RemovePageBreakBefore()
	} else {
		if err := pPr.GetOrAddPageBreakBefore().SetVal(*v); err != nil {
			return err
		}
	}
	return nil
}

// WidowControlVal returns the tri-state widowControl value.
func (pPr *CT_PPr) WidowControlVal() *bool {
	return pPr.pPrBoolVal("w:widowControl")
}

// SetWidowControlVal sets widowControl. nil removes the element.
func (pPr *CT_PPr) SetWidowControlVal(v *bool) error {
	if v == nil {
		pPr.RemoveWidowControl()
	} else {
		if err := pPr.GetOrAddWidowControl().SetVal(*v); err != nil {
			return err
		}
	}
	return nil
}

// --- CT_TabStops custom methods ---

// InsertTabInOrder inserts a new <w:tab> child element in position order.
func (tabs *CT_TabStops) InsertTabInOrder(pos int, align enum.WdTabAlignment, leader enum.WdTabLeader) (*CT_TabStop, error) {
	newTab := tabs.newTab()
	if err := newTab.SetPos(pos); err != nil {
		return nil, err
	}
	if err := newTab.SetVal(align); err != nil {
		return nil, fmt.Errorf("InsertTabInOrder: %w", err)
	}
	if err := newTab.SetLeader(leader); err != nil {
		return nil, fmt.Errorf("InsertTabInOrder: %w", err)
	}

	for _, tab := range tabs.TabList() {
		tabPos, err := tab.Pos()
		if err == nil && pos < tabPos {
			insertBefore(tabs.e, newTab.e, tab.e)
			return newTab, nil
		}
	}
	tabs.e.AddChild(newTab.e)
	return newTab, nil
}
