package enum

// ---------------------------------------------------------------------------
// WdTableAlignment
// ---------------------------------------------------------------------------

// WdTableAlignment specifies table justification type.
// MS API name: WdRowAlignment
type WdTableAlignment int

const (
	WdTableAlignmentLeft   WdTableAlignment = 0
	WdTableAlignmentCenter WdTableAlignment = 1
	WdTableAlignmentRight  WdTableAlignment = 2
)

var wdTableAlignmentToXml = map[WdTableAlignment]string{
	WdTableAlignmentLeft:   "left",
	WdTableAlignmentCenter: "center",
	WdTableAlignmentRight:  "right",
}

var wdTableAlignmentFromXml = invertMap(wdTableAlignmentToXml)

// ToXml returns the XML attribute value for this table alignment.
func (v WdTableAlignment) ToXml() (string, error) { return ToXml(wdTableAlignmentToXml, v) }

// WdTableAlignmentFromXml returns the table alignment for the given XML value.
func WdTableAlignmentFromXml(s string) (WdTableAlignment, error) {
	return FromXml(wdTableAlignmentFromXml, s)
}

// ---------------------------------------------------------------------------
// WdTableDirection — no XML mapping (BaseEnum equivalent)
// ---------------------------------------------------------------------------

// WdTableDirection specifies the direction in which an application orders cells.
// MS API name: WdTableDirection
type WdTableDirection int

const (
	WdTableDirectionLTR WdTableDirection = 0
	WdTableDirectionRTL WdTableDirection = 1
)

// ---------------------------------------------------------------------------
// WdCellVerticalAlignment (alias: WdAlignVertical)
// ---------------------------------------------------------------------------

// WdCellVerticalAlignment specifies the vertical alignment of text in table cells.
// MS API name: WdCellVerticalAlignment
type WdCellVerticalAlignment int

const (
	WdCellVerticalAlignmentTop    WdCellVerticalAlignment = 0
	WdCellVerticalAlignmentCenter WdCellVerticalAlignment = 1
	WdCellVerticalAlignmentBottom WdCellVerticalAlignment = 3
	WdCellVerticalAlignmentBoth   WdCellVerticalAlignment = 101
)

// WdAlignVertical is an alias for WdCellVerticalAlignment.
type WdAlignVertical = WdCellVerticalAlignment

var wdCellVerticalAlignmentToXml = map[WdCellVerticalAlignment]string{
	WdCellVerticalAlignmentTop:    "top",
	WdCellVerticalAlignmentCenter: "center",
	WdCellVerticalAlignmentBottom: "bottom",
	WdCellVerticalAlignmentBoth:   "both",
}

var wdCellVerticalAlignmentFromXml = invertMap(wdCellVerticalAlignmentToXml)

// ToXml returns the XML attribute value for this vertical alignment.
func (v WdCellVerticalAlignment) ToXml() (string, error) {
	return ToXml(wdCellVerticalAlignmentToXml, v)
}

// WdCellVerticalAlignmentFromXml returns the vertical alignment for the given XML value.
func WdCellVerticalAlignmentFromXml(s string) (WdCellVerticalAlignment, error) {
	return FromXml(wdCellVerticalAlignmentFromXml, s)
}

// ---------------------------------------------------------------------------
// WdRowHeightRule (alias: WdRowHeight)
// ---------------------------------------------------------------------------

// WdRowHeightRule specifies the rule for determining the height of a table row.
// MS API name: WdRowHeightRule
type WdRowHeightRule int

const (
	WdRowHeightRuleAuto    WdRowHeightRule = 0
	WdRowHeightRuleAtLeast WdRowHeightRule = 1
	WdRowHeightRuleExactly WdRowHeightRule = 2
)

// WdRowHeight is an alias for WdRowHeightRule.
type WdRowHeight = WdRowHeightRule

var wdRowHeightRuleToXml = map[WdRowHeightRule]string{
	WdRowHeightRuleAuto:    "auto",
	WdRowHeightRuleAtLeast: "atLeast",
	WdRowHeightRuleExactly: "exact",
}

var wdRowHeightRuleFromXml = invertMap(wdRowHeightRuleToXml)

// ToXml returns the XML attribute value for this row height rule.
func (v WdRowHeightRule) ToXml() (string, error) { return ToXml(wdRowHeightRuleToXml, v) }

// WdRowHeightRuleFromXml returns the row height rule for the given XML value.
func WdRowHeightRuleFromXml(s string) (WdRowHeightRule, error) {
	return FromXml(wdRowHeightRuleFromXml, s)
}
