package enum

import "fmt"

// ---------------------------------------------------------------------------
// WdParagraphAlignment (alias: WdAlignParagraph)
// ---------------------------------------------------------------------------

// WdParagraphAlignment specifies paragraph justification type.
// MS API name: WdParagraphAlignment
type WdParagraphAlignment int

const (
	WdParagraphAlignmentLeft       WdParagraphAlignment = 0
	WdParagraphAlignmentCenter     WdParagraphAlignment = 1
	WdParagraphAlignmentRight      WdParagraphAlignment = 2
	WdParagraphAlignmentJustify    WdParagraphAlignment = 3
	WdParagraphAlignmentDistribute WdParagraphAlignment = 4
	WdParagraphAlignmentJustifyMed WdParagraphAlignment = 5
	WdParagraphAlignmentJustifyHi  WdParagraphAlignment = 7
	WdParagraphAlignmentJustifyLow WdParagraphAlignment = 8
	WdParagraphAlignmentThaiJustify WdParagraphAlignment = 9
)

// WdAlignParagraph is an alias for WdParagraphAlignment.
type WdAlignParagraph = WdParagraphAlignment

var wdParagraphAlignmentToXml = map[WdParagraphAlignment]string{
	WdParagraphAlignmentLeft:        "left",
	WdParagraphAlignmentCenter:      "center",
	WdParagraphAlignmentRight:       "right",
	WdParagraphAlignmentJustify:     "both",
	WdParagraphAlignmentDistribute:  "distribute",
	WdParagraphAlignmentJustifyMed:  "mediumKashida",
	WdParagraphAlignmentJustifyHi:   "highKashida",
	WdParagraphAlignmentJustifyLow:  "lowKashida",
	WdParagraphAlignmentThaiJustify: "thaiDistribute",
}

var wdParagraphAlignmentFromXml = invertMap(wdParagraphAlignmentToXml)

// ToXml returns the XML attribute value for this alignment.
func (v WdParagraphAlignment) ToXml() (string, error) {
	return ToXml(wdParagraphAlignmentToXml, v)
}

// WdParagraphAlignmentFromXml returns the alignment for the given XML value.
func WdParagraphAlignmentFromXml(s string) (WdParagraphAlignment, error) {
	return FromXml(wdParagraphAlignmentFromXml, s)
}

// ---------------------------------------------------------------------------
// WdBreakType
// ---------------------------------------------------------------------------

// WdBreakType specifies the type of break.
// MS API name: WdBreakType
type WdBreakType int

const (
	WdBreakTypeColumn             WdBreakType = 8
	WdBreakTypeLine               WdBreakType = 6
	WdBreakTypeLineClearLeft      WdBreakType = 9
	WdBreakTypeLineClearRight     WdBreakType = 10
	WdBreakTypeLineClearAll       WdBreakType = 11
	WdBreakTypePage               WdBreakType = 7
	WdBreakTypeSectionContinuous  WdBreakType = 3
	WdBreakTypeSectionEvenPage    WdBreakType = 4
	WdBreakTypeSectionNextPage    WdBreakType = 2
	WdBreakTypeSectionOddPage     WdBreakType = 5
	WdBreakTypeTextWrapping       WdBreakType = 11
)

// WdBreak is an alias for WdBreakType.
type WdBreak = WdBreakType

// ---------------------------------------------------------------------------
// WdColorIndex
// ---------------------------------------------------------------------------

// WdColorIndex specifies a standard preset color for font highlighting.
// MS API name: WdColorIndex
type WdColorIndex int

const (
	WdColorIndexInherited   WdColorIndex = -1
	WdColorIndexAuto        WdColorIndex = 0
	WdColorIndexBlack       WdColorIndex = 1
	WdColorIndexBlue        WdColorIndex = 2
	WdColorIndexTurquoise   WdColorIndex = 3
	WdColorIndexBrightGreen WdColorIndex = 4
	WdColorIndexPink        WdColorIndex = 5
	WdColorIndexRed         WdColorIndex = 6
	WdColorIndexYellow      WdColorIndex = 7
	WdColorIndexWhite       WdColorIndex = 8
	WdColorIndexDarkBlue    WdColorIndex = 9
	WdColorIndexTeal        WdColorIndex = 10
	WdColorIndexGreen       WdColorIndex = 11
	WdColorIndexViolet      WdColorIndex = 12
	WdColorIndexDarkRed     WdColorIndex = 13
	WdColorIndexDarkYellow  WdColorIndex = 14
	WdColorIndexGray50      WdColorIndex = 15
	WdColorIndexGray25      WdColorIndex = 16
)

// WdColor is an alias for WdColorIndex.
type WdColor = WdColorIndex

var wdColorIndexToXml = map[WdColorIndex]string{
	WdColorIndexAuto:        "default",
	WdColorIndexBlack:       "black",
	WdColorIndexBlue:        "blue",
	WdColorIndexTurquoise:   "cyan",
	WdColorIndexBrightGreen: "green",
	WdColorIndexPink:        "magenta",
	WdColorIndexRed:         "red",
	WdColorIndexYellow:      "yellow",
	WdColorIndexWhite:       "white",
	WdColorIndexDarkBlue:    "darkBlue",
	WdColorIndexTeal:        "darkCyan",
	WdColorIndexGreen:       "darkGreen",
	WdColorIndexViolet:      "darkMagenta",
	WdColorIndexDarkRed:     "darkRed",
	WdColorIndexDarkYellow:  "darkYellow",
	WdColorIndexGray50:      "darkGray",
	WdColorIndexGray25:      "lightGray",
}

var wdColorIndexFromXml = invertMap(wdColorIndexToXml)

// ToXml returns the XML attribute value for this color index.
// INHERITED has no XML representation.
func (v WdColorIndex) ToXml() (string, error) { return ToXml(wdColorIndexToXml, v) }

// WdColorIndexFromXml returns the color index for the given XML value.
func WdColorIndexFromXml(s string) (WdColorIndex, error) {
	return FromXml(wdColorIndexFromXml, s)
}

// ---------------------------------------------------------------------------
// WdLineSpacing
// ---------------------------------------------------------------------------

// WdLineSpacing specifies a line spacing format to be applied to a paragraph.
// MS API name: WdLineSpacing
type WdLineSpacing int

const (
	WdLineSpacingSingle       WdLineSpacing = 0
	WdLineSpacingOnePointFive WdLineSpacing = 1
	WdLineSpacingDouble       WdLineSpacing = 2
	WdLineSpacingAtLeast      WdLineSpacing = 3
	WdLineSpacingExactly      WdLineSpacing = 4
	WdLineSpacingMultiple     WdLineSpacing = 5
)

var wdLineSpacingToXml = map[WdLineSpacing]string{
	WdLineSpacingAtLeast:  "atLeast",
	WdLineSpacingExactly:  "exact",
	WdLineSpacingMultiple: "auto",
}

var wdLineSpacingFromXml = invertMap(wdLineSpacingToXml)

// ToXml returns the XML attribute value for this line spacing rule.
// SINGLE, ONE_POINT_FIVE, and DOUBLE have no direct XML mapping (they are UNMAPPED).
func (v WdLineSpacing) ToXml() (string, error) {
	return ToXml(wdLineSpacingToXml, v)
}

// WdLineSpacingFromXml returns the line spacing rule for the given XML value.
func WdLineSpacingFromXml(s string) (WdLineSpacing, error) {
	return FromXml(wdLineSpacingFromXml, s)
}

// ---------------------------------------------------------------------------
// WdTabAlignment
// ---------------------------------------------------------------------------

// WdTabAlignment specifies the tab stop alignment.
// MS API name: WdTabAlignment
type WdTabAlignment int

const (
	WdTabAlignmentLeft    WdTabAlignment = 0
	WdTabAlignmentCenter  WdTabAlignment = 1
	WdTabAlignmentRight   WdTabAlignment = 2
	WdTabAlignmentDecimal WdTabAlignment = 3
	WdTabAlignmentBar     WdTabAlignment = 4
	WdTabAlignmentList    WdTabAlignment = 6
	WdTabAlignmentClear   WdTabAlignment = 101
	WdTabAlignmentEnd     WdTabAlignment = 102
	WdTabAlignmentNum     WdTabAlignment = 103
	WdTabAlignmentStart   WdTabAlignment = 104
)

var wdTabAlignmentToXml = map[WdTabAlignment]string{
	WdTabAlignmentLeft:    "left",
	WdTabAlignmentCenter:  "center",
	WdTabAlignmentRight:   "right",
	WdTabAlignmentDecimal: "decimal",
	WdTabAlignmentBar:     "bar",
	WdTabAlignmentList:    "list",
	WdTabAlignmentClear:   "clear",
	WdTabAlignmentEnd:     "end",
	WdTabAlignmentNum:     "num",
	WdTabAlignmentStart:   "start",
}

var wdTabAlignmentFromXml = invertMap(wdTabAlignmentToXml)

// ToXml returns the XML attribute value for this tab alignment.
func (v WdTabAlignment) ToXml() (string, error) { return ToXml(wdTabAlignmentToXml, v) }

// WdTabAlignmentFromXml returns the tab alignment for the given XML value.
func WdTabAlignmentFromXml(s string) (WdTabAlignment, error) {
	return FromXml(wdTabAlignmentFromXml, s)
}

// ---------------------------------------------------------------------------
// WdTabLeader
// ---------------------------------------------------------------------------

// WdTabLeader specifies the character to use as the leader with formatted tabs.
// MS API name: WdTabLeader
type WdTabLeader int

const (
	WdTabLeaderSpaces    WdTabLeader = 0
	WdTabLeaderDots      WdTabLeader = 1
	WdTabLeaderDashes    WdTabLeader = 2
	WdTabLeaderLines     WdTabLeader = 3
	WdTabLeaderHeavy     WdTabLeader = 4
	WdTabLeaderMiddleDot WdTabLeader = 5
)

var wdTabLeaderToXml = map[WdTabLeader]string{
	WdTabLeaderSpaces:    "none",
	WdTabLeaderDots:      "dot",
	WdTabLeaderDashes:    "hyphen",
	WdTabLeaderLines:     "underscore",
	WdTabLeaderHeavy:     "heavy",
	WdTabLeaderMiddleDot: "middleDot",
}

var wdTabLeaderFromXml = invertMap(wdTabLeaderToXml)

// ToXml returns the XML attribute value for this tab leader.
func (v WdTabLeader) ToXml() (string, error) { return ToXml(wdTabLeaderToXml, v) }

// WdTabLeaderFromXml returns the tab leader for the given XML value.
func WdTabLeaderFromXml(s string) (WdTabLeader, error) {
	return FromXml(wdTabLeaderFromXml, s)
}

// ---------------------------------------------------------------------------
// WdUnderline
// ---------------------------------------------------------------------------

// WdUnderline specifies the style of underline applied to a run of characters.
// MS API name: WdUnderline
type WdUnderline int

const (
	WdUnderlineInherited       WdUnderline = -1
	WdUnderlineNone            WdUnderline = 0
	WdUnderlineSingle          WdUnderline = 1
	WdUnderlineWords           WdUnderline = 2
	WdUnderlineDouble          WdUnderline = 3
	WdUnderlineDotted          WdUnderline = 4
	WdUnderlineThick           WdUnderline = 6
	WdUnderlineDash            WdUnderline = 7
	WdUnderlineDotDash         WdUnderline = 9
	WdUnderlineDotDotDash      WdUnderline = 10
	WdUnderlineWavy            WdUnderline = 11
	WdUnderlineDottedHeavy     WdUnderline = 20
	WdUnderlineDashHeavy       WdUnderline = 23
	WdUnderlineDotDashHeavy    WdUnderline = 25
	WdUnderlineDotDotDashHeavy WdUnderline = 26
	WdUnderlineWavyHeavy       WdUnderline = 27
	WdUnderlineDashLong        WdUnderline = 39
	WdUnderlineWavyDouble      WdUnderline = 43
	WdUnderlineDashLongHeavy   WdUnderline = 55
)

var wdUnderlineToXml = map[WdUnderline]string{
	WdUnderlineNone:            "none",
	WdUnderlineSingle:          "single",
	WdUnderlineWords:           "words",
	WdUnderlineDouble:          "double",
	WdUnderlineDotted:          "dotted",
	WdUnderlineThick:           "thick",
	WdUnderlineDash:            "dash",
	WdUnderlineDotDash:         "dotDash",
	WdUnderlineDotDotDash:      "dotDotDash",
	WdUnderlineWavy:            "wave",
	WdUnderlineDottedHeavy:     "dottedHeavy",
	WdUnderlineDashHeavy:       "dashedHeavy",
	WdUnderlineDotDashHeavy:    "dashDotHeavy",
	WdUnderlineDotDotDashHeavy: "dashDotDotHeavy",
	WdUnderlineWavyHeavy:       "wavyHeavy",
	WdUnderlineDashLong:        "dashLong",
	WdUnderlineWavyDouble:      "wavyDouble",
	WdUnderlineDashLongHeavy:   "dashLongHeavy",
}

var wdUnderlineFromXml = invertMap(wdUnderlineToXml)

// ToXml returns the XML attribute value for this underline style.
// INHERITED has no XML representation.
func (v WdUnderline) ToXml() (string, error) {
	return ToXml(wdUnderlineToXml, v)
}

// WdUnderlineFromXml returns the underline style for the given XML value.
func WdUnderlineFromXml(s string) (WdUnderline, error) {
	return FromXml(wdUnderlineFromXml, s)
}

// String returns a human-readable name for the WdUnderline value.
func (v WdUnderline) String() string {
	names := map[WdUnderline]string{
		WdUnderlineInherited:       "INHERITED",
		WdUnderlineNone:            "NONE",
		WdUnderlineSingle:          "SINGLE",
		WdUnderlineWords:           "WORDS",
		WdUnderlineDouble:          "DOUBLE",
		WdUnderlineDotted:          "DOTTED",
		WdUnderlineThick:           "THICK",
		WdUnderlineDash:            "DASH",
		WdUnderlineDotDash:         "DOT_DASH",
		WdUnderlineDotDotDash:      "DOT_DOT_DASH",
		WdUnderlineWavy:            "WAVY",
		WdUnderlineDottedHeavy:     "DOTTED_HEAVY",
		WdUnderlineDashHeavy:       "DASH_HEAVY",
		WdUnderlineDotDashHeavy:    "DOT_DASH_HEAVY",
		WdUnderlineDotDotDashHeavy: "DOT_DOT_DASH_HEAVY",
		WdUnderlineWavyHeavy:       "WAVY_HEAVY",
		WdUnderlineDashLong:        "DASH_LONG",
		WdUnderlineWavyDouble:      "WAVY_DOUBLE",
		WdUnderlineDashLongHeavy:   "DASH_LONG_HEAVY",
	}
	if name, ok := names[v]; ok {
		return fmt.Sprintf("WdUnderline.%s (%d)", name, int(v))
	}
	return fmt.Sprintf("WdUnderline(%d)", int(v))
}
