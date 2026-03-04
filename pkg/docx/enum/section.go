package enum

// ---------------------------------------------------------------------------
// WdHeaderFooterIndex (alias: WdHeaderFooter)
// ---------------------------------------------------------------------------

// WdHeaderFooterIndex specifies one of the three possible header/footer definitions.
// MS API name: WdHeaderFooterIndex
type WdHeaderFooterIndex int

const (
	WdHeaderFooterIndexPrimary   WdHeaderFooterIndex = 1
	WdHeaderFooterIndexFirstPage WdHeaderFooterIndex = 2
	WdHeaderFooterIndexEvenPage  WdHeaderFooterIndex = 3
)

// WdHeaderFooter is an alias for WdHeaderFooterIndex.
type WdHeaderFooter = WdHeaderFooterIndex

var wdHeaderFooterIndexToXml = map[WdHeaderFooterIndex]string{
	WdHeaderFooterIndexPrimary:   "default",
	WdHeaderFooterIndexFirstPage: "first",
	WdHeaderFooterIndexEvenPage:  "even",
}

var wdHeaderFooterIndexFromXml = invertMap(wdHeaderFooterIndexToXml)

// ToXml returns the XML attribute value for this header/footer index.
func (v WdHeaderFooterIndex) ToXml() (string, error) { return ToXml(wdHeaderFooterIndexToXml, v) }

// WdHeaderFooterIndexFromXml returns the header/footer index for the given XML value.
func WdHeaderFooterIndexFromXml(s string) (WdHeaderFooterIndex, error) {
	return FromXml(wdHeaderFooterIndexFromXml, s)
}

// ---------------------------------------------------------------------------
// WdOrientation (alias: WdOrient)
// ---------------------------------------------------------------------------

// WdOrientation specifies the page layout orientation.
// MS API name: WdOrientation
type WdOrientation int

const (
	WdOrientationPortrait  WdOrientation = 0
	WdOrientationLandscape WdOrientation = 1
)

// WdOrient is an alias for WdOrientation.
type WdOrient = WdOrientation

var wdOrientationToXml = map[WdOrientation]string{
	WdOrientationPortrait:  "portrait",
	WdOrientationLandscape: "landscape",
}

var wdOrientationFromXml = invertMap(wdOrientationToXml)

// ToXml returns the XML attribute value for this orientation.
func (v WdOrientation) ToXml() (string, error) { return ToXml(wdOrientationToXml, v) }

// WdOrientationFromXml returns the orientation for the given XML value.
func WdOrientationFromXml(s string) (WdOrientation, error) {
	return FromXml(wdOrientationFromXml, s)
}

// ---------------------------------------------------------------------------
// WdSectionStart (alias: WdSection)
// ---------------------------------------------------------------------------

// WdSectionStart specifies the start type of a section break.
// MS API name: WdSectionStart
type WdSectionStart int

const (
	WdSectionStartContinuous WdSectionStart = 0
	WdSectionStartNewColumn  WdSectionStart = 1
	WdSectionStartNewPage    WdSectionStart = 2
	WdSectionStartEvenPage   WdSectionStart = 3
	WdSectionStartOddPage    WdSectionStart = 4
)

// WdSection is an alias for WdSectionStart.
type WdSection = WdSectionStart

var wdSectionStartToXml = map[WdSectionStart]string{
	WdSectionStartContinuous: "continuous",
	WdSectionStartNewColumn:  "nextColumn",
	WdSectionStartNewPage:    "nextPage",
	WdSectionStartEvenPage:   "evenPage",
	WdSectionStartOddPage:    "oddPage",
}

var wdSectionStartFromXml = invertMap(wdSectionStartToXml)

// ToXml returns the XML attribute value for this section start type.
func (v WdSectionStart) ToXml() (string, error) { return ToXml(wdSectionStartToXml, v) }

// WdSectionStartFromXml returns the section start type for the given XML value.
func WdSectionStartFromXml(s string) (WdSectionStart, error) {
	return FromXml(wdSectionStartFromXml, s)
}
