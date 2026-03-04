package enum

// ---------------------------------------------------------------------------
// MsoColorType — no XML mapping (BaseEnum equivalent)
// ---------------------------------------------------------------------------

// MsoColorType specifies the color specification scheme.
// MS API name: MsoColorType
type MsoColorType int

const (
	MsoColorTypeRGB   MsoColorType = 1
	MsoColorTypeTheme MsoColorType = 2
	MsoColorTypeAuto  MsoColorType = 101
)

// ---------------------------------------------------------------------------
// MsoThemeColorIndex (alias: MsoThemeColor)
// ---------------------------------------------------------------------------

// MsoThemeColorIndex indicates the Office theme color.
// MS API name: MsoThemeColorIndex
type MsoThemeColorIndex int

const (
	MsoThemeColorIndexNotThemeColor    MsoThemeColorIndex = 0
	MsoThemeColorIndexDark1            MsoThemeColorIndex = 1
	MsoThemeColorIndexLight1           MsoThemeColorIndex = 2
	MsoThemeColorIndexDark2            MsoThemeColorIndex = 3
	MsoThemeColorIndexLight2           MsoThemeColorIndex = 4
	MsoThemeColorIndexAccent1          MsoThemeColorIndex = 5
	MsoThemeColorIndexAccent2          MsoThemeColorIndex = 6
	MsoThemeColorIndexAccent3          MsoThemeColorIndex = 7
	MsoThemeColorIndexAccent4          MsoThemeColorIndex = 8
	MsoThemeColorIndexAccent5          MsoThemeColorIndex = 9
	MsoThemeColorIndexAccent6          MsoThemeColorIndex = 10
	MsoThemeColorIndexHyperlink        MsoThemeColorIndex = 11
	MsoThemeColorIndexFollowedHyperlink MsoThemeColorIndex = 12
	MsoThemeColorIndexText1            MsoThemeColorIndex = 13
	MsoThemeColorIndexBackground1      MsoThemeColorIndex = 14
	MsoThemeColorIndexText2            MsoThemeColorIndex = 15
	MsoThemeColorIndexBackground2      MsoThemeColorIndex = 16
)

// MsoThemeColor is an alias for MsoThemeColorIndex.
type MsoThemeColor = MsoThemeColorIndex

var msoThemeColorIndexToXml = map[MsoThemeColorIndex]string{
	MsoThemeColorIndexDark1:             "dark1",
	MsoThemeColorIndexLight1:            "light1",
	MsoThemeColorIndexDark2:             "dark2",
	MsoThemeColorIndexLight2:            "light2",
	MsoThemeColorIndexAccent1:           "accent1",
	MsoThemeColorIndexAccent2:           "accent2",
	MsoThemeColorIndexAccent3:           "accent3",
	MsoThemeColorIndexAccent4:           "accent4",
	MsoThemeColorIndexAccent5:           "accent5",
	MsoThemeColorIndexAccent6:           "accent6",
	MsoThemeColorIndexHyperlink:         "hyperlink",
	MsoThemeColorIndexFollowedHyperlink: "followedHyperlink",
	MsoThemeColorIndexText1:             "text1",
	MsoThemeColorIndexBackground1:       "background1",
	MsoThemeColorIndexText2:             "text2",
	MsoThemeColorIndexBackground2:       "background2",
}

var msoThemeColorIndexFromXml = invertMap(msoThemeColorIndexToXml)

// ToXml returns the XML attribute value for this theme color.
// NOT_THEME_COLOR has no XML representation (UNMAPPED).
func (v MsoThemeColorIndex) ToXml() (string, error) {
	return ToXml(msoThemeColorIndexToXml, v)
}

// MsoThemeColorIndexFromXml returns the theme color index for the given XML value.
func MsoThemeColorIndexFromXml(s string) (MsoThemeColorIndex, error) {
	return FromXml(msoThemeColorIndexFromXml, s)
}
