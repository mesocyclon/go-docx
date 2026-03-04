package enum

// ---------------------------------------------------------------------------
// WdStyleType
// ---------------------------------------------------------------------------

// WdStyleType specifies one of the four style types: paragraph, character, list, or table.
// MS API name: WdStyleType
type WdStyleType int

const (
	WdStyleTypeParagraph WdStyleType = 1
	WdStyleTypeCharacter WdStyleType = 2
	WdStyleTypeTable     WdStyleType = 3
	WdStyleTypeList      WdStyleType = 4
)

var wdStyleTypeToXml = map[WdStyleType]string{
	WdStyleTypeParagraph: "paragraph",
	WdStyleTypeCharacter: "character",
	WdStyleTypeTable:     "table",
	WdStyleTypeList:      "numbering",
}

var wdStyleTypeFromXml = invertMap(wdStyleTypeToXml)

// ToXml returns the XML attribute value for this style type.
func (v WdStyleType) ToXml() (string, error) { return ToXml(wdStyleTypeToXml, v) }

// WdStyleTypeFromXml returns the style type for the given XML value.
func WdStyleTypeFromXml(s string) (WdStyleType, error) {
	return FromXml(wdStyleTypeFromXml, s)
}

// ---------------------------------------------------------------------------
// WdBuiltinStyle (alias: WdStyle) — no XML mapping
// ---------------------------------------------------------------------------

// WdBuiltinStyle specifies a built-in Microsoft Word style.
// MS API name: WdBuiltinStyle
type WdBuiltinStyle int

// WdStyle is an alias for WdBuiltinStyle.
type WdStyle = WdBuiltinStyle

const (
	WdBuiltinStyleNormal                       WdBuiltinStyle = -1
	WdBuiltinStyleHeading1                     WdBuiltinStyle = -2
	WdBuiltinStyleHeading2                     WdBuiltinStyle = -3
	WdBuiltinStyleHeading3                     WdBuiltinStyle = -4
	WdBuiltinStyleHeading4                     WdBuiltinStyle = -5
	WdBuiltinStyleHeading5                     WdBuiltinStyle = -6
	WdBuiltinStyleHeading6                     WdBuiltinStyle = -7
	WdBuiltinStyleHeading7                     WdBuiltinStyle = -8
	WdBuiltinStyleHeading8                     WdBuiltinStyle = -9
	WdBuiltinStyleHeading9                     WdBuiltinStyle = -10
	WdBuiltinStyleIndex1                       WdBuiltinStyle = -11
	WdBuiltinStyleIndex2                       WdBuiltinStyle = -12
	WdBuiltinStyleIndex3                       WdBuiltinStyle = -13
	WdBuiltinStyleIndex4                       WdBuiltinStyle = -14
	WdBuiltinStyleIndex5                       WdBuiltinStyle = -15
	WdBuiltinStyleIndex6                       WdBuiltinStyle = -16
	WdBuiltinStyleIndex7                       WdBuiltinStyle = -17
	WdBuiltinStyleIndex8                       WdBuiltinStyle = -18
	WdBuiltinStyleIndex9                       WdBuiltinStyle = -19
	WdBuiltinStyleTOC1                         WdBuiltinStyle = -20
	WdBuiltinStyleTOC2                         WdBuiltinStyle = -21
	WdBuiltinStyleTOC3                         WdBuiltinStyle = -22
	WdBuiltinStyleTOC4                         WdBuiltinStyle = -23
	WdBuiltinStyleTOC5                         WdBuiltinStyle = -24
	WdBuiltinStyleTOC6                         WdBuiltinStyle = -25
	WdBuiltinStyleTOC7                         WdBuiltinStyle = -26
	WdBuiltinStyleTOC8                         WdBuiltinStyle = -27
	WdBuiltinStyleTOC9                         WdBuiltinStyle = -28
	WdBuiltinStyleNormalIndent                 WdBuiltinStyle = -29
	WdBuiltinStyleFootnoteText                 WdBuiltinStyle = -30
	WdBuiltinStyleCommentText                  WdBuiltinStyle = -31
	WdBuiltinStyleHeader                       WdBuiltinStyle = -32
	WdBuiltinStyleFooter                       WdBuiltinStyle = -33
	WdBuiltinStyleIndexHeading                 WdBuiltinStyle = -34
	WdBuiltinStyleCaption                      WdBuiltinStyle = -35
	WdBuiltinStyleTableOfFigures               WdBuiltinStyle = -36
	WdBuiltinStyleEnvelopeAddress              WdBuiltinStyle = -37
	WdBuiltinStyleEnvelopeReturn               WdBuiltinStyle = -38
	WdBuiltinStyleFootnoteReference            WdBuiltinStyle = -39
	WdBuiltinStyleCommentReference             WdBuiltinStyle = -40
	WdBuiltinStyleLineNumber                   WdBuiltinStyle = -41
	WdBuiltinStylePageNumber                   WdBuiltinStyle = -42
	WdBuiltinStyleEndnoteReference             WdBuiltinStyle = -43
	WdBuiltinStyleEndnoteText                  WdBuiltinStyle = -44
	WdBuiltinStyleTableOfAuthorities           WdBuiltinStyle = -45
	WdBuiltinStyleMacroText                    WdBuiltinStyle = -46
	WdBuiltinStyleTOAHeading                   WdBuiltinStyle = -47
	WdBuiltinStyleList                         WdBuiltinStyle = -48
	WdBuiltinStyleListBullet                   WdBuiltinStyle = -49
	WdBuiltinStyleListNumber                   WdBuiltinStyle = -50
	WdBuiltinStyleList2                        WdBuiltinStyle = -51
	WdBuiltinStyleList3                        WdBuiltinStyle = -52
	WdBuiltinStyleList4                        WdBuiltinStyle = -53
	WdBuiltinStyleList5                        WdBuiltinStyle = -54
	WdBuiltinStyleListBullet2                  WdBuiltinStyle = -55
	WdBuiltinStyleListBullet3                  WdBuiltinStyle = -56
	WdBuiltinStyleListBullet4                  WdBuiltinStyle = -57
	WdBuiltinStyleListBullet5                  WdBuiltinStyle = -58
	WdBuiltinStyleListNumber2                  WdBuiltinStyle = -59
	WdBuiltinStyleListNumber3                  WdBuiltinStyle = -60
	WdBuiltinStyleListNumber4                  WdBuiltinStyle = -61
	WdBuiltinStyleListNumber5                  WdBuiltinStyle = -62
	WdBuiltinStyleTitle                        WdBuiltinStyle = -63
	WdBuiltinStyleClosing                      WdBuiltinStyle = -64
	WdBuiltinStyleSignature                    WdBuiltinStyle = -65
	WdBuiltinStyleDefaultParagraphFont         WdBuiltinStyle = -66
	WdBuiltinStyleBodyText                     WdBuiltinStyle = -67
	WdBuiltinStyleBodyTextIndent               WdBuiltinStyle = -68
	WdBuiltinStyleListContinue                 WdBuiltinStyle = -69
	WdBuiltinStyleListContinue2                WdBuiltinStyle = -70
	WdBuiltinStyleListContinue3                WdBuiltinStyle = -71
	WdBuiltinStyleListContinue4                WdBuiltinStyle = -72
	WdBuiltinStyleListContinue5                WdBuiltinStyle = -73
	WdBuiltinStyleMessageHeader                WdBuiltinStyle = -74
	WdBuiltinStyleSubtitle                     WdBuiltinStyle = -75
	WdBuiltinStyleSalutation                   WdBuiltinStyle = -76
	WdBuiltinStyleDate                         WdBuiltinStyle = -77
	WdBuiltinStyleBodyTextFirstIndent          WdBuiltinStyle = -78
	WdBuiltinStyleBodyTextFirstIndent2         WdBuiltinStyle = -79
	WdBuiltinStyleNoteHeading                  WdBuiltinStyle = -80
	WdBuiltinStyleBodyText2                    WdBuiltinStyle = -81
	WdBuiltinStyleBodyText3                    WdBuiltinStyle = -82
	WdBuiltinStyleBodyTextIndent2              WdBuiltinStyle = -83
	WdBuiltinStyleBodyTextIndent3              WdBuiltinStyle = -84
	WdBuiltinStyleBlockQuotation               WdBuiltinStyle = -85
	WdBuiltinStyleHyperlink                    WdBuiltinStyle = -86
	WdBuiltinStyleHyperlinkFollowed            WdBuiltinStyle = -87
	WdBuiltinStyleStrong                       WdBuiltinStyle = -88
	WdBuiltinStyleEmphasis                     WdBuiltinStyle = -89
	WdBuiltinStyleNavPane                      WdBuiltinStyle = -90
	WdBuiltinStylePlainText                    WdBuiltinStyle = -91
	WdBuiltinStyleHTMLNormal                   WdBuiltinStyle = -95
	WdBuiltinStyleHTMLAcronym                  WdBuiltinStyle = -96
	WdBuiltinStyleHTMLAddress                  WdBuiltinStyle = -97
	WdBuiltinStyleHTMLCite                     WdBuiltinStyle = -98
	WdBuiltinStyleHTMLCode                     WdBuiltinStyle = -99
	WdBuiltinStyleHTMLDfn                      WdBuiltinStyle = -100
	WdBuiltinStyleHTMLKbd                      WdBuiltinStyle = -101
	WdBuiltinStyleHTMLPre                      WdBuiltinStyle = -102
	WdBuiltinStyleHTMLSamp                     WdBuiltinStyle = -103
	WdBuiltinStyleHTMLTt                       WdBuiltinStyle = -104
	WdBuiltinStyleHTMLVar                      WdBuiltinStyle = -105
	WdBuiltinStyleNormalTable                  WdBuiltinStyle = -106
	WdBuiltinStyleNormalObject                 WdBuiltinStyle = -158
	WdBuiltinStyleTableLightShading            WdBuiltinStyle = -159
	WdBuiltinStyleTableLightList               WdBuiltinStyle = -160
	WdBuiltinStyleTableLightGrid               WdBuiltinStyle = -161
	WdBuiltinStyleTableMediumShading1          WdBuiltinStyle = -162
	WdBuiltinStyleTableMediumShading2          WdBuiltinStyle = -163
	WdBuiltinStyleTableMediumList1             WdBuiltinStyle = -164
	WdBuiltinStyleTableMediumList2             WdBuiltinStyle = -165
	WdBuiltinStyleTableMediumGrid1             WdBuiltinStyle = -166
	WdBuiltinStyleTableMediumGrid2             WdBuiltinStyle = -167
	WdBuiltinStyleTableMediumGrid3             WdBuiltinStyle = -168
	WdBuiltinStyleTableDarkList                WdBuiltinStyle = -169
	WdBuiltinStyleTableColorfulShading         WdBuiltinStyle = -170
	WdBuiltinStyleTableColorfulList            WdBuiltinStyle = -171
	WdBuiltinStyleTableColorfulGrid            WdBuiltinStyle = -172
	WdBuiltinStyleTableLightShadingAccent1     WdBuiltinStyle = -173
	WdBuiltinStyleTableLightListAccent1        WdBuiltinStyle = -174
	WdBuiltinStyleTableLightGridAccent1        WdBuiltinStyle = -175
	WdBuiltinStyleTableMediumShading1Accent1   WdBuiltinStyle = -176
	WdBuiltinStyleTableMediumShading2Accent1   WdBuiltinStyle = -177
	WdBuiltinStyleTableMediumList1Accent1      WdBuiltinStyle = -178
	WdBuiltinStyleListParagraph                WdBuiltinStyle = -180
	WdBuiltinStyleQuote                        WdBuiltinStyle = -181
	WdBuiltinStyleIntenseQuote                 WdBuiltinStyle = -182
	WdBuiltinStyleSubtleEmphasis               WdBuiltinStyle = -261
	WdBuiltinStyleIntenseEmphasis              WdBuiltinStyle = -262
	WdBuiltinStyleSubtleReference              WdBuiltinStyle = -263
	WdBuiltinStyleIntenseReference             WdBuiltinStyle = -264
	WdBuiltinStyleBookTitle                    WdBuiltinStyle = -265
)
