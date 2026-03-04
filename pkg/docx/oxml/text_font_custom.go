package oxml

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// --- CT_RPr custom methods ---

// getBoolVal reads a tri-state boolean value from a CT_OnOff child element.
//   - nil: element not present (inherit)
//   - *true: element present with val=true or val absent (e.g. <w:b/>)
//   - *false: element present with val=false (e.g. <w:b w:val="false"/>)
func (rPr *CT_RPr) getBoolVal(tag string) *bool {
	child := rPr.FindChild(tag)
	if child == nil {
		return nil
	}
	onOff := &CT_OnOff{Element{e: child}}
	v := onOff.Val()
	return &v
}

// setBoolVal sets a tri-state boolean for a CT_OnOff child element.
//   - nil: remove the element
//   - *true: add element without val attr (e.g. <w:b/>)
//   - *false: add element with val="false" (e.g. <w:b w:val="false"/>)
//
// getOrAdd and remove are function params to handle the correct child tag.
func (rPr *CT_RPr) setBoolValWith(val *bool, getOrAdd func() *CT_OnOff, remove func()) error {
	if val == nil {
		remove()
		return nil
	}
	el := getOrAdd()
	return el.SetVal(*val)
}

// --- Bold ---

// BoldVal returns the tri-state bold value: nil (inherit), *true, or *false.
func (rPr *CT_RPr) BoldVal() *bool {
	return rPr.getBoolVal("w:b")
}

// SetBoldVal sets the bold tri-state.
func (rPr *CT_RPr) SetBoldVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddB, rPr.RemoveB)
}

// --- Italic ---

// ItalicVal returns the tri-state italic value.
func (rPr *CT_RPr) ItalicVal() *bool {
	return rPr.getBoolVal("w:i")
}

// SetItalicVal sets the italic tri-state.
func (rPr *CT_RPr) SetItalicVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddI, rPr.RemoveI)
}

// --- Caps ---

// CapsVal returns the tri-state all-caps value.
func (rPr *CT_RPr) CapsVal() *bool {
	return rPr.getBoolVal("w:caps")
}

// SetCapsVal sets the all-caps tri-state.
func (rPr *CT_RPr) SetCapsVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddCaps, rPr.RemoveCaps)
}

// --- SmallCaps ---

// SmallCapsVal returns the tri-state small-caps value.
func (rPr *CT_RPr) SmallCapsVal() *bool {
	return rPr.getBoolVal("w:smallCaps")
}

// SetSmallCapsVal sets the small-caps tri-state.
func (rPr *CT_RPr) SetSmallCapsVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddSmallCaps, rPr.RemoveSmallCaps)
}

// --- Strike ---

// StrikeVal returns the tri-state strikethrough value.
func (rPr *CT_RPr) StrikeVal() *bool {
	return rPr.getBoolVal("w:strike")
}

// SetStrikeVal sets the strikethrough tri-state.
func (rPr *CT_RPr) SetStrikeVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddStrike, rPr.RemoveStrike)
}

// --- Dstrike (double strikethrough) ---

// DstrikeVal returns the tri-state double-strikethrough value.
func (rPr *CT_RPr) DstrikeVal() *bool {
	return rPr.getBoolVal("w:dstrike")
}

// SetDstrikeVal sets the double-strikethrough tri-state.
func (rPr *CT_RPr) SetDstrikeVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddDstrike, rPr.RemoveDstrike)
}

// --- Outline ---

// OutlineVal returns the tri-state outline value.
func (rPr *CT_RPr) OutlineVal() *bool {
	return rPr.getBoolVal("w:outline")
}

// SetOutlineVal sets the outline tri-state.
func (rPr *CT_RPr) SetOutlineVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddOutline, rPr.RemoveOutline)
}

// --- Shadow ---

// ShadowVal returns the tri-state shadow value.
func (rPr *CT_RPr) ShadowVal() *bool {
	return rPr.getBoolVal("w:shadow")
}

// SetShadowVal sets the shadow tri-state.
func (rPr *CT_RPr) SetShadowVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddShadow, rPr.RemoveShadow)
}

// --- Emboss ---

// EmbossVal returns the tri-state emboss value.
func (rPr *CT_RPr) EmbossVal() *bool {
	return rPr.getBoolVal("w:emboss")
}

// SetEmbossVal sets the emboss tri-state.
func (rPr *CT_RPr) SetEmbossVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddEmboss, rPr.RemoveEmboss)
}

// --- Imprint ---

// ImprintVal returns the tri-state imprint value.
func (rPr *CT_RPr) ImprintVal() *bool {
	return rPr.getBoolVal("w:imprint")
}

// SetImprintVal sets the imprint tri-state.
func (rPr *CT_RPr) SetImprintVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddImprint, rPr.RemoveImprint)
}

// --- NoProof ---

// NoProofVal returns the tri-state noProof value.
func (rPr *CT_RPr) NoProofVal() *bool {
	return rPr.getBoolVal("w:noProof")
}

// SetNoProofVal sets the noProof tri-state.
func (rPr *CT_RPr) SetNoProofVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddNoProof, rPr.RemoveNoProof)
}

// --- SnapToGrid ---

// SnapToGridVal returns the tri-state snapToGrid value.
func (rPr *CT_RPr) SnapToGridVal() *bool {
	return rPr.getBoolVal("w:snapToGrid")
}

// SetSnapToGridVal sets the snapToGrid tri-state.
func (rPr *CT_RPr) SetSnapToGridVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddSnapToGrid, rPr.RemoveSnapToGrid)
}

// --- Vanish ---

// VanishVal returns the tri-state vanish value.
func (rPr *CT_RPr) VanishVal() *bool {
	return rPr.getBoolVal("w:vanish")
}

// SetVanishVal sets the vanish tri-state.
func (rPr *CT_RPr) SetVanishVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddVanish, rPr.RemoveVanish)
}

// --- WebHidden ---

// WebHiddenVal returns the tri-state webHidden value.
func (rPr *CT_RPr) WebHiddenVal() *bool {
	return rPr.getBoolVal("w:webHidden")
}

// SetWebHiddenVal sets the webHidden tri-state.
func (rPr *CT_RPr) SetWebHiddenVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddWebHidden, rPr.RemoveWebHidden)
}

// --- SpecVanish ---

// SpecVanishVal returns the tri-state specVanish value.
func (rPr *CT_RPr) SpecVanishVal() *bool {
	return rPr.getBoolVal("w:specVanish")
}

// SetSpecVanishVal sets the specVanish tri-state.
func (rPr *CT_RPr) SetSpecVanishVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddSpecVanish, rPr.RemoveSpecVanish)
}

// --- OMath ---

// OMathVal returns the tri-state oMath value.
func (rPr *CT_RPr) OMathVal() *bool {
	return rPr.getBoolVal("w:oMath")
}

// SetOMathVal sets the oMath tri-state.
func (rPr *CT_RPr) SetOMathVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddOMath, rPr.RemoveOMath)
}

// --- Color ---

// ColorVal returns the hex color string from w:color/@w:val, or nil if not present.
func (rPr *CT_RPr) ColorVal() (*string, error) {
	color := rPr.Color()
	if color == nil {
		return nil, nil
	}
	val, err := color.Val()
	if err != nil {
		return nil, err
	}
	return &val, nil
}

// SetColorVal sets the color hex value. Passing nil removes the color element.
func (rPr *CT_RPr) SetColorVal(v *string) error {
	if v == nil {
		rPr.RemoveColor()
		return nil
	}
	color := rPr.GetOrAddColor()
	if err := color.SetVal(*v); err != nil {
		return err
	}
	return nil
}

// ColorTheme returns the theme color from w:color/@w:themeColor, or nil if not present.
func (rPr *CT_RPr) ColorTheme() (*enum.MsoThemeColorIndex, error) {
	color := rPr.Color()
	if color == nil {
		return nil, nil
	}
	tc := color.ThemeColor()
	if tc == "" {
		return nil, nil
	}
	v, err := enum.MsoThemeColorIndexFromXml(tc)
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetColorTheme sets the theme color. Passing nil removes the themeColor attribute.
func (rPr *CT_RPr) SetColorTheme(v *enum.MsoThemeColorIndex) error {
	if v == nil {
		color := rPr.Color()
		if color != nil {
			if err := color.SetThemeColor(""); err != nil {
				return err
			}
		}
		return nil
	}
	color := rPr.GetOrAddColor()
	xml, err := v.ToXml()
	if err != nil {
		return fmt.Errorf("CT_RPr.SetColorTheme: %w", err)
	}
	return color.SetThemeColor(xml)
}

// --- Size ---

// SzVal returns the font size from w:sz/@w:val as half-points, or nil if not present.
func (rPr *CT_RPr) SzVal() (*int64, error) {
	sz := rPr.Sz()
	if sz == nil {
		return nil, nil
	}
	val, err := sz.Val()
	if err != nil {
		return nil, err
	}
	return &val, nil
}

// SetSzVal sets the font size in half-points. Passing nil removes the sz element.
func (rPr *CT_RPr) SetSzVal(v *int64) error {
	if v == nil {
		rPr.RemoveSz()
		return nil
	}
	sz := rPr.GetOrAddSz()
	if err := sz.SetVal(*v); err != nil {
		return err
	}
	return nil
}

// --- Fonts ---

// RFontsAscii returns the ascii font name, or nil if not present.
func (rPr *CT_RPr) RFontsAscii() *string {
	rFonts := rPr.RFonts()
	if rFonts == nil {
		return nil
	}
	v := rFonts.Ascii()
	if v == "" {
		return nil
	}
	return &v
}

// SetRFontsAscii sets the ascii font name. Passing nil removes the rFonts element.
func (rPr *CT_RPr) SetRFontsAscii(v *string) error {
	if v == nil {
		rPr.RemoveRFonts()
		return nil
	}
	rFonts := rPr.GetOrAddRFonts()
	if err := rFonts.SetAscii(*v); err != nil {
		return err
	}
	return nil
}

// RFontsHAnsi returns the hAnsi font name, or nil if not present.
func (rPr *CT_RPr) RFontsHAnsi() *string {
	rFonts := rPr.RFonts()
	if rFonts == nil {
		return nil
	}
	v := rFonts.HAnsi()
	if v == "" {
		return nil
	}
	return &v
}

// SetRFontsHAnsi sets the hAnsi font name. Passing nil leaves rFonts alone.
func (rPr *CT_RPr) SetRFontsHAnsi(v *string) error {
	if v == nil && rPr.RFonts() == nil {
		return nil
	}
	rFonts := rPr.GetOrAddRFonts()
	if v == nil {
		if err := rFonts.SetHAnsi(""); err != nil {
			return err
		}
	} else {
		if err := rFonts.SetHAnsi(*v); err != nil {
			return err
		}
	}
	return nil
}

// --- Underline ---

// UVal returns the underline style from w:u/@w:val, or nil if not present.
func (rPr *CT_RPr) UVal() *string {
	u := rPr.U()
	if u == nil {
		return nil
	}
	v := u.Val()
	if v == "" {
		return nil
	}
	return &v
}

// SetUVal sets the underline style. Passing nil removes the u element.
func (rPr *CT_RPr) SetUVal(v *string) error {
	rPr.RemoveU()
	if v != nil {
		u := rPr.addU()
		if err := u.SetVal(*v); err != nil {
			return err
		}
	}
	return nil
}

// --- Highlight ---

// HighlightVal returns the highlight color string, or nil if not present.
func (rPr *CT_RPr) HighlightVal() (*string, error) {
	h := rPr.Highlight()
	if h == nil {
		return nil, nil
	}
	v, err := h.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetHighlightVal sets the highlight color. Passing nil removes the highlight element.
func (rPr *CT_RPr) SetHighlightVal(v *string) error {
	if v == nil {
		rPr.RemoveHighlight()
		return nil
	}
	h := rPr.GetOrAddHighlight()
	if err := h.SetVal(*v); err != nil {
		return err
	}
	return nil
}

// --- Style ---

// StyleVal returns the run style string from w:rStyle/@w:val, or nil if not present.
func (rPr *CT_RPr) StyleVal() (*string, error) {
	rStyle := rPr.RStyle()
	if rStyle == nil {
		return nil, nil
	}
	v, err := rStyle.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetStyleVal sets the run style. Passing nil removes the rStyle element.
func (rPr *CT_RPr) SetStyleVal(v *string) error {
	if v == nil {
		rPr.RemoveRStyle()
		return nil
	}
	rStyle := rPr.RStyle()
	if rStyle == nil {
		s := rPr.addRStyle()
		if err := s.SetVal(*v); err != nil {
			return err
		}
	} else {
		if err := rStyle.SetVal(*v); err != nil {
			return err
		}
	}
	return nil
}

// --- Subscript / Superscript ---

// Subscript returns true if vertAlign is "subscript", false if it's something else,
// nil if vertAlign is not present.
func (rPr *CT_RPr) Subscript() (*bool, error) {
	va := rPr.VertAlign()
	if va == nil {
		return nil, nil
	}
	v, err := va.Val()
	if err != nil {
		return nil, err
	}
	result := v == "subscript"
	return &result, nil
}

// SetSubscript sets the subscript state. nil removes vertAlign,
// true sets it to "subscript", false clears only if currently "subscript".
func (rPr *CT_RPr) SetSubscript(v *bool) error {
	if v == nil {
		rPr.RemoveVertAlign()
	} else if *v {
		if err := rPr.GetOrAddVertAlign().SetVal("subscript"); err != nil {
			return err
		}
	} else {
		va := rPr.VertAlign()
		if va != nil {
			val, err := va.Val()
			if err != nil {
				return fmt.Errorf("oxml: reading vertAlign val for subscript clear: %w", err)
			}
			if val == "subscript" {
				rPr.RemoveVertAlign()
			}
		}
	}
	return nil
}

// Superscript returns true if vertAlign is "superscript", false if it's something else,
// nil if vertAlign is not present.
func (rPr *CT_RPr) Superscript() (*bool, error) {
	va := rPr.VertAlign()
	if va == nil {
		return nil, nil
	}
	v, err := va.Val()
	if err != nil {
		return nil, err
	}
	result := v == "superscript"
	return &result, nil
}

// --- ComplexScript (w:cs) ---

// ComplexScriptVal returns the tri-state complex-script value.
func (rPr *CT_RPr) ComplexScriptVal() *bool {
	return rPr.getBoolVal("w:cs")
}

// SetComplexScriptVal sets the complex-script tri-state.
func (rPr *CT_RPr) SetComplexScriptVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddCs, rPr.RemoveCs)
}

// --- CsBold (w:bCs) ---

// CsBoldVal returns the tri-state complex-script bold value.
func (rPr *CT_RPr) CsBoldVal() *bool {
	return rPr.getBoolVal("w:bCs")
}

// SetCsBoldVal sets the complex-script bold tri-state.
func (rPr *CT_RPr) SetCsBoldVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddBCs, rPr.RemoveBCs)
}

// --- CsItalic (w:iCs) ---

// CsItalicVal returns the tri-state complex-script italic value.
func (rPr *CT_RPr) CsItalicVal() *bool {
	return rPr.getBoolVal("w:iCs")
}

// SetCsItalicVal sets the complex-script italic tri-state.
func (rPr *CT_RPr) SetCsItalicVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddICs, rPr.RemoveICs)
}

// --- Rtl (w:rtl) ---

// RtlVal returns the tri-state right-to-left value.
func (rPr *CT_RPr) RtlVal() *bool {
	return rPr.getBoolVal("w:rtl")
}

// SetRtlVal sets the right-to-left tri-state.
func (rPr *CT_RPr) SetRtlVal(v *bool) error {
	return rPr.setBoolValWith(v, rPr.GetOrAddRtl, rPr.RemoveRtl)
}

// SetSuperscript sets the superscript state.
func (rPr *CT_RPr) SetSuperscript(v *bool) error {
	if v == nil {
		rPr.RemoveVertAlign()
	} else if *v {
		if err := rPr.GetOrAddVertAlign().SetVal("superscript"); err != nil {
			return err
		}
	} else {
		va := rPr.VertAlign()
		if va != nil {
			val, err := va.Val()
			if err != nil {
				return fmt.Errorf("oxml: reading vertAlign val for superscript clear: %w", err)
			}
			if val == "superscript" {
				rPr.RemoveVertAlign()
			}
		}
	}
	return nil
}
