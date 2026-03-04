package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// rPrProvider is any XML element that contains a <w:rPr> child.
// Satisfied by *oxml.CT_R (run) and *oxml.CT_Style (style definition).
//
// This mirrors the Python duck-typing where Font.__init__ accepts CT_R
// but CharacterStyle.font passes CT_Style — both have .rPr and
// .get_or_add_rPr(). Follows the same pattern as pPrProvider in parfmt.go.
type rPrProvider interface {
	RPr() *oxml.CT_RPr
	GetOrAddRPr() *oxml.CT_RPr
}

// Font wraps an element containing <w:rPr>, providing access to character
// formatting properties such as font name, size, bold, and subscript.
//
// Mirrors Python Font(ElementProxy).
type Font struct {
	rPrOwner rPrProvider
}

// newFont creates a Font proxy from a run element.
func newFont(r *oxml.CT_R) *Font {
	return &Font{rPrOwner: r}
}

// newFontFromStyle creates a Font proxy from a style element.
// Used by BaseStyle.Font() — mirrors Python CharacterStyle.font which
// passes CT_Style to Font.__init__.
func newFontFromStyle(s *oxml.CT_Style) *Font {
	return &Font{rPrOwner: s}
}

// --------------------------------------------------------------------------
// Tri-state boolean properties — all 20 from Python Font
// --------------------------------------------------------------------------

// AllCaps returns the tri-state all-caps value.
func (f *Font) AllCaps() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.CapsVal() })
}

// SetAllCaps sets the tri-state all-caps value.
func (f *Font) SetAllCaps(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetCapsVal(v) })
}

// Bold returns the tri-state bold value.
func (f *Font) Bold() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.BoldVal() })
}

// SetBold sets the tri-state bold value.
func (f *Font) SetBold(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetBoldVal(v) })
}

// ComplexScript returns the tri-state complex-script value.
func (f *Font) ComplexScript() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.ComplexScriptVal() })
}

// SetComplexScript sets the tri-state complex-script value.
func (f *Font) SetComplexScript(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetComplexScriptVal(v) })
}

// CsBold returns the tri-state complex-script bold value.
func (f *Font) CsBold() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.CsBoldVal() })
}

// SetCsBold sets the tri-state complex-script bold value.
func (f *Font) SetCsBold(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetCsBoldVal(v) })
}

// CsItalic returns the tri-state complex-script italic value.
func (f *Font) CsItalic() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.CsItalicVal() })
}

// SetCsItalic sets the tri-state complex-script italic value.
func (f *Font) SetCsItalic(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetCsItalicVal(v) })
}

// DoubleStrike returns the tri-state double-strikethrough value.
func (f *Font) DoubleStrike() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.DstrikeVal() })
}

// SetDoubleStrike sets the tri-state double-strikethrough value.
func (f *Font) SetDoubleStrike(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetDstrikeVal(v) })
}

// Emboss returns the tri-state emboss value.
func (f *Font) Emboss() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.EmbossVal() })
}

// SetEmboss sets the tri-state emboss value.
func (f *Font) SetEmboss(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetEmbossVal(v) })
}

// Hidden returns the tri-state vanish/hidden value.
func (f *Font) Hidden() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.VanishVal() })
}

// SetHidden sets the tri-state hidden value.
func (f *Font) SetHidden(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetVanishVal(v) })
}

// Italic returns the tri-state italic value.
func (f *Font) Italic() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.ItalicVal() })
}

// SetItalic sets the tri-state italic value.
func (f *Font) SetItalic(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetItalicVal(v) })
}

// Imprint returns the tri-state imprint value.
func (f *Font) Imprint() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.ImprintVal() })
}

// SetImprint sets the tri-state imprint value.
func (f *Font) SetImprint(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetImprintVal(v) })
}

// Math returns the tri-state oMath value.
func (f *Font) Math() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.OMathVal() })
}

// SetMath sets the tri-state oMath value.
func (f *Font) SetMath(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetOMathVal(v) })
}

// NoProof returns the tri-state noProof value.
func (f *Font) NoProof() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.NoProofVal() })
}

// SetNoProof sets the tri-state noProof value.
func (f *Font) SetNoProof(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetNoProofVal(v) })
}

// Outline returns the tri-state outline value.
func (f *Font) Outline() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.OutlineVal() })
}

// SetOutline sets the tri-state outline value.
func (f *Font) SetOutline(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetOutlineVal(v) })
}

// Rtl returns the tri-state rtl value.
func (f *Font) Rtl() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.RtlVal() })
}

// SetRtl sets the tri-state rtl value.
func (f *Font) SetRtl(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetRtlVal(v) })
}

// Shadow returns the tri-state shadow value.
func (f *Font) Shadow() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.ShadowVal() })
}

// SetShadow sets the tri-state shadow value.
func (f *Font) SetShadow(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetShadowVal(v) })
}

// SmallCaps returns the tri-state small-caps value.
func (f *Font) SmallCaps() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.SmallCapsVal() })
}

// SetSmallCaps sets the tri-state small-caps value.
func (f *Font) SetSmallCaps(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetSmallCapsVal(v) })
}

// SnapToGrid returns the tri-state snapToGrid value.
func (f *Font) SnapToGrid() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.SnapToGridVal() })
}

// SetSnapToGrid sets the tri-state snapToGrid value.
func (f *Font) SetSnapToGrid(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetSnapToGridVal(v) })
}

// SpecVanish returns the tri-state specVanish value.
func (f *Font) SpecVanish() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.SpecVanishVal() })
}

// SetSpecVanish sets the tri-state specVanish value.
func (f *Font) SetSpecVanish(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetSpecVanishVal(v) })
}

// Strike returns the tri-state strikethrough value.
func (f *Font) Strike() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.StrikeVal() })
}

// SetStrike sets the tri-state strikethrough value.
func (f *Font) SetStrike(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetStrikeVal(v) })
}

// WebHidden returns the tri-state webHidden value.
func (f *Font) WebHidden() *bool {
	return f.getBoolProp(func(rPr *oxml.CT_RPr) *bool { return rPr.WebHiddenVal() })
}

// SetWebHidden sets the tri-state webHidden value.
func (f *Font) SetWebHidden(v *bool) error {
	return f.setBoolProp(func(rPr *oxml.CT_RPr) error { return rPr.SetWebHiddenVal(v) })
}

// --------------------------------------------------------------------------
// Non-boolean properties
// --------------------------------------------------------------------------

// Color returns the ColorFormat for this font.
func (f *Font) Color() *ColorFormat {
	return newColorFormat(f.rPrOwner)
}

// HighlightColor returns the highlight color index, or nil if not set.
func (f *Font) HighlightColor() (*enum.WdColorIndex, error) {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	val, err := rPr.HighlightVal()
	if err != nil {
		return nil, err
	}
	if val == nil {
		return nil, nil
	}
	ci, err := enum.WdColorIndexFromXml(*val)
	if err != nil {
		return nil, err
	}
	return &ci, nil
}

// SetHighlightColor sets the highlight color. Passing nil removes it.
func (f *Font) SetHighlightColor(v *enum.WdColorIndex) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	if v == nil {
		return rPr.SetHighlightVal(nil)
	}
	xml, err := v.ToXml()
	if err != nil {
		return fmt.Errorf("docx: invalid highlight color: %w", err)
	}
	return rPr.SetHighlightVal(&xml)
}

// Name returns the font name (ascii), or nil if not set.
func (f *Font) Name() *string {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil
	}
	return rPr.RFontsAscii()
}

// SetName sets the font name. Sets both ascii and hAnsi. Passing nil removes it.
func (f *Font) SetName(v *string) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	if err := rPr.SetRFontsAscii(v); err != nil {
		return err
	}
	return rPr.SetRFontsHAnsi(v)
}

// Size returns the font size as a Length (EMU), or nil if inherited.
//
// Mirrors Python Font.size (getter) — returns Length (EMU).
// Internally the XML stores half-points: <w:sz val="24"/> = 12pt = 152400 EMU.
// Python converts via Pt(int(str_value) / 2.0), i.e. float64 arithmetic.
func (f *Font) Size() (*Length, error) {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	hp, err := rPr.SzVal()
	if err != nil {
		return nil, fmt.Errorf("docx: reading font size: %w", err)
	}
	if hp == nil {
		return nil, nil
	}
	// half-points → EMU via float64, matching Python's Pt(int(val) / 2.0):
	//   Pt(x) = int(x * 12700)
	emu := Length(int64(float64(*hp) / 2.0 * float64(EmusPerPt)))
	return &emu, nil
}

// SetSize sets the font size. Passing nil removes it.
// Accepts Length (EMU). E.g. Pt(12) for 12-point font.
//
// Mirrors Python Font.size (setter) — accepts EMU.
// Python converts via int(emu.pt * 2), i.e. float64 arithmetic.
func (f *Font) SetSize(v *Length) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	if v == nil {
		return rPr.SetSzVal(nil)
	}
	// EMU → half-points via float64, matching Python's int(Emu(value).pt * 2):
	//   emu.pt = emu / 12700.0
	hp := int64(float64(*v) / float64(EmusPerPt) * 2.0)
	return rPr.SetSzVal(&hp)
}

// Subscript returns the tri-state subscript value.
func (f *Font) Subscript() (*bool, error) {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	return rPr.Subscript()
}

// SetSubscript sets the tri-state subscript value.
func (f *Font) SetSubscript(v *bool) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	return rPr.SetSubscript(v)
}

// Superscript returns the tri-state superscript value.
func (f *Font) Superscript() (*bool, error) {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	return rPr.Superscript()
}

// SetSuperscript sets the tri-state superscript value.
func (f *Font) SetSuperscript(v *bool) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	return rPr.SetSuperscript(v)
}

// Underline returns the underline setting, or nil if inherited.
//
// Mirrors Python Font.underline (getter). Python's OptionalAttribute
// descriptor raises on unrecognised w:val values; we propagate the error
// rather than silently returning nil (which means "inherited").
func (f *Font) Underline() (*UnderlineVal, error) {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil, nil
	}
	uVal := rPr.UVal()
	if uVal == nil {
		return nil, nil // inherited
	}
	switch *uVal {
	case "single":
		v := UnderlineSingle()
		return &v, nil
	case "none":
		v := UnderlineNone()
		return &v, nil
	default:
		wdu, err := enum.WdUnderlineFromXml(*uVal)
		if err != nil {
			return nil, fmt.Errorf("docx: parsing underline %q: %w", *uVal, err)
		}
		v := UnderlineStyle(wdu)
		return &v, nil
	}
}

// SetUnderline sets the underline value. Pass nil to inherit.
//
// Mirrors Python Font.underline (setter).
func (f *Font) SetUnderline(v *UnderlineVal) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	if v == nil {
		return rPr.SetUVal(nil)
	}
	switch {
	case v.IsSingle():
		s := "single"
		return rPr.SetUVal(&s)
	case v.IsNone():
		s := "none"
		return rPr.SetUVal(&s)
	case v.IsStyle():
		xml, err := v.Style().ToXml()
		if err != nil {
			return fmt.Errorf("docx: invalid underline: %w", err)
		}
		return rPr.SetUVal(&xml)
	default:
		return rPr.SetUVal(nil)
	}
}

// --------------------------------------------------------------------------
// Internal helpers
// --------------------------------------------------------------------------

func (f *Font) getBoolProp(getter func(*oxml.CT_RPr) *bool) *bool {
	rPr := f.rPrOwner.RPr()
	if rPr == nil {
		return nil
	}
	return getter(rPr)
}

func (f *Font) setBoolProp(setter func(*oxml.CT_RPr) error) error {
	rPr := f.rPrOwner.GetOrAddRPr()
	return setter(rPr)
}
