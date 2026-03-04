package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// pPrProvider is implemented by CT_P and CT_Style — both have PPr/GetOrAddPPr
// with correct XSD element ordering via generated insert methods.
type pPrProvider interface {
	PPr() *oxml.CT_PPr
	GetOrAddPPr() *oxml.CT_PPr
}

// ParagraphFormat provides access to paragraph formatting such as justification,
// indentation, line spacing, space before and after, and widow/orphan control.
//
// Mirrors Python ParagraphFormat(ElementProxy).
type ParagraphFormat struct {
	provider pPrProvider
}

// newParagraphFormatFromP creates a ParagraphFormat wrapping a CT_P element.
func newParagraphFormatFromP(p *oxml.CT_P) *ParagraphFormat {
	return &ParagraphFormat{provider: p}
}

// newParagraphFormatFromStyle creates a ParagraphFormat wrapping a CT_Style element.
func newParagraphFormatFromStyle(s *oxml.CT_Style) *ParagraphFormat {
	return &ParagraphFormat{provider: s}
}

// Alignment returns the paragraph alignment, or nil if inherited.
func (pf *ParagraphFormat) Alignment() (*enum.WdParagraphAlignment, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.JcVal()
}

// SetAlignment sets the paragraph alignment. Passing nil removes the setting.
func (pf *ParagraphFormat) SetAlignment(v *enum.WdParagraphAlignment) error {
	return pf.provider.GetOrAddPPr().SetJcVal(v)
}

// FirstLineIndent returns the first-line indent in twips, or nil if inherited.
//
// Note: the OXML layer stores and returns raw twips (twentieths of a point).
// Python returns EMU (Length) because its OXML layer auto-converts.
// To convert to EMU: emu = twips * 635.
func (pf *ParagraphFormat) FirstLineIndent() (*int, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.FirstLineIndent()
}

// SetFirstLineIndent sets the first-line indent in twips. Passing nil removes it.
func (pf *ParagraphFormat) SetFirstLineIndent(v *int) error {
	return pf.provider.GetOrAddPPr().SetFirstLineIndent(v)
}

// KeepTogether returns the tri-state keep-together value, or nil if inherited.
func (pf *ParagraphFormat) KeepTogether() *bool {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil
	}
	return pPr.KeepLinesVal()
}

// SetKeepTogether sets the keep-together value.
func (pf *ParagraphFormat) SetKeepTogether(v *bool) error {
	return pf.provider.GetOrAddPPr().SetKeepLinesVal(v)
}

// KeepWithNext returns the tri-state keep-with-next value, or nil if inherited.
func (pf *ParagraphFormat) KeepWithNext() *bool {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil
	}
	return pPr.KeepNextVal()
}

// SetKeepWithNext sets the keep-with-next value.
func (pf *ParagraphFormat) SetKeepWithNext(v *bool) error {
	return pf.provider.GetOrAddPPr().SetKeepNextVal(v)
}

// LeftIndent returns the left indent in twips, or nil if inherited.
func (pf *ParagraphFormat) LeftIndent() (*int, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.IndLeft()
}

// SetLeftIndent sets the left indent in twips. Passing nil removes it.
func (pf *ParagraphFormat) SetLeftIndent(v *int) error {
	return pf.provider.GetOrAddPPr().SetIndLeft(v)
}

// RightIndent returns the right indent in twips, or nil if inherited.
func (pf *ParagraphFormat) RightIndent() (*int, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.IndRight()
}

// SetRightIndent sets the right indent in twips. Passing nil removes it.
func (pf *ParagraphFormat) SetRightIndent(v *int) error {
	return pf.provider.GetOrAddPPr().SetIndRight(v)
}

// SpaceAfter returns the space after in twips, or nil if inherited.
func (pf *ParagraphFormat) SpaceAfter() (*int, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.SpacingAfter()
}

// SetSpaceAfter sets the space after in twips. Passing nil removes it.
func (pf *ParagraphFormat) SetSpaceAfter(v *int) error {
	return pf.provider.GetOrAddPPr().SetSpacingAfter(v)
}

// SpaceBefore returns the space before in twips, or nil if inherited.
func (pf *ParagraphFormat) SpaceBefore() (*int, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	return pPr.SpacingBefore()
}

// SetSpaceBefore sets the space before in twips. Passing nil removes it.
func (pf *ParagraphFormat) SetSpaceBefore(v *int) error {
	return pf.provider.GetOrAddPPr().SetSpacingBefore(v)
}

// LineSpacing returns the line spacing value.
// Returns a float64 (multiple of line height) when rule is MULTIPLE,
// or an int (raw twips) for absolute height. Returns nil if inherited.
//
// Mirrors Python ParagraphFormat.line_spacing.
func (pf *ParagraphFormat) LineSpacing() (*LineSpacingVal, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	line, err := pPr.SpacingLine()
	if err != nil {
		return nil, err
	}
	if line == nil {
		return nil, nil
	}
	rule, err := pPr.SpacingLineRule()
	if err != nil {
		return nil, err
	}
	return toLineSpacingVal(*line, rule), nil
}

// SetLineSpacing sets the line spacing. Pass nil to inherit, a float64 for
// multiple (e.g. 2.0), or an int for absolute twips.
//
// Mirrors Python ParagraphFormat.line_spacing setter.
func (pf *ParagraphFormat) SetLineSpacing(v *LineSpacingVal) error {
	pPr := pf.provider.GetOrAddPPr()
	if v == nil {
		if err := pPr.SetSpacingLine(nil); err != nil {
			return err
		}
		return pPr.SetSpacingLineRule(nil) // remove
	}
	if v.IsMultiple() {
		// Multiple: val * 240 twips.
		twips := int(v.Multiple() * 240)
		if err := pPr.SetSpacingLine(&twips); err != nil {
			return err
		}
		return pPr.SetSpacingLineRule(wdlsPtr(enum.WdLineSpacingMultiple))
	}
	// Absolute twips.
	tw := v.Twips()
	if err := pPr.SetSpacingLine(&tw); err != nil {
		return err
	}
	rule, err := pPr.SpacingLineRule()
	if err != nil {
		return fmt.Errorf("docx: reading line spacing rule: %w", err)
	}
	if rule != nil && *rule == enum.WdLineSpacingAtLeast {
		return nil // preserve AT_LEAST
	}
	return pPr.SetSpacingLineRule(wdlsPtr(enum.WdLineSpacingExactly))
}

// LineSpacingRule returns the line spacing rule, or nil if inherited.
//
// Mirrors Python ParagraphFormat.line_spacing_rule.
func (pf *ParagraphFormat) LineSpacingRule() (*enum.WdLineSpacing, error) {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil, nil
	}
	line, err := pPr.SpacingLine()
	if err != nil {
		return nil, err
	}
	rule, err := pPr.SpacingLineRule()
	if err != nil {
		return nil, err
	}
	return lineSpacingRule(line, rule), nil
}

// SetLineSpacingRule sets the line spacing rule.
//
// Mirrors Python ParagraphFormat.line_spacing_rule setter.
func (pf *ParagraphFormat) SetLineSpacingRule(v enum.WdLineSpacing) error {
	pPr := pf.provider.GetOrAddPPr()
	switch v {
	case enum.WdLineSpacingSingle:
		twips := 240
		if err := pPr.SetSpacingLine(&twips); err != nil {
			return err
		}
		return pPr.SetSpacingLineRule(wdlsPtr(enum.WdLineSpacingMultiple))
	case enum.WdLineSpacingOnePointFive:
		twips := 360
		if err := pPr.SetSpacingLine(&twips); err != nil {
			return err
		}
		return pPr.SetSpacingLineRule(wdlsPtr(enum.WdLineSpacingMultiple))
	case enum.WdLineSpacingDouble:
		twips := 480
		if err := pPr.SetSpacingLine(&twips); err != nil {
			return err
		}
		return pPr.SetSpacingLineRule(wdlsPtr(enum.WdLineSpacingMultiple))
	default:
		return pPr.SetSpacingLineRule(&v)
	}
}

// PageBreakBefore returns the tri-state page-break-before value, or nil if inherited.
func (pf *ParagraphFormat) PageBreakBefore() *bool {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil
	}
	return pPr.PageBreakBeforeVal()
}

// SetPageBreakBefore sets the page-break-before value.
func (pf *ParagraphFormat) SetPageBreakBefore(v *bool) error {
	return pf.provider.GetOrAddPPr().SetPageBreakBeforeVal(v)
}

// WidowControl returns the tri-state widow-control value, or nil if inherited.
func (pf *ParagraphFormat) WidowControl() *bool {
	pPr := pf.provider.PPr()
	if pPr == nil {
		return nil
	}
	return pPr.WidowControlVal()
}

// SetWidowControl sets the widow-control value.
func (pf *ParagraphFormat) SetWidowControl(v *bool) error {
	return pf.provider.GetOrAddPPr().SetWidowControlVal(v)
}

// TabStops returns the TabStops providing access to tab stop definitions.
//
// Mirrors Python ParagraphFormat.tab_stops (lazyproperty).
func (pf *ParagraphFormat) TabStops() *TabStops {
	pPr := pf.provider.GetOrAddPPr()
	return newTabStops(pPr)
}

// --- internal helpers ---

// lineSpacing mirrors Python ParagraphFormat._line_spacing static method.
//
// spacingLine is raw twips from OXML. Python works with EMU but the ratio
// is identical: 240twips / 240twips == 152400EMU / 152400EMU == 1.0.
// toLineSpacingVal converts raw OXML values to a typed LineSpacingVal.
func toLineSpacingVal(spacingLine int, spacingLineRule *enum.WdLineSpacing) *LineSpacingVal {
	if spacingLineRule != nil && *spacingLineRule == enum.WdLineSpacingMultiple {
		v := LineSpacingMultiple(float64(spacingLine) / 240.0)
		return &v
	}
	v := LineSpacingTwips(spacingLine)
	return &v
}

// lineSpacingRule mirrors Python ParagraphFormat._line_spacing_rule static method.
//
// Compares raw twips (Python compares EMU equivalents: Twips(240)=152400, etc.).
func lineSpacingRule(line *int, lineRule *enum.WdLineSpacing) *enum.WdLineSpacing {
	if lineRule != nil && *lineRule == enum.WdLineSpacingMultiple && line != nil {
		switch *line {
		case 240:
			v := enum.WdLineSpacingSingle
			return &v
		case 360:
			v := enum.WdLineSpacingOnePointFive
			return &v
		case 480:
			v := enum.WdLineSpacingDouble
			return &v
		}
	}
	return lineRule
}

// wdlsPtr returns a pointer to the given WdLineSpacing value.
func wdlsPtr(v enum.WdLineSpacing) *enum.WdLineSpacing { return &v }
