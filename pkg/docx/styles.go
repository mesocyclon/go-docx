package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// --------------------------------------------------------------------------
// BabelFish — style name translation (canonical source: oxml.BabelFish)
// --------------------------------------------------------------------------

// UI2Internal converts a UI style name to internal form.
// Delegates to the canonical oxml.UI2Internal.
func UI2Internal(name string) string { return oxml.UI2Internal(name) }

// Internal2UI converts an internal style name to UI form.
// Delegates to the canonical oxml.Internal2UI.
func Internal2UI(name string) string { return oxml.Internal2UI(name) }

// --------------------------------------------------------------------------
// styleFactory
// --------------------------------------------------------------------------

// styleFactory creates the appropriate BaseStyle subtype for a CT_Style element.
//
// Mirrors Python StyleFactory.
func styleFactory(styleElm *oxml.CT_Style) *BaseStyle {
	return &BaseStyle{element: styleElm}
}

// --------------------------------------------------------------------------
// Styles
// --------------------------------------------------------------------------

// Styles wraps CT_Styles, providing access to the styles in a document.
//
// Mirrors Python Styles(ElementProxy).
type Styles struct {
	element *oxml.CT_Styles
}

// newStyles creates a new Styles proxy.
func newStyles(element *oxml.CT_Styles) *Styles {
	return &Styles{element: element}
}

// Contains returns true if a style with the given UI name exists.
func (s *Styles) Contains(name string) bool {
	internalName := UI2Internal(name)
	for _, st := range s.element.StyleList() {
		nm, err := st.NameVal()
		if err == nil && nm == internalName {
			return true
		}
	}
	return false
}

// Get returns the style with the given UI name.
func (s *Styles) Get(name string) (*BaseStyle, error) {
	internalName := UI2Internal(name)
	st := s.element.GetByName(internalName)
	if st != nil {
		return styleFactory(st), nil
	}
	// Fallback: try by ID (deprecated)
	st = s.element.GetByID(name)
	if st != nil {
		return styleFactory(st), nil
	}
	return nil, fmt.Errorf("docx: no style with name %q", name)
}

// Iter returns all styles.
func (s *Styles) Iter() []*BaseStyle {
	lst := s.element.StyleList()
	result := make([]*BaseStyle, len(lst))
	for i, st := range lst {
		result[i] = styleFactory(st)
	}
	return result
}

// Len returns the number of styles.
func (s *Styles) Len() int {
	return len(s.element.StyleList())
}

// AddStyle adds a new style with the given name and type.
//
// Mirrors Python Styles.add_style.
func (s *Styles) AddStyle(name string, styleType enum.WdStyleType, builtin bool) (*BaseStyle, error) {
	styleName := UI2Internal(name)
	if s.Contains(name) {
		return nil, fmt.Errorf("docx: document already contains style %q", name)
	}
	st, err := s.element.AddStyleOfType(styleName, styleType, builtin)
	if err != nil {
		return nil, err
	}
	return styleFactory(st), nil
}

// Default returns the default style for the given type, or nil.
//
// Mirrors Python Styles.default.
func (s *Styles) Default(styleType enum.WdStyleType) (*BaseStyle, error) {
	st, err := s.element.DefaultFor(styleType)
	if err != nil {
		return nil, fmt.Errorf("docx: getting default style: %w", err)
	}
	if st == nil {
		return nil, nil
	}
	return styleFactory(st), nil
}

// GetByID returns the style matching styleID and styleType. Returns the default
// style if styleID is nil or not found.
//
// Mirrors Python Styles.get_by_id.
func (s *Styles) GetByID(styleID *string, styleType enum.WdStyleType) (*BaseStyle, error) {
	if styleID == nil {
		return s.Default(styleType)
	}
	st := s.element.GetByID(*styleID)
	if st == nil {
		return s.Default(styleType)
	}
	stTypeXml, err := styleType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("docx: invalid style type: %w", err)
	}
	if st.Type() != stTypeXml {
		return s.Default(styleType)
	}
	return styleFactory(st), nil
}

// GetStyleID returns the style ID for the given style or name.
//
// Mirrors Python Styles.get_style_id.
// GetStyleID returns the style ID for the given style or name.
// styleOrName must be a StyleName, *BaseStyle, or nil.
//
// Mirrors Python Styles.get_style_id.
func (s *Styles) GetStyleID(styleOrName StyleRef, styleType enum.WdStyleType) (*string, error) {
	if styleOrName == nil {
		return nil, nil
	}
	switch v := styleOrName.(type) {
	case *BaseStyle:
		return s.getStyleIDFromStyle(v, styleType)
	case StyleName:
		return s.getStyleIDFromName(string(v), styleType)
	default:
		return nil, fmt.Errorf("docx: unsupported style type %T", styleOrName)
	}
}

// LatentStyles returns the LatentStyles object.
func (s *Styles) LatentStyles() *LatentStyles {
	ls := s.element.GetOrAddLatentStyles()
	return &LatentStyles{element: ls}
}

func (s *Styles) getStyleIDFromName(name string, styleType enum.WdStyleType) (*string, error) {
	return s.element.GetStyleIDByName(name, styleType)
}

func (s *Styles) getStyleIDFromStyle(style *BaseStyle, styleType enum.WdStyleType) (*string, error) {
	st, err := style.Type()
	if err != nil {
		return nil, err
	}
	if st != styleType {
		return nil, fmt.Errorf("docx: assigned style is type %v, need type %v", st, styleType)
	}
	def, err := s.Default(styleType)
	if err != nil {
		return nil, fmt.Errorf("docx: getting default style: %w", err)
	}
	if def != nil && style.StyleID() == def.StyleID() {
		return nil, nil
	}
	id := style.StyleID()
	return &id, nil
}

// --------------------------------------------------------------------------
// BaseStyle
// --------------------------------------------------------------------------

// BaseStyle is the base for all style objects.
//
// Mirrors Python BaseStyle(ElementProxy).
type BaseStyle struct {
	element *oxml.CT_Style
}

// Builtin returns true if this is a built-in style.
func (s *BaseStyle) Builtin() bool {
	return s.element.IsBuiltin()
}

// Delete removes this style definition from the document.
func (s *BaseStyle) Delete() {
	s.element.Delete()
}

// Hidden returns true if this style is semi-hidden.
func (s *BaseStyle) Hidden() bool {
	return s.element.SemiHiddenVal()
}

// SetHidden sets the semi-hidden value.
func (s *BaseStyle) SetHidden(v bool) error {
	return s.element.SetSemiHiddenVal(v)
}

// Locked returns true if this style is locked.
func (s *BaseStyle) Locked() bool {
	return s.element.LockedVal()
}

// SetLocked sets the locked value.
func (s *BaseStyle) SetLocked(v bool) error {
	return s.element.SetLockedVal(v)
}

// Name returns the UI name of this style.
func (s *BaseStyle) Name() (string, error) {
	name, err := s.element.NameVal()
	if err != nil {
		return "", fmt.Errorf("docx: reading style name: %w", err)
	}
	return Internal2UI(name), nil
}

// SetName sets the style name.
func (s *BaseStyle) SetName(v string) error {
	return s.element.SetNameVal(v)
}

// Priority returns the sort priority, or nil if not set.
func (s *BaseStyle) Priority() (*int, error) {
	return s.element.UiPriorityVal()
}

// SetPriority sets the sort priority.
func (s *BaseStyle) SetPriority(v *int) error {
	return s.element.SetUiPriorityVal(v)
}

// QuickStyle returns true if this style should appear in the style gallery.
func (s *BaseStyle) QuickStyle() bool {
	return s.element.QFormatVal()
}

// SetQuickStyle sets the quick-style flag.
func (s *BaseStyle) SetQuickStyle(v bool) {
	s.element.SetQFormatVal(v)
}

// StyleID returns the unique key name for this style.
func (s *BaseStyle) StyleID() string {
	return s.element.StyleId()
}

// SetStyleID sets the style ID.
func (s *BaseStyle) SetStyleID(v string) error {
	return s.element.SetStyleId(v)
}

// Type returns the style type as WdStyleType.
//
// Mirrors Python BaseStyle.type: absent w:type defaults to PARAGRAPH
// (per OOXML spec); an unrecognised value is an error.
func (s *BaseStyle) Type() (enum.WdStyleType, error) {
	xmlType := s.element.Type()
	if xmlType == "" {
		return enum.WdStyleTypeParagraph, nil // absent → paragraph (matches Python)
	}
	st, err := enum.WdStyleTypeFromXml(xmlType)
	if err != nil {
		return 0, fmt.Errorf("docx: parsing style type %q: %w", xmlType, err)
	}
	return st, nil
}

// isStyleRef implements StyleRef, allowing *BaseStyle to be passed directly
// as a style argument to AddParagraph, SetStyle, etc.
func (*BaseStyle) isStyleRef() {}

// UnhideWhenUsed returns true if this style should be unhidden when applied.
func (s *BaseStyle) UnhideWhenUsed() bool {
	return s.element.UnhideWhenUsedVal()
}

// SetUnhideWhenUsed sets the unhide-when-used flag.
func (s *BaseStyle) SetUnhideWhenUsed(v bool) error {
	return s.element.SetUnhideWhenUsedVal(v)
}

// BaseStyleObj returns the style this one inherits from, or nil.
func (s *BaseStyle) BaseStyleObj() *BaseStyle {
	base := s.element.BaseStyle()
	if base == nil {
		return nil
	}
	return styleFactory(base)
}

// SetBaseStyle sets the base style. Passing nil removes the basedOn.
func (s *BaseStyle) SetBaseStyle(style *BaseStyle) error {
	if style == nil {
		return s.element.SetBasedOnVal("")
	}
	return s.element.SetBasedOnVal(style.StyleID())
}

// Font returns the Font providing access to character formatting for this style.
//
// Mirrors Python CharacterStyle.font → Font(self._element).
func (s *BaseStyle) Font() *Font {
	return newFontFromStyle(s.element)
}

// ParagraphFormat returns the ParagraphFormat for this style.
func (s *BaseStyle) ParagraphFormat() *ParagraphFormat {
	return newParagraphFormatFromStyle(s.element)
}

// NextParagraphStyle returns the style applied to the next paragraph.
// Returns self if none defined.
func (s *BaseStyle) NextParagraphStyle() *BaseStyle {
	next := s.element.NextStyle()
	if next == nil {
		return s
	}
	if next.Type() != s.element.Type() {
		return s
	}
	return styleFactory(next)
}

// CT_Style returns the underlying oxml element.
func (s *BaseStyle) CT_Style() *oxml.CT_Style { return s.element }

// --------------------------------------------------------------------------
// LatentStyles
// --------------------------------------------------------------------------

// LatentStyles provides access to default behaviors for latent styles.
//
// Mirrors Python LatentStyles(ElementProxy).
type LatentStyles struct {
	element *oxml.CT_LatentStyles
}

// Get returns the latent style with the given UI name.
func (ls *LatentStyles) Get(name string) (*LatentStyle, error) {
	internalName := UI2Internal(name)
	exc := ls.element.GetByName(internalName)
	if exc == nil {
		return nil, fmt.Errorf("docx: no latent style with name %q", name)
	}
	return &LatentStyle{element: exc}, nil
}

// Iter returns all latent style exceptions.
func (ls *LatentStyles) Iter() []*LatentStyle {
	lst := ls.element.LsdExceptionList()
	result := make([]*LatentStyle, len(lst))
	for i, exc := range lst {
		result[i] = &LatentStyle{element: exc}
	}
	return result
}

// Len returns the number of latent style exceptions.
func (ls *LatentStyles) Len() int {
	return len(ls.element.LsdExceptionList())
}

// AddLatentStyle adds a new latent style override.
func (ls *LatentStyles) AddLatentStyle(name string) *LatentStyle {
	exc := ls.element.AddLsdException()
	exc.RawElement().CreateAttr("w:name", UI2Internal(name))
	return &LatentStyle{element: exc}
}

// DefaultPriority returns the default UI priority, or nil.
func (ls *LatentStyles) DefaultPriority() (*int, error) {
	return ls.element.DefUIPriority()
}

// SetDefaultPriority sets the default UI priority. Passing nil removes it.
func (ls *LatentStyles) SetDefaultPriority(v *int) error {
	return ls.element.SetDefUIPriority(v)
}

// DefaultToHidden returns whether latent styles are hidden by default.
func (ls *LatentStyles) DefaultToHidden() bool {
	return ls.element.BoolProp("defSemiHidden")
}

// SetDefaultToHidden sets the default hidden behavior.
func (ls *LatentStyles) SetDefaultToHidden(v bool) error {
	return ls.element.SetBoolProp("defSemiHidden", v)
}

// DefaultToLocked returns whether latent styles are locked by default.
func (ls *LatentStyles) DefaultToLocked() bool {
	return ls.element.BoolProp("defLockedState")
}

// SetDefaultToLocked sets the default locked behavior.
func (ls *LatentStyles) SetDefaultToLocked(v bool) error {
	return ls.element.SetBoolProp("defLockedState", v)
}

// DefaultToQuickStyle returns whether latent styles appear in gallery by default.
func (ls *LatentStyles) DefaultToQuickStyle() bool {
	return ls.element.BoolProp("defQFormat")
}

// SetDefaultToQuickStyle sets the default quick-style behavior.
func (ls *LatentStyles) SetDefaultToQuickStyle(v bool) error {
	return ls.element.SetBoolProp("defQFormat", v)
}

// DefaultToUnhideWhenUsed returns whether latent styles unhide by default when used.
func (ls *LatentStyles) DefaultToUnhideWhenUsed() bool {
	return ls.element.BoolProp("defUnhideWhenUsed")
}

// SetDefaultToUnhideWhenUsed sets the default unhide-when-used behavior.
func (ls *LatentStyles) SetDefaultToUnhideWhenUsed(v bool) error {
	return ls.element.SetBoolProp("defUnhideWhenUsed", v)
}

// LoadCount returns the number of built-in styles to initialize, or nil.
func (ls *LatentStyles) LoadCount() (*int, error) {
	return ls.element.Count()
}

// SetLoadCount sets the load count.
func (ls *LatentStyles) SetLoadCount(v *int) error {
	return ls.element.SetCount(v)
}

// --------------------------------------------------------------------------
// LatentStyle
// --------------------------------------------------------------------------

// LatentStyle is a proxy for a w:lsdException element.
//
// Mirrors Python _LatentStyle(ElementProxy).
type LatentStyle struct {
	element *oxml.CT_LsdException
}

// Delete removes this latent style exception.
func (ls *LatentStyle) Delete() {
	ls.element.Delete()
}

// Hidden returns the tri-state hidden value.
func (ls *LatentStyle) Hidden() *bool {
	return ls.element.OnOffProp("w:semiHidden")
}

// SetHidden sets the hidden value.
func (ls *LatentStyle) SetHidden(v *bool) error {
	return ls.element.SetOnOffProp("w:semiHidden", v)
}

// Locked returns the tri-state locked value.
func (ls *LatentStyle) Locked() *bool {
	return ls.element.OnOffProp("w:locked")
}

// SetLocked sets the locked value.
func (ls *LatentStyle) SetLocked(v *bool) error {
	return ls.element.SetOnOffProp("w:locked", v)
}

// Name returns the style name.
func (ls *LatentStyle) Name() string {
	name := ls.element.RawElement().SelectAttrValue("w:name", "")
	return Internal2UI(name)
}

// Priority returns the sort priority, or nil.
func (ls *LatentStyle) Priority() (*int, error) {
	return ls.element.UiPriority()
}

// SetPriority sets the sort priority.
func (ls *LatentStyle) SetPriority(v *int) error {
	return ls.element.SetUiPriority(v)
}

// QuickStyle returns the tri-state quick-style value.
func (ls *LatentStyle) QuickStyle() *bool {
	return ls.element.OnOffProp("w:qFormat")
}

// SetQuickStyle sets the quick-style value.
func (ls *LatentStyle) SetQuickStyle(v *bool) error {
	return ls.element.SetOnOffProp("w:qFormat", v)
}

// UnhideWhenUsed returns the tri-state unhide-when-used value.
func (ls *LatentStyle) UnhideWhenUsed() *bool {
	return ls.element.OnOffProp("w:unhideWhenUsed")
}

// SetUnhideWhenUsed sets the unhide-when-used value.
func (ls *LatentStyle) SetUnhideWhenUsed(v *bool) error {
	return ls.element.SetOnOffProp("w:unhideWhenUsed", v)
}
