package oxml

import (
	"fmt"
	"strings"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ===========================================================================
// styleIdFromName utility
// ===========================================================================

// StyleIdFromName returns the style ID corresponding to a style name.
// Special-case names like "Heading 1" map to "Heading1", etc.
// Default behaviour: remove spaces.
func StyleIdFromName(name string) string {
	special := map[string]string{
		"caption":   "Caption",
		"heading 1": "Heading1",
		"heading 2": "Heading2",
		"heading 3": "Heading3",
		"heading 4": "Heading4",
		"heading 5": "Heading5",
		"heading 6": "Heading6",
		"heading 7": "Heading7",
		"heading 8": "Heading8",
		"heading 9": "Heading9",
	}
	lower := strings.ToLower(name)
	if v, ok := special[lower]; ok {
		return v
	}
	return strings.ReplaceAll(name, " ", "")
}

// ===========================================================================
// CT_Styles — custom methods
// ===========================================================================

// GetByID returns the w:style element whose @w:styleId matches styleID, or nil.
func (ss *CT_Styles) GetByID(styleID string) *CT_Style {
	for _, s := range ss.StyleList() {
		if s.StyleId() == styleID {
			return s
		}
	}
	return nil
}

// GetByName returns the w:style element whose w:name/@w:val matches name, or nil.
func (ss *CT_Styles) GetByName(name string) *CT_Style {
	for _, s := range ss.StyleList() {
		n, err := s.NameVal()
		if err == nil && n == name {
			return s
		}
	}
	return nil
}

// DefaultFor returns the default style for the given type, or nil.
// If multiple defaults exist, returns the last one (per OOXML spec).
func (ss *CT_Styles) DefaultFor(styleType enum.WdStyleType) (*CT_Style, error) {
	xmlType, err := styleType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("oxml: invalid style type for DefaultFor: %w", err)
	}
	var last *CT_Style
	for _, s := range ss.StyleList() {
		if s.Type() == xmlType && s.Default() {
			last = s
		}
	}
	return last, nil
}

// AddStyleOfType creates and adds a new w:style element with the given name, type,
// and builtin flag. Returns the new style element.
func (ss *CT_Styles) AddStyleOfType(name string, styleType enum.WdStyleType, builtin bool) (*CT_Style, error) {
	xmlType, err := styleType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("oxml: invalid style type: %w", err)
	}
	style := ss.AddStyle()
	if err := style.SetType(xmlType); err != nil {
		return nil, err
	}
	if !builtin {
		if err := style.SetCustomStyle(true); err != nil {
			return nil, err
		}
	}
	if err := style.SetStyleId(StyleIdFromName(name)); err != nil {
		return nil, err
	}
	if err := style.SetNameVal(name); err != nil {
		return nil, err
	}
	return style, nil
}

// ===========================================================================
// CT_Style — custom methods
// ===========================================================================

// NameVal returns the value of w:name/@w:val, or "" if not present.
func (s *CT_Style) NameVal() (string, error) {
	n := s.Name()
	if n == nil {
		return "", nil
	}
	return n.Val()
}

// SetNameVal sets the w:name/@w:val. Passing "" removes the name element.
func (s *CT_Style) SetNameVal(name string) error {
	s.RemoveName()
	if name == "" {
		return nil
	}
	if err := s.GetOrAddName().SetVal(name); err != nil {
		return err
	}
	return nil
}

// BasedOnVal returns the value of w:basedOn/@w:val, or "" if not present.
func (s *CT_Style) BasedOnVal() (string, error) {
	b := s.BasedOn()
	if b == nil {
		return "", nil
	}
	return b.Val()
}

// SetBasedOnVal sets the basedOn value. Passing "" removes the element.
func (s *CT_Style) SetBasedOnVal(v string) error {
	s.RemoveBasedOn()
	if v == "" {
		return nil
	}
	if err := s.GetOrAddBasedOn().SetVal(v); err != nil {
		return err
	}
	return nil
}

// NextVal returns the value of w:next/@w:val, or "" if not present.
func (s *CT_Style) NextVal() (string, error) {
	n := s.Next()
	if n == nil {
		return "", nil
	}
	return n.Val()
}

// SetNextVal sets the next style ID. Passing "" removes the element.
func (s *CT_Style) SetNextVal(v string) error {
	s.RemoveNext()
	if v == "" {
		return nil
	}
	if err := s.GetOrAddNext().SetVal(v); err != nil {
		return err
	}
	return nil
}

// LockedVal returns the value of w:locked, or false if not present.
func (s *CT_Style) LockedVal() bool {
	l := s.Locked()
	if l == nil {
		return false
	}
	return l.Val()
}

// SetLockedVal sets the locked flag. Passing false removes the element.
func (s *CT_Style) SetLockedVal(v bool) error {
	s.RemoveLocked()
	if v {
		if err := s.GetOrAddLocked().SetVal(true); err != nil {
			return err
		}
	}
	return nil
}

// SemiHiddenVal returns the value of w:semiHidden, or false if not present.
func (s *CT_Style) SemiHiddenVal() bool {
	sh := s.SemiHidden()
	if sh == nil {
		return false
	}
	return sh.Val()
}

// SetSemiHiddenVal sets the semiHidden flag.
func (s *CT_Style) SetSemiHiddenVal(v bool) error {
	s.RemoveSemiHidden()
	if v {
		if err := s.GetOrAddSemiHidden().SetVal(true); err != nil {
			return err
		}
	}
	return nil
}

// UnhideWhenUsedVal returns the value of w:unhideWhenUsed, or false.
func (s *CT_Style) UnhideWhenUsedVal() bool {
	u := s.UnhideWhenUsed()
	if u == nil {
		return false
	}
	return u.Val()
}

// SetUnhideWhenUsedVal sets the unhideWhenUsed flag.
func (s *CT_Style) SetUnhideWhenUsedVal(v bool) error {
	s.RemoveUnhideWhenUsed()
	if v {
		if err := s.GetOrAddUnhideWhenUsed().SetVal(true); err != nil {
			return err
		}
	}
	return nil
}

// QFormatVal returns the value of w:qFormat, or false.
func (s *CT_Style) QFormatVal() bool {
	q := s.QFormat()
	if q == nil {
		return false
	}
	return q.Val()
}

// SetQFormatVal sets the qFormat flag.
func (s *CT_Style) SetQFormatVal(v bool) {
	s.RemoveQFormat()
	if v {
		s.GetOrAddQFormat()
	}
}

// UiPriorityVal returns the value of w:uiPriority/@w:val, or nil.
func (s *CT_Style) UiPriorityVal() (*int, error) {
	u := s.UiPriority()
	if u == nil {
		return nil, nil
	}
	v, err := u.Val()
	if err != nil {
		return nil, err
	}
	return &v, nil
}

// SetUiPriorityVal sets the uiPriority. Passing nil removes the element.
func (s *CT_Style) SetUiPriorityVal(v *int) error {
	s.RemoveUiPriority()
	if v == nil {
		return nil
	}
	if err := s.GetOrAddUiPriority().SetVal(*v); err != nil {
		return err
	}
	return nil
}

// BaseStyle returns the sibling CT_Style that this style is based on, or nil.
func (s *CT_Style) BaseStyle() *CT_Style {
	basedOn, err := s.BasedOnVal()
	if err != nil || basedOn == "" {
		return nil
	}
	parent := s.e.Parent()
	if parent == nil {
		return nil
	}
	styles := &CT_Styles{Element{e: parent}}
	return styles.GetByID(basedOn)
}

// NextStyle returns the sibling CT_Style identified by w:next, or nil.
func (s *CT_Style) NextStyle() *CT_Style {
	nextVal, err := s.NextVal()
	if err != nil || nextVal == "" {
		return nil
	}
	parent := s.e.Parent()
	if parent == nil {
		return nil
	}
	styles := &CT_Styles{Element{e: parent}}
	return styles.GetByID(nextVal)
}

// Delete removes this w:style element from its parent w:styles.
func (s *CT_Style) Delete() {
	parent := s.e.Parent()
	if parent != nil {
		parent.RemoveChild(s.e)
	}
}

// IsBuiltin returns true if this is a built-in style (customStyle is not set or false).
func (s *CT_Style) IsBuiltin() bool {
	return !s.CustomStyle()
}

// ===========================================================================
// CT_Styles — style ID resolution
// ===========================================================================

// GetStyleIDByName resolves a UI style name to its XML style ID string.
//
// Applies BabelFish translation, looks up by name (with fallback to ID),
// validates the style type, and returns nil if the style is the default
// for its type.
//
// This is the full algorithm that Python spreads across:
//   - Styles.__getitem__     (BabelFish + get_by_name + fallback get_by_id)
//   - Styles._get_style_id_from_name  (delegates to __getitem__ + _get_style_id_from_style)
//   - Styles._get_style_id_from_style (type check + default check)
func (ss *CT_Styles) GetStyleIDByName(uiName string, styleType enum.WdStyleType) (*string, error) {
	// 1. BabelFish translation (Python: BabelFish.ui2internal)
	internalName := UI2Internal(uiName)

	// 2. Lookup by name (Python: self._element.get_by_name)
	s := ss.GetByName(internalName)

	// 3. Fallback to ID (Python: self._element.get_by_id — deprecated path)
	if s == nil {
		s = ss.GetByID(uiName)
	}

	// 4. Not found → error (Python: raise KeyError)
	if s == nil {
		return nil, fmt.Errorf("oxml: no style with name %q", uiName)
	}

	// 5. Type check (Python: _get_style_id_from_style raises ValueError)
	xmlType, err := styleType.ToXml()
	if err != nil {
		return nil, fmt.Errorf("oxml: invalid style type: %w", err)
	}
	if s.Type() != xmlType {
		return nil, fmt.Errorf("oxml: style %q is type %q, need %q", uiName, s.Type(), xmlType)
	}

	// 6. Default check → return nil (Python: if style == self.default(style_type): return None)
	def, err := ss.DefaultFor(styleType)
	if err != nil {
		return nil, fmt.Errorf("oxml: getting default style: %w", err)
	}
	if def != nil && def.StyleId() == s.StyleId() {
		return nil, nil
	}

	// 7. Return style ID
	id := s.StyleId()
	return &id, nil
}

// ===========================================================================
// CT_LatentStyles — custom methods
// ===========================================================================

// GetByName returns the lsdException child with the given name, or nil.
func (ls *CT_LatentStyles) GetByName(name string) *CT_LsdException {
	for _, exc := range ls.LsdExceptionList() {
		n, err := exc.Name()
		if err == nil && n == name {
			return exc
		}
	}
	return nil
}

// BoolProp returns the boolean value of the named attribute, or false if absent.
func (ls *CT_LatentStyles) BoolProp(attrName string) bool {
	val, ok := ls.GetAttr(attrName)
	if !ok {
		return false
	}
	return parseBoolAttr(val)
}

// SetBoolProp sets the named on/off attribute.
func (ls *CT_LatentStyles) SetBoolProp(attrName string, val bool) error {
	s, err := formatBoolAttr(val)
	if err != nil {
		return err
	}
	ls.SetAttr(attrName, s)
	return nil
}

// ===========================================================================
// CT_LsdException — custom methods
// ===========================================================================

// Delete removes this lsdException element from its parent.
func (exc *CT_LsdException) Delete() {
	parent := exc.e.Parent()
	if parent != nil {
		parent.RemoveChild(exc.e)
	}
}

// OnOffProp returns the boolean value of the named attribute, or nil if absent.
func (exc *CT_LsdException) OnOffProp(attrName string) *bool {
	val, ok := exc.GetAttr(attrName)
	if !ok {
		return nil
	}
	v := parseBoolAttr(val)
	return &v
}

// SetOnOffProp sets the named on/off attribute. Passing nil removes the attribute.
func (exc *CT_LsdException) SetOnOffProp(attrName string, val *bool) error {
	if val == nil {
		exc.RemoveAttr(attrName)
		return nil
	}
	s, err := formatBoolAttr(*val)
	if err != nil {
		return err
	}
	exc.SetAttr(attrName, s)
	return nil
}
