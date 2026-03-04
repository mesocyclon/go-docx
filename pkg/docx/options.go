package docx

import "github.com/vortex/go-docx/pkg/docx/enum"

// ---------------------------------------------------------------------------
// StyleRef — replaces interface{} for style parameters
// ---------------------------------------------------------------------------

// StyleRef identifies a style for paragraphs, runs, and tables.
// Use StyleName("Heading 1") for a name or pass a *BaseStyle obtained from
// the Styles collection.
type StyleRef interface {
	isStyleRef() // sealed
}

// StyleName references a style by its display name (e.g. "Heading 1").
type StyleName string

func (StyleName) isStyleRef() {}

// resolveStyleRef extracts the raw value (string or *BaseStyle) from an
// optional variadic StyleRef for passing to the parts layer.
func resolveStyleRef(style []StyleRef) any {
	if len(style) == 0 || style[0] == nil {
		return nil
	}
	switch v := style[0].(type) {
	case StyleName:
		return string(v)
	case *BaseStyle:
		return v
	default:
		return nil
	}
}

// ---------------------------------------------------------------------------
// UnderlineVal — replaces interface{} for underline get/set
// ---------------------------------------------------------------------------

// UnderlineVal represents an underline setting.
// Use the Underline* constructors below.
type UnderlineVal struct {
	kind underlineKind
	wdU  enum.WdUnderline
}

type underlineKind int

const (
	underlineKindSingle underlineKind = iota
	underlineKindNone
	underlineKindStyle
)

// UnderlineSingle returns an UnderlineVal for standard single underline.
func UnderlineSingle() UnderlineVal { return UnderlineVal{kind: underlineKindSingle} }

// UnderlineNone returns an UnderlineVal that explicitly disables underline.
func UnderlineNone() UnderlineVal { return UnderlineVal{kind: underlineKindNone} }

// UnderlineStyle returns an UnderlineVal for a specific underline style.
func UnderlineStyle(u enum.WdUnderline) UnderlineVal {
	return UnderlineVal{kind: underlineKindStyle, wdU: u}
}

// IsSingle reports whether this is a standard single underline.
func (u UnderlineVal) IsSingle() bool { return u.kind == underlineKindSingle }

// IsNone reports whether underline is explicitly disabled.
func (u UnderlineVal) IsNone() bool { return u.kind == underlineKindNone }

// IsStyle reports whether this is a specific WdUnderline style.
func (u UnderlineVal) IsStyle() bool { return u.kind == underlineKindStyle }

// Style returns the WdUnderline style. Only meaningful when IsStyle() is true.
func (u UnderlineVal) Style() enum.WdUnderline { return u.wdU }

// ---------------------------------------------------------------------------
// LineSpacingVal — replaces interface{} for line spacing get/set
// ---------------------------------------------------------------------------

// LineSpacingVal represents a line spacing value.
// It is either a multiple of the normal line height (e.g. 1.5) or an
// absolute distance in twips.
type LineSpacingVal struct {
	isMultiple bool
	multiple   float64
	twips      int
}

// LineSpacingMultiple creates a LineSpacingVal expressed as a multiple of the
// standard single-line height (e.g. 2.0 for double-spacing).
func LineSpacingMultiple(v float64) LineSpacingVal {
	return LineSpacingVal{isMultiple: true, multiple: v}
}

// LineSpacingTwips creates a LineSpacingVal expressed as an absolute distance
// in twips (twentieth of a point).
func LineSpacingTwips(v int) LineSpacingVal {
	return LineSpacingVal{isMultiple: false, twips: v}
}

// IsMultiple reports whether this value is a line-height multiple.
func (v LineSpacingVal) IsMultiple() bool { return v.isMultiple }

// Multiple returns the multiple value (e.g. 2.0). Only meaningful when
// IsMultiple() is true.
func (v LineSpacingVal) Multiple() float64 { return v.multiple }

// Twips returns the absolute twips value. Only meaningful when IsMultiple()
// is false.
func (v LineSpacingVal) Twips() int { return v.twips }

// ---------------------------------------------------------------------------
// InlineItem — replaces []interface{} for IterInnerContent
// ---------------------------------------------------------------------------

// InlineItem represents either a *Run or a *Hyperlink found inside a
// paragraph. Callers inspect the type via Run() / Hyperlink().
type InlineItem struct {
	run       *Run
	hyperlink *Hyperlink
}

// IsRun reports whether this item is a Run.
func (it *InlineItem) IsRun() bool { return it.run != nil }

// IsHyperlink reports whether this item is a Hyperlink.
func (it *InlineItem) IsHyperlink() bool { return it.hyperlink != nil }

// Run returns the Run, or nil if this item is a Hyperlink.
func (it *InlineItem) Run() *Run { return it.run }

// Hyperlink returns the Hyperlink, or nil if this item is a Run.
func (it *InlineItem) Hyperlink() *Hyperlink { return it.hyperlink }
