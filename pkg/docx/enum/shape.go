package enum

// ---------------------------------------------------------------------------
// WdInlineShapeType (alias: WdInlineShape) — no XML mapping
// ---------------------------------------------------------------------------

// WdInlineShapeType specifies the type of an inline shape.
// MS API name: WdInlineShapeType
type WdInlineShapeType int

const (
	WdInlineShapeTypeChart          WdInlineShapeType = 12
	WdInlineShapeTypeLinkedPicture  WdInlineShapeType = 4
	WdInlineShapeTypePicture        WdInlineShapeType = 3
	WdInlineShapeTypeSmartArt       WdInlineShapeType = 15
	WdInlineShapeTypeNotImplemented WdInlineShapeType = -6
)

// WdInlineShape is an alias for WdInlineShapeType.
type WdInlineShape = WdInlineShapeType
