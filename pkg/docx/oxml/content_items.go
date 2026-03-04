package oxml

// BlockItem is a sealed interface representing block-level content elements
// that can appear inside story containers (body, table cell, comment, header/footer).
// Implemented by *CT_P and *CT_Tbl.
//
// This corresponds to Python's `CT_P | CT_Tbl` union type used in
// inner_content_elements properties.
type BlockItem interface {
	isBlockItem()
}

// Verify at compile time that CT_P and CT_Tbl implement BlockItem.
var (
	_ BlockItem = (*CT_P)(nil)
	_ BlockItem = (*CT_Tbl)(nil)
)

func (*CT_P) isBlockItem()   {}
func (*CT_Tbl) isBlockItem() {}

// InlineItem is a sealed interface representing inline-level content elements
// that can appear inside a paragraph.
// Implemented by *CT_R and *CT_Hyperlink.
//
// This corresponds to Python's `CT_R | CT_Hyperlink` union type used in
// CT_P.inner_content_elements.
type InlineItem interface {
	isInlineItem()
}

// Verify at compile time that CT_R and CT_Hyperlink implement InlineItem.
var (
	_ InlineItem = (*CT_R)(nil)
	_ InlineItem = (*CT_Hyperlink)(nil)
)

func (*CT_R) isInlineItem()         {}
func (*CT_Hyperlink) isInlineItem() {}
