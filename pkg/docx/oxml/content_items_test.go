package oxml

import (
	"testing"
)

func TestBlockItemInterface(t *testing.T) {
	t.Parallel()
	var bi BlockItem

	bi = &CT_P{Element{e: OxmlElement("w:p")}}
	bi.isBlockItem() // compile-time + runtime verification

	bi = &CT_Tbl{Element{e: OxmlElement("w:tbl")}}
	bi.isBlockItem()
	_ = bi
}

func TestInlineItemInterface(t *testing.T) {
	t.Parallel()
	var ii InlineItem

	ii = &CT_R{Element{e: OxmlElement("w:r")}}
	ii.isInlineItem()

	ii = &CT_Hyperlink{Element{e: OxmlElement("w:hyperlink")}}
	ii.isInlineItem()
	_ = ii
}
