package main

import "github.com/vortex/go-docx/pkg/docx"

type replacement struct {
	tag    string
	source *docx.Document
}

func allReplacements(sources map[string]*docx.Document) []replacement {
	s := func(name string) *docx.Document { return sources[name] }
	return []replacement{
		// ── A. Tag position tests ──────────────────────────────────
		{"(<ENTIRE>)", s("src_paragraph.docx")},
		{"(<AFTER>)", s("src_paragraph.docx")},
		{"(<BEFORE>)", s("src_paragraph.docx")},
		{"(<BOTH>)", s("src_paragraph.docx")},
		{"(<CROSS>)", s("src_paragraph.docx")},
		{"(<MULTI>)", s("src_paragraph.docx")},
		{"(<SECT>)", s("src_paragraph.docx")},
		{"(<CELL_FULL>)", s("src_paragraph.docx")},
		{"(<CELL_MIX>)", s("src_paragraph.docx")},
		{"(<HDR>)", s("src_paragraph.docx")},
		{"(<FTR>)", s("src_paragraph.docx")},
		{"(<FP_HDR>)", s("src_paragraph.docx")},
		{"(<FP_FTR>)", s("src_paragraph.docx")},
		{"(<COMMENT>)", s("src_paragraph.docx")},
		{"(<HDR_CELL>)", s("src_paragraph.docx")},
		{"(<REPEAT>)", s("src_paragraph.docx")},
		{"(<TAG_START>)", s("src_paragraph.docx")},
		{"(<TAG_END>)", s("src_paragraph.docx")},
		{"(<BETWEEN>)", s("src_paragraph.docx")},

		// ── B. Source content type tests ───────────────────────────
		{"(<SRC_MULTI>)", s("src_multi_para.docx")},
		{"(<SRC_TABLE>)", s("src_table.docx")},
		{"(<SRC_COMPLEX>)", s("src_complex.docx")},
		{"(<SRC_IMAGE>)", s("src_image.docx")},
		{"(<SRC_HEADING>)", s("src_heading.docx")},
		{"(<SRC_EMPTY>)", s("src_empty.docx")},
		{"(<SRC_LARGE>)", s("src_large.docx")},
		{"(<SRC_STYLED>)", s("src_styled.docx")},

		// ── C. Cyrillic tests ─────────────────────────────────────
		{"(<КИРИЛЛИЦА>)", s("src_cyrillic.docx")},
		{"(<ЗАМЕНА>)", s("src_paragraph_ru.docx")},

		// ── D. Multi-source tests ─────────────────────────────────
		{"(<SRC_A>)", s("src_a.docx")},
		{"(<SRC_B>)", s("src_b.docx")},

		// ── E. No-match (tag absent from template) ────────────────
		{"(<NO_MATCH>)", s("src_paragraph.docx")},
	}
	// (<SELF>) and "" handled separately in main().
}
