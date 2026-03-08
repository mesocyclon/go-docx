package main

import "github.com/vortex/go-docx/pkg/docx"

type replacement struct {
	tag    string
	source *docx.Document
	format docx.ImportFormatMode
	opts   docx.ImportFormatOptions
}

func allReplacements(sources map[string]*docx.Document) []replacement {
	s := func(name string) *docx.Document { return sources[name] }
	return []replacement{
		// ── A. Tag position tests ──────────────────────────────────
		{tag: "(<ENTIRE>)", source: s("src_paragraph.docx")},
		{tag: "(<AFTER>)", source: s("src_paragraph.docx")},
		{tag: "(<BEFORE>)", source: s("src_paragraph.docx")},
		{tag: "(<BOTH>)", source: s("src_paragraph.docx")},
		{tag: "(<CROSS>)", source: s("src_paragraph.docx")},
		{tag: "(<MULTI>)", source: s("src_paragraph.docx")},
		{tag: "(<SECT>)", source: s("src_paragraph.docx")},
		{tag: "(<CELL_FULL>)", source: s("src_paragraph.docx")},
		{tag: "(<CELL_MIX>)", source: s("src_paragraph.docx")},
		{tag: "(<HDR>)", source: s("src_paragraph.docx")},
		{tag: "(<FTR>)", source: s("src_paragraph.docx")},
		{tag: "(<FP_HDR>)", source: s("src_paragraph.docx")},
		{tag: "(<FP_FTR>)", source: s("src_paragraph.docx")},
		{tag: "(<COMMENT>)", source: s("src_paragraph.docx")},
		{tag: "(<HDR_CELL>)", source: s("src_paragraph.docx")},
		{tag: "(<REPEAT>)", source: s("src_paragraph.docx")},
		{tag: "(<TAG_START>)", source: s("src_paragraph.docx")},
		{tag: "(<TAG_END>)", source: s("src_paragraph.docx")},
		{tag: "(<BETWEEN>)", source: s("src_paragraph.docx")},

		// ── B. Source content type tests ───────────────────────────
		{tag: "(<SRC_MULTI>)", source: s("src_multi_para.docx")},
		{tag: "(<SRC_TABLE>)", source: s("src_table.docx")},
		{tag: "(<SRC_COMPLEX>)", source: s("src_complex.docx")},
		{tag: "(<SRC_IMAGE>)", source: s("src_image.docx")},
		{tag: "(<SRC_HEADING>)", source: s("src_heading.docx")},
		{tag: "(<SRC_EMPTY>)", source: s("src_empty.docx")},
		{tag: "(<SRC_LARGE>)", source: s("src_large.docx")},
		{tag: "(<SRC_STYLED>)", source: s("src_styled.docx")},

		// ── C. Cyrillic tests ─────────────────────────────────────
		{tag: "(<КИРИЛЛИЦА>)", source: s("src_cyrillic.docx")},
		{tag: "(<ЗАМЕНА>)", source: s("src_paragraph_ru.docx")},

		// ── D. Multi-source tests ─────────────────────────────────
		{tag: "(<SRC_A>)", source: s("src_a.docx")},
		{tag: "(<SRC_B>)", source: s("src_b.docx")},

		// ── E. No-match (tag absent from template) ────────────────
		{tag: "(<NO_MATCH>)", source: s("src_paragraph.docx")},

		// ── F. ImportFormatMode tests ────────────────────────────
		// Е1: UseDestinationStyles — target Heading 1 styling wins.
		{tag: "(<IFM_USE_DEST>)", source: s("src_conflict_style.docx"),
			format: docx.UseDestinationStyles},
		// Е2: KeepSourceFormatting — source red Heading 1 expanded to direct attrs.
		{tag: "(<IFM_KEEP_SRC>)", source: s("src_conflict_style.docx"),
			format: docx.KeepSourceFormatting},
		// Е3: KeepSourceFormatting + ForceCopyStyles — style copied with suffix.
		{tag: "(<IFM_KEEP_SRC_FORCE>)", source: s("src_conflict_style.docx"),
			format: docx.KeepSourceFormatting, opts: docx.ImportFormatOptions{ForceCopyStyles: true}},
		// Е4: KeepDifferentStyles — formatting differs → expand to direct attrs.
		{tag: "(<IFM_KEEP_DIFF>)", source: s("src_conflict_style.docx"),
			format: docx.KeepDifferentStyles},
		// Е5: KeepDifferentStyles + ForceCopyStyles — different → copy with suffix.
		{tag: "(<IFM_KEEP_DIFF_FORCE>)", source: s("src_conflict_style.docx"),
			format: docx.KeepDifferentStyles, opts: docx.ImportFormatOptions{ForceCopyStyles: true}},
		// Е6: KeepSourceNumbering — list numbering preserved as separate definition.
		{tag: "(<IFM_KEEP_NUM>)", source: s("src_numbered_list.docx"),
			opts: docx.ImportFormatOptions{KeepSourceNumbering: true}},
		// Е7: Default numbering merge (KeepSourceNumbering=false).
		{tag: "(<IFM_MERGE_NUM>)", source: s("src_numbered_list.docx")},
		// Е8: KeepSourceFormatting + unique style not in target → deep-copied.
		{tag: "(<IFM_UNIQUE_STYLE>)", source: s("src_unique_style.docx"),
			format: docx.KeepSourceFormatting},
	}
	// (<SELF>) and "" handled separately in main().
}
