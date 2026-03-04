package main

import "github.com/vortex/go-docx/pkg/docx"

// repl pairs a search string with the TableData to insert.
type repl struct {
	old string
	td  docx.TableData
}

// tbl is a shorthand for building TableData with "Table Grid" style.
func tbl(rows ...[]string) docx.TableData {
	return docx.TableData{
		Rows:  rows,
		Style: docx.StyleName("Table Grid"),
	}
}

// row is a shorthand for building a table row.
func row(cells ...string) []string { return cells }

// allReplacements returns every (tag, TableData) pair applied to the "after" doc.
func allReplacements() []repl {
	return []repl{
		// §1  Tag is entire paragraph text
		{"|<ENTIRE>|", tbl(row("A1", "B1"), row("A2", "B2"))},

		// §2  Text only before tag
		{"|<AFTER_TEXT>|", tbl(row("Col1", "Col2"), row("X", "Y"))},

		// §3  Text only after tag
		{"|<BEFORE_TEXT>|", tbl(row("Alpha", "Beta"), row("1", "2"))},

		// §4  Text on both sides
		{"|<BOTH_SIDES>|", tbl(row("L", "R"), row("left", "right"))},

		// §5  Cross-run tag (split across formatted runs)
		{"|<CROSSRUN>|", tbl(row("Cross", "Run"), row("tag", "split"))},

		// §6  Multiple tags in one paragraph
		{"|<MULTI_TAG>|", tbl(row("M1", "M2"))},

		// §7  Section break paragraph
		{"|<SECT_BRK>|", tbl(row("Sect", "Data"), row("s1", "s2"))},

		// §8  Tag in table cell → nested table
		{"|<NESTED>|", tbl(row("Inner1", "Inner2"), row("Nest1", "Nest2"))},

		// §9  Tag in header
		{"|<HDR_TBL>|", tbl(row("Header", "Table"))},

		// §10 Tag in footer
		{"|<FTR_TBL>|", tbl(row("Footer", "Table"))},

		// §11 Multiple occurrences across paragraphs
		{"|<MULTI_OCC>|", tbl(row("Occ1", "Occ2"))},

		// §12 No match — applied to a tag that doesn't exist
		{"|<NO_MATCH_TAG>|", tbl(row("Never", "Seen"))},

		// §13 Empty old — returns 0
		{"", tbl(row("Should", "Not", "Appear"))},

		// §14 Cyrillic / UTF-8 tag
		{"|<КИРИЛЛИЦА>|", tbl(row("Кол1", "Кол2"), row("Дан1", "Дан2"))},

		// §15 Tag at paragraph start (no text before)
		{"|<AT_START>|", tbl(row("Start1", "Start2"))},

		// §16 Tag at paragraph end (no text after)
		{"|<AT_END>|", tbl(row("End1", "End2"))},

		// §17 Tag is entire text of table cell
		{"|<CELL_ENTIRE>|", tbl(row("CE1", "CE2"))},

		// §18 Tag in cell with surrounding text
		{"|<CELL_MIXED>|", tbl(row("CM1", "CM2"))},

		// §19 Comment body with tag
		{"|<COMMENT_TBL>|", tbl(row("Cmt1", "Cmt2"))},

		// §20 First-page header
		{"|<FIRST_HDR_TBL>|", tbl(row("1stHdr", "Table"))},

		// §21 First-page footer
		{"|<FIRST_FTR_TBL>|", tbl(row("1stFtr", "Table"))},

		// §22 Large table (many rows/cols)
		{"|<BIG_TABLE>|", tbl(
			row("H1", "H2", "H3", "H4"),
			row("R1C1", "R1C2", "R1C3", "R1C4"),
			row("R2C1", "R2C2", "R2C3", "R2C4"),
			row("R3C1", "R3C2", "R3C3", "R3C4"),
			row("R4C1", "R4C2", "R4C3", "R4C4"),
		)},

		// §23 Single-cell table
		{"|<SINGLE_CELL>|", tbl(row("Only cell"))},

		// §24 Tag between two section breaks
		{"|<MID_SECT>|", tbl(row("Mid", "Section"))},

		// §25 Multiple tags in one paragraph (second tag, uses same |<MULTI_TAG>|)
		// — handled by the same |<MULTI_TAG>| replacement (appears 2× in one para)

		// §26 Tag in header with table (tag inside a table cell in the header)
		{"|<HDR_CELL_TBL>|", tbl(row("HC1", "HC2"))},
	}
}
