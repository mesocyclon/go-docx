package main

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
	. "github.com/vortex/go-docx/visual-regtest/internal/docfmt"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// buildBeforeDocument creates the "before" document from scratch.
func buildBeforeDocument() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	// ---- Section 1 setup: headers / footers ----
	sect, err := doc.Sections().Get(0)
	if err != nil {
		return nil, fmt.Errorf("getting default section: %w", err)
	}

	// Primary header: tag for table replacement
	hdr := sect.Header()
	BuildHighlightedParagraph(hdr, "|<HDR_TBL>|")

	// Primary header: also a table in header, with tag inside a cell
	hdrTbl, err := hdr.AddTable(1, 2, 9000)
	if err != nil {
		return nil, err
	}
	hc0, _ := hdrTbl.CellAt(0, 0)
	hc0.SetText("Header Left")
	hc1, _ := hdrTbl.CellAt(0, 1)
	hc1.SetText("|<HDR_CELL_TBL>|")

	// Primary footer: tag for table replacement
	ftr := sect.Footer()
	BuildHighlightedParagraph(ftr, "|<FTR_TBL>|")

	// First-page header/footer
	if err := sect.SetDifferentFirstPageHeaderFooter(true); err != nil {
		return nil, err
	}
	BuildHighlightedParagraph(sect.FirstPageHeader(), "|<FIRST_HDR_TBL>|")
	BuildHighlightedParagraph(sect.FirstPageFooter(), "|<FIRST_FTR_TBL>|")

	// ================================================================
	// Legend
	// ================================================================
	Heading(doc, "ReplaceWithTable — Visual Regression Test", 0)

	lp1, _ := doc.AddParagraph("")
	AddPlain(lp1, "How to read this document:  ")
	r, _ := lp1.AddRun("Yellow highlight")
	_ = r.SetBold(BoolPtr(true))
	SetHighlightYellow(r)
	AddPlain(lp1, " = placeholder (will be replaced with a table).  ")
	r, _ = lp1.AddRun("Green highlight")
	_ = r.SetBold(BoolPtr(true))
	SetHighlightGreen(r)
	AddPlain(lp1, " = expected table content (shown in spec lines).")

	lp2, _ := doc.AddParagraph("")
	AddPlain(lp2, "In ")
	AddBold(lp2, "02_after_replace_tbl.docx")
	AddPlain(lp2, ": yellow text should be replaced by bordered tables. "+
		"Compare inserted table content vs green expected values in each spec line.")

	Para(doc, "")

	// ================================================================
	// §1 — Tag is entire paragraph text (case 1)
	// ================================================================
	Heading(doc, "1. Tag Is Entire Paragraph Text", 1)
	SpecTable(doc, "|<ENTIRE>|", "A1,B1 / A2,B2")
	Note(doc, "The paragraph containing only the tag is removed entirely. "+
		"A 2×2 bordered table replaces it. No empty paragraphs before or after.")

	TagPara(doc, "", "|<ENTIRE>|", "")

	// ================================================================
	// §2 — Text only before tag (case 2)
	// ================================================================
	Heading(doc, "2. Text Only Before Tag", 1)
	SpecTable(doc, "|<AFTER_TEXT>|", "Col1,Col2 / X,Y")
	Note(doc, "Paragraph is split: text before stays as a paragraph, "+
		"table inserted after it. No text after.")

	TagPara(doc, "Text before the table: ", "|<AFTER_TEXT>|", "")

	// ================================================================
	// §3 — Text only after tag (case 3)
	// ================================================================
	Heading(doc, "3. Text Only After Tag", 1)
	SpecTable(doc, "|<BEFORE_TEXT>|", "Alpha,Beta / 1,2")
	Note(doc, "Table inserted first, text after stays as a paragraph below.")

	TagPara(doc, "", "|<BEFORE_TEXT>|", " — text after the table")

	// ================================================================
	// §4 — Text on both sides (case 4)
	// ================================================================
	Heading(doc, "4. Text On Both Sides", 1)
	SpecTable(doc, "|<BOTH_SIDES>|", "L,R / left,right")
	Note(doc, "Paragraph splits into: text-before paragraph, table, text-after paragraph.")

	TagPara(doc, "Left text → ", "|<BOTH_SIDES>|", " ← right text")

	// ================================================================
	// §5 — Cross-run tag (case 5)
	// ================================================================
	Heading(doc, "5. Cross-Run Tag (Formatting Preserved)", 1)
	SpecTable(doc, "|<CROSSRUN>|", "Cross,Run / tag,split")
	Note(doc, "Tag split across 3 runs: '|<CROSS' (bold red) + 'RU' (italic blue) + 'N>|' (normal). "+
		"Surrounding text formatting preserved. Paragraph splits around the table.")

	p5, _ := doc.AddParagraph("")
	AddPlain(p5, "Before marker: ")
	cr1, _ := p5.AddRun("|<CROSS")
	_ = cr1.SetBold(BoolPtr(true))
	_ = cr1.Font().Color().SetRGB(&ColorRed)
	SetHighlightYellow(cr1)
	cr2, _ := p5.AddRun("RU")
	_ = cr2.SetItalic(BoolPtr(true))
	_ = cr2.Font().Color().SetRGB(&ColorBlue)
	SetHighlightYellow(cr2)
	cr3, _ := p5.AddRun("N>|")
	SetHighlightYellow(cr3)
	AddPlain(p5, " — after marker")

	// ================================================================
	// §6 — Multiple tags in one paragraph (case 6)
	// ================================================================
	Heading(doc, "6. Multiple Tags In One Paragraph", 1)
	SpecTable(doc, "|<MULTI_TAG>| (×2 in same paragraph)", "M1,M2")
	Note(doc, "Two occurrences of the same tag in one paragraph. "+
		"Result: text, table, text, table, text — five sibling elements.")

	p6, _ := doc.AddParagraph("")
	AddPlain(p6, "First: ")
	AddHighlighted(p6, "|<MULTI_TAG>|")
	AddPlain(p6, " middle text ")
	AddHighlighted(p6, "|<MULTI_TAG>|")
	AddPlain(p6, " last.")

	// ================================================================
	// §7 — Section break paragraph (case 7)
	// ================================================================
	Heading(doc, "7. Section Break Paragraph (sectPr Preserved)", 1)
	SpecTable(doc, "|<SECT_BRK>|", "Sect,Data / s1,s2")
	Note(doc, "Tag is in the last paragraph of a section (carries sectPr). "+
		"After replacement: table is in the same section, sectPr moves to a trailing paragraph. "+
		"Page layout must not change.")

	TagPara(doc, "Before section break table: ", "|<SECT_BRK>|", " — end of section 1")

	// Add section break — this paragraph becomes the last in section 1
	sect2, err := doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		return nil, err
	}
	_ = sect2

	Heading(doc, "— Page 2 (Section 2 — verifies sectPr survived) —", 2)
	Note(doc, "If you see this on a new page, section break was preserved correctly.")

	// ================================================================
	// §8 — Tag in table cell → nested table (case 9)
	// ================================================================
	Heading(doc, "8. Tag In Table Cell (Nested Table)", 1)
	SpecTable(doc, "|<NESTED>|", "Inner1,Inner2 / Nest1,Nest2")
	Note(doc, "Outer 2×2 table. Cell (0,1) contains the tag. "+
		"After replacement: inner table appears inside that cell. "+
		"Cell must still have at least one <w:p> (OOXML invariant).")

	outer, _ := doc.AddTable(2, 2)
	oc00, _ := outer.CellAt(0, 0)
	oc00.SetText("Outer 0,0 (untouched)")
	oc01, _ := outer.CellAt(0, 1)
	oc01.SetText("|<NESTED>|")
	oc10, _ := outer.CellAt(1, 0)
	oc10.SetText("Outer 1,0 (untouched)")
	oc11, _ := outer.CellAt(1, 1)
	oc11.SetText("Outer 1,1 (untouched)")

	// ================================================================
	// §9 — Tag in header (case 10) — already set up in header above
	// ================================================================
	Heading(doc, "9. Tag In Primary Header", 1)
	SpecTable(doc, "|<HDR_TBL>|", "Header,Table")
	Note(doc, "Primary header contains the tag as a standalone paragraph. "+
		"After replacement: a table appears in the header (visible on page 2+).")

	// ================================================================
	// §10 — Tag in footer (case 10) — already set up in footer above
	// ================================================================
	Heading(doc, "10. Tag In Primary Footer", 1)
	SpecTable(doc, "|<FTR_TBL>|", "Footer,Table")
	Note(doc, "Primary footer contains the tag. "+
		"After replacement: a table appears in the footer (visible on page 2+).")

	// ================================================================
	// §11 — Multiple occurrences across paragraphs
	// ================================================================
	Heading(doc, "11. Multiple Occurrences Across Paragraphs", 1)
	SpecTable(doc, "|<MULTI_OCC>| (×3 across paragraphs)", "Occ1,Occ2")
	Note(doc, "Three paragraphs each contain the tag. "+
		"All three are replaced independently. Each gets its own fresh table.")

	TagPara(doc, "First occurrence: ", "|<MULTI_OCC>|", "")
	TagPara(doc, "Second occurrence: ", "|<MULTI_OCC>|", "")
	TagPara(doc, "Third occurrence: ", "|<MULTI_OCC>|", "")

	// ================================================================
	// §12 — No match (tag not present)
	// ================================================================
	Heading(doc, "12. No Match (Tag Not In Document)", 1)
	Note(doc, "|<NO_MATCH_TAG>| is not present anywhere. Returns 0 replacements. "+
		"Document is unchanged. This paragraph must be identical in both files.")
	Para(doc, "This text has no tag. Nothing changes here.")

	// ================================================================
	// §13 — Empty old (returns 0)
	// ================================================================
	Heading(doc, "13. Empty Search String (old == \"\")", 1)
	Note(doc, "ReplaceWithTable(\"\", ...) returns 0. No modifications. "+
		"This paragraph must be identical in both files.")
	Para(doc, "Nothing should change here either. Absolutely identical in both files.")

	// ================================================================
	// §14 — Cyrillic / UTF-8 tag
	// ================================================================
	Heading(doc, "14. Cyrillic / UTF-8 Tag", 1)
	SpecTable(doc, "|<КИРИЛЛИЦА>|", "Кол1,Кол2 / Дан1,Дан2")
	Note(doc, "Multibyte tag. Must be found and replaced correctly. "+
		"Surrounding Cyrillic text preserved.")

	TagPara(doc, "Перед таблицей: ", "|<КИРИЛЛИЦА>|", " — после таблицы")

	// ================================================================
	// §15 — Tag at paragraph start
	// ================================================================
	Heading(doc, "15. Tag At Paragraph Start", 1)
	SpecTable(doc, "|<AT_START>|", "Start1,Start2")
	Note(doc, "Tag at the very beginning. No before-paragraph created (empty text omitted). "+
		"Only table + after-paragraph.")

	TagPara(doc, "", "|<AT_START>|", " is at the very beginning.")

	// ================================================================
	// §16 — Tag at paragraph end
	// ================================================================
	Heading(doc, "16. Tag At Paragraph End", 1)
	SpecTable(doc, "|<AT_END>|", "End1,End2")
	Note(doc, "Tag at the very end. No after-paragraph created (empty text omitted). "+
		"Only before-paragraph + table.")

	TagPara(doc, "This paragraph ends with a table: ", "|<AT_END>|", "")

	// ================================================================
	// §17 — Tag is entire text of table cell (case 1 + case 9)
	// ================================================================
	Heading(doc, "17. Tag Is Entire Text of Table Cell", 1)
	SpecTable(doc, "|<CELL_ENTIRE>|", "CE1,CE2")
	Note(doc, "The tag is the only content of a table cell. "+
		"After replacement: cell contains nested table + mandatory trailing <w:p>. "+
		"Other cells unchanged.")

	tbl17, _ := doc.AddTable(2, 2)
	ce00, _ := tbl17.CellAt(0, 0)
	ce00.SetText("|<CELL_ENTIRE>|")
	ce01, _ := tbl17.CellAt(0, 1)
	ce01.SetText("Adjacent (untouched)")
	ce10, _ := tbl17.CellAt(1, 0)
	ce10.SetText("Below (untouched)")
	ce11, _ := tbl17.CellAt(1, 1)
	ce11.SetText("Diagonal (untouched)")

	// ================================================================
	// §18 — Tag in cell with surrounding text (case 4 + case 9)
	// ================================================================
	Heading(doc, "18. Tag In Cell With Surrounding Text", 1)
	SpecTable(doc, "|<CELL_MIXED>|", "CM1,CM2")
	Note(doc, "Cell text: 'Before |<CELL_MIXED>| after'. "+
		"After replacement: cell has before-paragraph, nested table, after-paragraph.")

	tbl18, _ := doc.AddTable(1, 2)
	cm0, _ := tbl18.CellAt(0, 0)
	cm0.SetText("Before |<CELL_MIXED>| after")
	cm1, _ := tbl18.CellAt(0, 1)
	cm1.SetText("Other cell (untouched)")

	// ================================================================
	// §19 — Comment body with tag
	// ================================================================
	Heading(doc, "19. Tag In Comment Body", 1)
	SpecTable(doc, "|<COMMENT_TBL>|", "Cmt1,Cmt2")
	Note(doc, "The comment body text contains the tag. "+
		"ReplaceWithTable must reach into word/comments.xml. "+
		"The annotated body text stays unchanged.")

	p19, _ := doc.AddParagraph("")
	AddPlain(p19, "This text has a comment whose body contains a table tag: ")
	p19run, _ := p19.AddRun("see comment")
	SetHighlightYellow(p19run)
	initials19 := "TB"
	if _, err := doc.AddComment(
		[]*docx.Run{p19run},
		"Review: |<COMMENT_TBL>| — needs a table here",
		"Table Comment Test", &initials19,
	); err != nil {
		return nil, fmt.Errorf("adding comment for §19: %w", err)
	}

	// ================================================================
	// §20 — First-page header (case 10) — already set up above
	// ================================================================
	Heading(doc, "20. Tag In First-Page Header", 1)
	SpecTable(doc, "|<FIRST_HDR_TBL>|", "1stHdr,Table")
	Note(doc, "First-page header contains the tag. "+
		"Visible on page 1. After replacement: table in first-page header.")

	// ================================================================
	// §21 — First-page footer (case 10) — already set up above
	// ================================================================
	Heading(doc, "21. Tag In First-Page Footer", 1)
	SpecTable(doc, "|<FIRST_FTR_TBL>|", "1stFtr,Table")
	Note(doc, "First-page footer contains the tag. "+
		"Visible on page 1. After replacement: table in first-page footer.")

	// ================================================================
	// §22 — Large table (stress test)
	// ================================================================
	Heading(doc, "22. Large Table (5×4)", 1)
	SpecTable(doc, "|<BIG_TABLE>|", "H1..H4 / R1..R4 × C1..C4")
	Note(doc, "Tag replaced with a 5-row, 4-column table. "+
		"Verifies column width distribution for wider tables.")

	TagPara(doc, "Before big table: ", "|<BIG_TABLE>|", " — after big table")

	// ================================================================
	// §23 — Single-cell table
	// ================================================================
	Heading(doc, "23. Single-Cell Table (1×1)", 1)
	SpecTable(doc, "|<SINGLE_CELL>|", "Only cell")
	Note(doc, "Minimal table: 1 row, 1 column. Still must have borders and valid structure.")

	TagPara(doc, "", "|<SINGLE_CELL>|", "")

	// ================================================================
	// §24 — Tag between two section breaks
	// ================================================================
	Heading(doc, "24. Tag Between Two Section Breaks", 1)
	SpecTable(doc, "|<MID_SECT>|", "Mid,Section")
	Note(doc, "Tag is in a paragraph that is between two section breaks "+
		"(in the middle section). Both section boundaries must survive.")

	// End current section (section 2 → section 3)
	TagPara(doc, "Mid-section tag: ", "|<MID_SECT>|", " — end of mid section")
	sect3, err := doc.AddSection(enum.WdSectionStartNewPage)
	if err != nil {
		return nil, err
	}
	_ = sect3

	Heading(doc, "— Page 3 (Section 3 — after mid-section test) —", 2)
	Note(doc, "If you see this on a new page, both section breaks survived.")

	// ================================================================
	// §25 — Unchanged paragraph (no match sanity check)
	// ================================================================
	Heading(doc, "25. Unchanged Paragraph (Sanity Check)", 1)
	Note(doc, "No tags here. Must be byte-identical in both files.")
	Para(doc, "The quick brown fox jumps over the lazy dog. 0123456789. "+
		"No |< markers whatsoever — just pipe-angle-bracket fragments that don't form a valid tag.")

	// ================================================================
	// §26 — Tag in header table cell (case 9 + case 10)
	// ================================================================
	Heading(doc, "26. Tag In Header Table Cell", 1)
	SpecTable(doc, "|<HDR_CELL_TBL>|", "HC1,HC2")
	Note(doc, "Primary header has a 1×2 table. Right cell contains the tag. "+
		"After replacement: nested table inside header table cell.")

	// Force page break for page 3 content
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	Heading(doc, "— Page 4 (additional content to verify multi-page stability) —", 2)
	Para(doc, "If all tests above pass, the ReplaceWithTable implementation is working correctly. "+
		"Headers/footers should show their replacement tables on pages 2+ (primary) and page 1 (first-page).")

	return doc, nil
}
