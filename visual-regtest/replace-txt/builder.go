package main

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx"
	. "github.com/vortex/go-docx/visual-regtest/internal/docfmt"
	"github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// buildBeforeDocument creates the "before" test document from scratch.
func buildBeforeDocument() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	// ---- headers / footers ----
	sect, err := doc.Sections().Get(0)
	if err != nil {
		return nil, fmt.Errorf("getting default section: %w", err)
	}

	// Primary header: highlighted text + table with placeholders
	hdr := sect.Header()
	BuildHighlightedParagraph(hdr, "HEADER_PLACEHOLDER")
	hdrTbl, err := hdr.AddTable(1, 2, 9000)
	if err != nil {
		return nil, err
	}
	hc0, _ := hdrTbl.CellAt(0, 0)
	hc0.SetText("HDR_TBL_LEFT")
	hc1, _ := hdrTbl.CellAt(0, 1)
	hc1.SetText("HDR_TBL_RIGHT")

	// Primary footer: highlighted text + table with placeholders
	ftr := sect.Footer()
	BuildHighlightedParagraph(ftr, "FOOTER_PLACEHOLDER")
	ftrTbl, err := ftr.AddTable(1, 2, 9000)
	if err != nil {
		return nil, err
	}
	fc0, _ := ftrTbl.CellAt(0, 0)
	fc0.SetText("FTR_TBL_LEFT")
	fc1, _ := ftrTbl.CellAt(0, 1)
	fc1.SetText("FTR_TBL_RIGHT")

	// First-page header/footer
	if err := sect.SetDifferentFirstPageHeaderFooter(true); err != nil {
		return nil, err
	}
	BuildHighlightedParagraph(sect.FirstPageHeader(), "FIRST_HDR")
	BuildHighlightedParagraph(sect.FirstPageFooter(), "FIRST_FTR")

	// ================================================================
	// Legend
	// ================================================================
	Heading(doc, "ReplaceText — Visual Regression Test", 0)

	lp1, _ := doc.AddParagraph("")
	AddPlain(lp1, "How to read this document:  ")
	r, _ := lp1.AddRun("Yellow highlight")
	_ = r.SetBold(regtest.BoolPtr(true))
	SetHighlightYellow(r)
	AddPlain(lp1, " = placeholder (will be replaced).  ")
	r, _ = lp1.AddRun("Green highlight")
	_ = r.SetBold(regtest.BoolPtr(true))
	SetHighlightGreen(r)
	AddPlain(lp1, " = expected result.")

	lp2, _ := doc.AddParagraph("")
	AddPlain(lp2, "In ")
	AddBold(lp2, "02_after_replace.docx")
	AddPlain(lp2, ": yellow text should now contain the expected value. "+
		"Compare yellow (actual) vs green (expected) in each spec line — they must match.")

	Para(doc, "")

	// ================================================================
	// §1 — Simple placeholder replacement
	// ================================================================
	Heading(doc, "1. Simple Placeholder Replacement", 1)
	SpecReplace(doc, "{{NAME}}", "Иван Петров")
	SpecReplace(doc, "{{DATE}}", "January 15, 2025")
	SpecReplace(doc, "{{COMPANY}}", "Acme Corp")

	TagPara(doc, "Name: ", "{{NAME}}", "")
	TagPara(doc, "Date: ", "{{DATE}}", "")
	TagPara(doc, "Company: ", "{{COMPANY}}", "")

	// ================================================================
	// §2 — Cross-run replacement (formatting preserved)
	// ================================================================
	Heading(doc, "2. Cross-Run Replacement (Formatting Preserved)", 1)
	SpecReplace(doc, "CROSSRUN_REPLACE", "DONE")
	Note(doc, "Placeholder split across 3 runs: 'CROSS' (bold red) + 'RUN_RE' (italic blue) + 'PLACE' (normal). "+
		"Result appears in first run. Middle/last runs become empty but keep their rPr.")

	p2, _ := doc.AddParagraph("")
	AddPlain(p2, "Before marker: ")
	cr1, _ := p2.AddRun("CROSS")
	_ = cr1.SetBold(regtest.BoolPtr(true))
	c1 := docx.NewRGBColor(0xFF, 0, 0)
	_ = cr1.Font().Color().SetRGB(&c1)
	SetHighlightYellow(cr1)
	cr2, _ := p2.AddRun("RUN_RE")
	_ = cr2.SetItalic(regtest.BoolPtr(true))
	c2 := docx.NewRGBColor(0, 0, 0xFF)
	_ = cr2.Font().Color().SetRGB(&c2)
	SetHighlightYellow(cr2)
	cr3, _ := p2.AddRun("PLACE")
	SetHighlightYellow(cr3)
	AddPlain(p2, " — after marker")

	// ================================================================
	// §3 — Table cell replacement
	// ================================================================
	Heading(doc, "3. Table Cell Replacement", 1)
	SpecReplace(doc, "CELL_OLD", "CELL_NEW")
	Note(doc, "3×3 table. 'CELL_OLD' appears in 3 cells. Other cells untouched.")

	tbl3, _ := doc.AddTable(3, 3)
	FillTable(tbl3, [][]string{
		{"Header A", "Header B", "Header C"},
		{"CELL_OLD value 1", "Normal text", "CELL_OLD value 2"},
		{"Row 3 Col 1", "CELL_OLD value 3", "Row 3 Col 3"},
	})

	// ================================================================
	// §4 — Nested table
	// ================================================================
	Heading(doc, "4. Nested Table Replacement", 1)
	SpecReplace(doc, "NESTED_OLD", "NESTED_NEW")
	Note(doc, "Table inside a cell. 'NESTED_OLD' in inner table rows. Outer cell untouched.")

	outer, _ := doc.AddTable(1, 2)
	oc0, _ := outer.CellAt(0, 0)
	oc0.SetText("Outer cell — no replacement")
	oc1, _ := outer.CellAt(0, 1)
	inner, _ := oc1.AddTable(2, 1)
	ic0, _ := inner.CellAt(0, 0)
	ic0.SetText("NESTED_OLD — row 1")
	ic1, _ := inner.CellAt(1, 0)
	ic1.SetText("NESTED_OLD — row 2")

	// ================================================================
	// §5 — Merged cells
	// ================================================================
	Heading(doc, "5. Merged Cell Replacement", 1)
	SpecReplace(doc, "MERGED_CELL_TEXT", "MERGED_REPLACED")
	Note(doc, "A1+B1 merged horizontally. Placeholder replaced exactly once, not duplicated.")

	tbl5, _ := doc.AddTable(2, 3)
	a1, _ := tbl5.CellAt(0, 0)
	b1, _ := tbl5.CellAt(0, 1)
	merged, _ := a1.Merge(b1)
	merged.SetText("MERGED_CELL_TEXT — spans two columns")
	cc, _ := tbl5.CellAt(0, 2)
	cc.SetText("Normal C1")
	for c := 0; c < 3; c++ {
		cl, _ := tbl5.CellAt(1, c)
		cl.SetText(fmt.Sprintf("Row 2, Col %d", c+1))
	}

	// ================================================================
	// §6 — Header replacement (text)
	// ================================================================
	Heading(doc, "6. Header Replacement (Text)", 1)
	SpecReplace(doc, "HEADER_PLACEHOLDER", "Real Header Title")
	Note(doc, "Primary page header contains yellow-highlighted placeholder. Visible on page 2+.")

	// ================================================================
	// §7 — Footer replacement (text)
	// ================================================================
	Heading(doc, "7. Footer Replacement (Text)", 1)
	SpecReplace(doc, "FOOTER_PLACEHOLDER", "Page 1 — Confidential")
	Note(doc, "Primary page footer contains yellow-highlighted placeholder. Visible on page 2+.")

	// ================================================================
	// §8 — First-page header/footer
	// ================================================================
	Heading(doc, "8. First-Page Header/Footer", 1)
	SpecReplace(doc, "FIRST_HDR", "First Page Header — Replaced")
	SpecReplace(doc, "FIRST_FTR", "First Page Footer — Replaced")
	Note(doc, "First page has separate header/footer. Primary header/footer visible on page 2 onward.")

	// ================================================================
	// §9 — Header table replacement
	// ================================================================
	Heading(doc, "9. Header Table Replacement", 1)
	SpecReplace(doc, "HDR_TBL_LEFT", "Company Name")
	SpecReplace(doc, "HDR_TBL_RIGHT", "Doc #12345")
	Note(doc, "Primary header has a 1×2 table. Both cells have placeholders. "+
		"Verifies ReplaceText reaches tables inside headers.")

	// ================================================================
	// §10 — Footer table replacement
	// ================================================================
	Heading(doc, "10. Footer Table Replacement", 1)
	SpecReplace(doc, "FTR_TBL_LEFT", "Legal Notice")
	SpecReplace(doc, "FTR_TBL_RIGHT", "Page X of Y")
	Note(doc, "Primary footer has a 1×2 table. Both cells have placeholders. "+
		"Verifies ReplaceText reaches tables inside footers.")

	// Force page 2 so primary header/footer become visible.
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	Heading(doc, "— Page 2 (primary header/footer visible here) —", 2)

	// ================================================================
	// §11 — Multiple occurrences
	// ================================================================
	Heading(doc, "11. Multiple Occurrences", 1)
	SpecReplace(doc, "MULTI", "REPLACED")
	Note(doc, "4 occurrences across 2 paragraphs — all must be replaced.")

	p11a, _ := doc.AddParagraph("")
	AddHighlighted(p11a, "MULTI")
	AddPlain(p11a, " is here, and ")
	AddHighlighted(p11a, "MULTI")
	AddPlain(p11a, " is there, and ")
	AddHighlighted(p11a, "MULTI")
	AddPlain(p11a, " is everywhere.")

	TagPara(doc, "Another paragraph also has ", "MULTI", " in it.")

	// ================================================================
	// §12 — Tab inside search string
	// ================================================================
	Heading(doc, "12. Tab Inside Search String", 1)
	SpecReplace(doc, "COL_A⟶COL_B (⟶ = tab)", "MERGED_AB")
	Note(doc, "The <w:tab> between A and B is consumed by the match. Trailing tab+COL_C remains.")
	Para(doc, "COL_A\tCOL_B\tCOL_C")

	// ================================================================
	// §13 — Newline inside search string
	// ================================================================
	Heading(doc, "13. Newline Inside Search String", 1)
	SpecReplace(doc, "LINE_ONE⏎LINE_TWO (⏎ = br)", "SINGLE_LINE")
	Note(doc, "The <w:br> is consumed. Trailing br+LINE_THREE remains.")
	Para(doc, "LINE_ONE\nLINE_TWO\nLINE_THREE")

	// ================================================================
	// §14 — Deletion
	// ================================================================
	Heading(doc, "14. Deletion (Replace with Empty String)", 1)
	SpecReplace(doc, "[DELETE_ME]", "⟨empty⟩")
	Note(doc, "Placeholder disappears completely. Surrounding text stays. "+
		"Yellow highlight vanishes from result.")

	TagPara(doc, "Before: ", "[DELETE_ME]", " — bracket text should vanish.")
	TagPara(doc, "Also here: ", "[DELETE_ME]", " gone.")

	// ================================================================
	// §15 — Short → long
	// ================================================================
	Heading(doc, "15. Short to Long Expansion", 1)
	SpecReplace(doc, "TINY", "THIS_IS_MUCH_LONGER_THAN_BEFORE")

	TagPara(doc, "Expand this: ", "TINY", " — done.")

	// ================================================================
	// §16 — Long → short
	// ================================================================
	Heading(doc, "16. Long to Short Contraction", 1)
	SpecReplace(doc, "VERY_LONG_PLACEHOLDER_TEXT_HERE", "Short")

	TagPara(doc, "Contract this: ", "VERY_LONG_PLACEHOLDER_TEXT_HERE", " — done.")

	// ================================================================
	// §17 — Cyrillic / UTF-8
	// ================================================================
	Heading(doc, "17. Cyrillic / UTF-8 (Single Run)", 1)
	SpecReplace(doc, "ЗАМЕНИТЬ", "ГОТОВО")
	SpecReplace(doc, "Шаблон", "Результат")
	Note(doc, "Multibyte byte offsets must be handled correctly.")

	TagPara(doc, "Нужно ", "ЗАМЕНИТЬ", " это слово.")
	TagPara(doc, "", "Шаблон", " документа — версия 1.0")

	// ================================================================
	// §18 — Cross-run Cyrillic
	// ================================================================
	Heading(doc, "18. Cross-Run Cyrillic", 1)
	SpecReplace(doc, "КРОССРАН", "OK")
	Note(doc, "Split: 'КРОСС' (bold) + 'РАН' (italic). Both runs keep formatting.")

	p18, _ := doc.AddParagraph("")
	AddPlain(p18, "Before: ")
	kr1, _ := p18.AddRun("КРОСС")
	_ = kr1.SetBold(regtest.BoolPtr(true))
	SetHighlightYellow(kr1)
	kr2, _ := p18.AddRun("РАН")
	_ = kr2.SetItalic(regtest.BoolPtr(true))
	SetHighlightYellow(kr2)
	AddPlain(p18, " — кросс-рановая кириллица")

	// ================================================================
	// §19 — Replacement at paragraph start
	// ================================================================
	Heading(doc, "19. Replacement at Paragraph Start", 1)
	SpecReplace(doc, "STARTWORD", "REPLACED_START")

	TagPara(doc, "", "STARTWORD", " is at the very beginning.")

	// ================================================================
	// §20 — Replacement at paragraph end
	// ================================================================
	Heading(doc, "20. Replacement at Paragraph End", 1)
	SpecReplace(doc, "ENDWORD", "REPLACED_END")

	TagPara(doc, "This paragraph ends with ", "ENDWORD", "")

	// ================================================================
	// §21 — No-op: old == ""
	// ================================================================
	Heading(doc, "21. No-Op: Empty Search String", 1)
	Note(doc, "ReplaceText(\"\", ...) returns 0. 'should_never_appear' must NOT appear. "+
		"Paragraph must be identical in both files.")
	Para(doc, "Nothing should change here. Absolutely identical in both files.")

	// ================================================================
	// §22 — No-op: old == new
	// ================================================================
	Heading(doc, "22. No-Op: old == new", 1)
	SpecReplace(doc, "NOOP_SAME", "NOOP_SAME")
	Note(doc, "Returns 0. No XML modification. Text identical in both files.")

	TagPara(doc, "The word ", "NOOP_SAME", " should remain as-is in both files.")

	// ================================================================
	// §23 — Comment-annotated text
	// ================================================================
	Heading(doc, "23. Comment-Annotated Text", 1)
	SpecReplace(doc, "COMMENTED_TEXT", "COMMENT_REPLACED")
	Note(doc, "A comment is attached to the placeholder. After replacement "+
		"comment markers and the comment body must survive.")

	p23, _ := doc.AddParagraph("")
	AddPlain(p23, "Commented word: ")
	p23run, _ := p23.AddRun("COMMENTED_TEXT")
	SetHighlightYellow(p23run)
	initials := "VR"
	if _, err := doc.AddComment([]*docx.Run{p23run}, "This is a test comment — must survive replacement", "Visual Regtest", &initials); err != nil {
		return nil, fmt.Errorf("adding comment: %w", err)
	}

	// ================================================================
	// §24 — Table: data row replacement
	// ================================================================
	Heading(doc, "24. Table Data Row Replacement", 1)
	SpecReplace(doc, "ROW_NAME", "Alice Johnson")
	SpecReplace(doc, "ROW_ROLE", "Lead Engineer")
	SpecReplace(doc, "ROW_DEPT", "Platform")
	Note(doc, "3×3 table with header row. Second row has 3 placeholders. "+
		"Header and third row untouched.")

	tbl24, _ := doc.AddTable(3, 3)
	FillTable(tbl24, [][]string{
		{"Name", "Role", "Department"},
		{"ROW_NAME", "ROW_ROLE", "ROW_DEPT"},
		{"Bob Smith", "Designer", "Creative"},
	})

	// ================================================================
	// §25 — Table: header row replacement
	// ================================================================
	Heading(doc, "25. Table Header Row Replacement", 1)
	SpecReplace(doc, "TH_COL1", "Employee")
	SpecReplace(doc, "TH_COL2", "Department")
	SpecReplace(doc, "TH_COL3", "Status")
	Note(doc, "Table header row cells contain placeholders. Data rows untouched.")

	tbl25, _ := doc.AddTable(3, 3)
	FillTable(tbl25, [][]string{
		{"TH_COL1", "TH_COL2", "TH_COL3"},
		{"John", "Engineering", "Active"},
		{"Jane", "Marketing", "On Leave"},
	})

	// ================================================================
	// §26 — Multiple different placeholders in one paragraph
	// ================================================================
	Heading(doc, "26. Multiple Different Placeholders in One Paragraph", 1)
	SpecReplace(doc, "{{NAME}}", "Иван Петров")
	SpecReplace(doc, "{{COMPANY}}", "Acme Corp")
	SpecReplace(doc, "{{DATE}}", "January 15, 2025")
	Note(doc, "Three separate ReplaceText calls each hit this paragraph. "+
		"All replaced, surrounding text intact.")

	p26, _ := doc.AddParagraph("")
	AddPlain(p26, "Dear ")
	AddHighlighted(p26, "{{NAME}}")
	AddPlain(p26, ", your order from ")
	AddHighlighted(p26, "{{COMPANY}}")
	AddPlain(p26, " on ")
	AddHighlighted(p26, "{{DATE}}")
	AddPlain(p26, " is confirmed.")

	// ================================================================
	// §27 — Unchanged paragraph (no match)
	// ================================================================
	Heading(doc, "27. Unchanged Paragraph (No Match)", 1)
	Note(doc, "No placeholders here. Must be byte-identical in both files.")
	Para(doc, "The quick brown fox jumps over the lazy dog. 0123456789. No markers whatsoever.")

	// ================================================================
	// §28 — Replacement inside one formatted run, sibling runs untouched
	// ================================================================
	Heading(doc, "28. Replacement Inside One Formatted Run", 1)
	SpecReplace(doc, "CELL_OLD", "CELL_NEW")
	Note(doc, "'CELL_OLD' sits inside a bold run followed by italic text. "+
		"After replacement bold stays bold, italic stays italic.")

	p28, _ := doc.AddParagraph("")
	fr1, _ := p28.AddRun("Bold CELL_OLD text")
	_ = fr1.SetBold(regtest.BoolPtr(true))
	SetHighlightYellow(fr1)
	fr2, _ := p28.AddRun(" then italic text")
	_ = fr2.SetItalic(regtest.BoolPtr(true))

	// ================================================================
	// §29 — Replacement inside comment body
	// ================================================================
	Heading(doc, "29. Replacement Inside Comment Body", 1)
	SpecReplace(doc, "COMMENT_BODY_OLD", "COMMENT_BODY_NEW")
	Note(doc, "The comment body itself contains a placeholder. ReplaceText must reach "+
		"into word/comments.xml and replace it. The annotated body text stays unchanged.")

	p29, _ := doc.AddParagraph("")
	AddPlain(p29, "This text has a comment whose body contains a placeholder: ")
	p29run, _ := p29.AddRun("see comment")
	SetHighlightYellow(p29run)
	initials29 := "CB"
	if _, err := doc.AddComment(
		[]*docx.Run{p29run},
		"Review status: COMMENT_BODY_OLD — needs update",
		"Comment Body Test", &initials29,
	); err != nil {
		return nil, fmt.Errorf("adding comment for §29: %w", err)
	}

	return doc, nil
}
