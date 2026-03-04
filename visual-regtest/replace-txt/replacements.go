package main

// repl is a simple old→new text replacement pair.
type repl struct{ old, new string }

// allReplacements returns every replacement pair applied to the "before" document.
func allReplacements() []repl {
	return []repl{
		// §1  simple placeholders
		{"{{NAME}}", "Иван Петров"},
		{"{{DATE}}", "January 15, 2025"},
		{"{{COMPANY}}", "Acme Corp"},

		// §2  cross-run (text split across differently-formatted runs)
		{"CROSSRUN_REPLACE", "DONE"},

		// §3  table cells
		{"CELL_OLD", "CELL_NEW"},

		// §4  nested table
		{"NESTED_OLD", "NESTED_NEW"},

		// §5  merged cells
		{"MERGED_CELL_TEXT", "MERGED_REPLACED"},

		// §6  header (primary)
		{"HEADER_PLACEHOLDER", "Real Header Title"},

		// §7  footer (primary)
		{"FOOTER_PLACEHOLDER", "Page 1 — Confidential"},

		// §8  first-page header/footer
		{"FIRST_HDR", "First Page Header — Replaced"},
		{"FIRST_FTR", "First Page Footer — Replaced"},

		// §9  header with table
		{"HDR_TBL_LEFT", "Company Name"},
		{"HDR_TBL_RIGHT", "Doc #12345"},

		// §10 footer with table
		{"FTR_TBL_LEFT", "Legal Notice"},
		{"FTR_TBL_RIGHT", "Page X of Y"},

		// §11  multiple occurrences
		{"MULTI", "REPLACED"},

		// §12 tab inside search string
		{"COL_A\tCOL_B", "MERGED_AB"},

		// §13 newline inside search string
		{"LINE_ONE\nLINE_TWO", "SINGLE_LINE"},

		// §14 deletion (replace with empty)
		{"[DELETE_ME]", ""},

		// §15 short → long expansion
		{"TINY", "THIS_IS_MUCH_LONGER_THAN_BEFORE"},

		// §16 long → short contraction
		{"VERY_LONG_PLACEHOLDER_TEXT_HERE", "Short"},

		// §17 Cyrillic single-run
		{"ЗАМЕНИТЬ", "ГОТОВО"},
		{"Шаблон", "Результат"},

		// §18 cross-run Cyrillic
		{"КРОССРАН", "OK"},

		// §19 replacement at paragraph start
		{"STARTWORD", "REPLACED_START"},

		// §20 replacement at paragraph end
		{"ENDWORD", "REPLACED_END"},

		// §21 no-op: old == ""
		{"", "should_never_appear"},

		// §22 no-op: old == new
		{"NOOP_SAME", "NOOP_SAME"},

		// §23 comment-annotated text
		{"COMMENTED_TEXT", "COMMENT_REPLACED"},

		// §24 table: full row replacement
		{"ROW_NAME", "Alice Johnson"},
		{"ROW_ROLE", "Lead Engineer"},
		{"ROW_DEPT", "Platform"},

		// §25 table: header row replacement
		{"TH_COL1", "Employee"},
		{"TH_COL2", "Department"},
		{"TH_COL3", "Status"},

		// §29 replacement inside comment body
		{"COMMENT_BODY_OLD", "COMMENT_BODY_NEW"},
	}
}
