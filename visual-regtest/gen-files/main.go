// gen-files generates .docx documents from scratch using the public API.
//
// Each gen* function produces one standalone document that exercises a
// specific area of the library. The output is written to --output.
//
// Run:  go run ./visual-regtest/gen-files --output ./visual-regtest/gen-files/out
// Or:   make gen-files   (from the visual-regtest directory)
package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// TestCase is one generated document.
type TestCase struct {
	Name string                         // output filename (without .docx)
	Gen  func() (*docx.Document, error) // generator
}

func main() {
	outputDir := flag.String("output", "", "directory for generated .docx files")
	flag.Parse()

	if *outputDir == "" {
		log.Fatal("--output is required")
	}
	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	tests := []TestCase{
		{"01_headings", genHeadings},
		{"02_paragraph_styles", genParagraphStyles},
		{"03_font_bold_italic_underline", genFontBasic},
		{"04_font_advanced", genFontAdvanced},
		{"05_font_color", genFontColor},
		{"06_font_size", genFontSize},
		{"07_paragraph_alignment", genParagraphAlignment},
		{"08_paragraph_indent", genParagraphIndent},
		{"09_paragraph_spacing", genParagraphSpacing},
		{"10_line_spacing", genLineSpacing},
		{"11_tab_stops", genTabStops},
		{"12_page_breaks", genPageBreaks},
		{"13_run_breaks", genRunBreaks},
		{"14_table_basic", genTableBasic},
		{"15_table_merged_cells", genTableMergedCells},
		{"16_table_alignment", genTableAlignment},
		{"17_table_nested", genTableNested},
		{"18_table_cell_valign", genTableCellVerticalAlign},
		{"19_table_add_row_col", genTableAddRowCol},
		{"20_sections_multi", genSectionsMulti},
		{"21_section_landscape", genSectionLandscape},
		{"22_section_margins", genSectionMargins},
		{"23_header_footer", genHeaderFooter},
		{"24_header_footer_first_page", genHeaderFooterFirstPage},
		{"25_comments", genComments},
		{"26_core_properties", genCoreProperties},
		{"27_custom_styles", genCustomStyles},
		{"28_mixed_content", genMixedContent},
		{"29_paragraph_format_flow", genParagraphFormatFlow},
		{"30_settings_odd_even", genSettingsOddEven},
		{"31_font_highlight", genFontHighlight},
		{"32_font_name", genFontName},
		{"33_underline_styles", genUnderlineStyles},
		{"34_table_row_height", genTableRowHeight},
		{"35_table_column_width", genTableColumnWidth},
		{"36_section_header_distance", genSectionHeaderDistance},
		{"37_insert_paragraph_before", genInsertParagraphBefore},
		{"38_inline_image", genInlineImage},
		{"39_multiple_runs", genMultipleRuns},
		{"40_paragraph_clear_set_text", genParagraphClearSetText},
		{"41_table_cell_set_text", genTableCellSetText},
		{"42_table_style", genTableStyle},
		{"43_table_bidi", genTableBidi},
		{"44_section_continuous_break", genSectionContinuousBreak},
		{"45_font_subscript_superscript", genFontSubSuperscript},
		{"46_tab_and_newline_in_text", genTabAndNewlineInText},
		{"47_large_document", genLargeDocument},
	}

	var results []regtest.FileResult
	for _, tc := range tests {
		start := time.Now()
		fname := tc.Name + ".docx"
		dstPath := filepath.Join(*outputDir, fname)

		doc, err := tc.Gen()
		if err != nil {
			results = append(results, regtest.FileResult{Name: fname, OK: false, Error: fmt.Sprintf("gen: %v", err), Elapsed: time.Since(start).String()})
			log.Printf("FAIL %s: %v", fname, err)
			continue
		}
		if err := doc.SaveFile(dstPath); err != nil {
			results = append(results, regtest.FileResult{Name: fname, OK: false, Error: fmt.Sprintf("save: %v", err), Elapsed: time.Since(start).String()})
			log.Printf("FAIL %s: save: %v", fname, err)
			continue
		}
		results = append(results, regtest.FileResult{Name: fname, OK: true, Elapsed: time.Since(start).String()})
		log.Printf("OK   %s (%s)", fname, time.Since(start))
	}

	manifestPath := filepath.Join(*outputDir, "manifest.json")
	if err := regtest.WriteManifest(manifestPath, results); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}

	okCount := 0
	for _, r := range results {
		if r.OK {
			okCount++
		}
	}
	log.Printf("done: %d/%d succeeded", okCount, len(results))
	if okCount != len(results) {
		os.Exit(1)
	}
}
