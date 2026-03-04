package main

import (
	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// 12 — Page breaks
func genPageBreaks() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Page 1", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content on page 1."); err != nil {
		return nil, err
	}

	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Page 2 (after AddPageBreak)", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content on page 2."); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("Page 3 — via PageBreakBefore property. This paragraph should start on a new page.")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetPageBreakBefore(BoolPtr(true)); err != nil {
		return nil, err
	}

	return doc, nil
}

// 13 — Run-level breaks
func genRunBreaks() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Run-Level Breaks", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	r1, err := p1.AddRun("Before line break")
	if err != nil {
		return nil, err
	}
	if err := r1.AddBreak(enum.WdBreakTypeLine); err != nil {
		return nil, err
	}
	r1.AddText("After line break (same paragraph)")

	p2, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	r2, err := p2.AddRun("Before column break")
	if err != nil {
		return nil, err
	}
	if err := r2.AddBreak(enum.WdBreakTypeColumn); err != nil {
		return nil, err
	}
	r2.AddText("After column break")

	p3, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	r3, err := p3.AddRun("Before page break (in run)")
	if err != nil {
		return nil, err
	}
	if err := r3.AddBreak(enum.WdBreakTypePage); err != nil {
		return nil, err
	}
	r3.AddText("After page break (same run, new page)")

	return doc, nil
}

// 20 — Multiple sections
func genSectionsMulti() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Section 1", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content in section 1."); err != nil {
		return nil, err
	}

	if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Section 2 (New Page)", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content in section 2."); err != nil {
		return nil, err
	}

	if _, err := doc.AddSection(enum.WdSectionStartOddPage); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Section 3 (Odd Page)", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content in section 3."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 21 — Landscape section
func genSectionLandscape() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	if sections.Len() > 0 {
		sect, _ := sections.Get(sections.Len() - 1)
		if err := sect.SetOrientation(enum.WdOrientationLandscape); err != nil {
			return nil, err
		}
		if err := sect.SetPageWidth(IntPtr(docx.Inches(11).Twips())); err != nil {
			return nil, err
		}
		if err := sect.SetPageHeight(IntPtr(docx.Inches(8.5).Twips())); err != nil {
			return nil, err
		}
	}

	if _, err := doc.AddHeading("Landscape Page", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This entire document is in landscape orientation."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 22 — Custom margins
func genSectionMargins() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)
	if err := sect.SetTopMargin(IntPtr(docx.Inches(2).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetBottomMargin(IntPtr(docx.Inches(2).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetLeftMargin(IntPtr(docx.Inches(1.5).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetRightMargin(IntPtr(docx.Inches(1.5).Twips())); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Custom Margins", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Top=2\", Bottom=2\", Left=1.5\", Right=1.5\""); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 23 — Header and Footer
func genHeaderFooter() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)

	header := sect.Header()
	if _, err := header.AddParagraph("Header: Document Title — Confidential"); err != nil {
		return nil, err
	}

	footer := sect.Footer()
	if _, err := footer.AddParagraph("Footer: Page X of Y — © 2025 Company"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Document with Header and Footer", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Check the header and footer areas of this document."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 24 — First-page header/footer
func genHeaderFooterFirstPage() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)
	if err := sect.SetDifferentFirstPageHeaderFooter(true); err != nil {
		return nil, err
	}

	header := sect.Header()
	if _, err := header.AddParagraph("Primary Header (pages 2+)"); err != nil {
		return nil, err
	}
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Primary Footer (pages 2+)"); err != nil {
		return nil, err
	}

	firstHeader := sect.FirstPageHeader()
	if _, err := firstHeader.AddParagraph("FIRST PAGE HEADER"); err != nil {
		return nil, err
	}
	firstFooter := sect.FirstPageFooter()
	if _, err := firstFooter.AddParagraph("FIRST PAGE FOOTER"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("First Page — Different Header/Footer", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This is the first page."); err != nil {
		return nil, err
	}
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Second Page", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This is the second page with the primary header/footer."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 30 — Settings: odd/even page header/footer
func genSettingsOddEven() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	settings, err := doc.Settings()
	if err != nil {
		return nil, err
	}
	if err := settings.SetOddAndEvenPagesHeaderFooter(true); err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)

	header := sect.Header()
	if _, err := header.AddParagraph("Odd Page Header"); err != nil {
		return nil, err
	}
	evenHeader := sect.EvenPageHeader()
	if _, err := evenHeader.AddParagraph("Even Page Header"); err != nil {
		return nil, err
	}
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Odd Page Footer"); err != nil {
		return nil, err
	}
	evenFooter := sect.EvenPageFooter()
	if _, err := evenFooter.AddParagraph("Even Page Footer"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Odd/Even Headers/Footers", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Page 1 (odd)"); err != nil {
		return nil, err
	}
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Page 2 (even)"); err != nil {
		return nil, err
	}
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Page 3 (odd)"); err != nil {
		return nil, err
	}

	return doc, nil
}

// 36 — Section header/footer distance
func genSectionHeaderDistance() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)
	if err := sect.SetHeaderDistance(IntPtr(docx.Inches(0.3).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetFooterDistance(IntPtr(docx.Inches(0.3).Twips())); err != nil {
		return nil, err
	}

	header := sect.Header()
	if _, err := header.AddParagraph("Header close to edge (0.3\")"); err != nil {
		return nil, err
	}
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Footer close to edge (0.3\")"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Header/Footer Distance", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Header and footer are 0.3 inches from the edge."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 44 — Section with continuous break
func genSectionContinuousBreak() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Continuous Section Break", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content before continuous break."); err != nil {
		return nil, err
	}

	if _, err := doc.AddSection(enum.WdSectionStartContinuous); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content after continuous section break (same page)."); err != nil {
		return nil, err
	}

	if _, err := doc.AddSection(enum.WdSectionStartEvenPage); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content after even-page section break."); err != nil {
		return nil, err
	}

	return doc, nil
}
