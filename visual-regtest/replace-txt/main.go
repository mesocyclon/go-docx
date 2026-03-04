// replace-txt generates a pair of .docx files to visually verify ReplaceText:
//
//	01_before_replace.docx — original document with highlighted placeholders and a spec for each test
//	02_after_replace.docx  — same document reopened, replacements applied, re-saved
//
// Visual verification:
//
//	BEFORE:  Yellow-highlighted text = placeholders that WILL be replaced.
//	         Green-highlighted text  = the expected replacement value (shown in the spec line).
//	AFTER:   Yellow-highlighted text now contains the replacement values.
//	         Compare yellow (actual) vs green (expected) in each section — they must match.
//
// The "before" file is created from scratch, saved to disk, then reopened via
// docx.OpenBytes (full serialization roundtrip) before applying replacements.
// This ensures the test exercises the real read→modify→write pipeline.
//
// Run:
//
//	go run ./visual-regtest/replace-txt --output ./visual-regtest/replace-txt/out
package main

import (
	"flag"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

func main() {
	outputDir := flag.String("output", "", "directory for generated .docx files")
	flag.Parse()

	if *outputDir == "" {
		log.Fatal("--output is required")
	}
	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	var results []regtest.FileResult

	// ---- Step 1: build and save the "before" document ----
	start := time.Now()
	beforeDoc, err := buildBeforeDocument()
	if err != nil {
		log.Fatalf("building before document: %v", err)
	}
	beforePath := filepath.Join(*outputDir, "01_before_replace.docx")
	if err := beforeDoc.SaveFile(beforePath); err != nil {
		log.Fatalf("saving before document: %v", err)
	}
	results = append(results, regtest.FileResult{
		Name: "01_before_replace.docx", OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("OK   01_before_replace.docx (%s)", time.Since(start))

	// ---- Step 2: reopen saved file and apply replacements ----
	start = time.Now()
	beforeBytes, err := os.ReadFile(beforePath)
	if err != nil {
		log.Fatalf("reading before file: %v", err)
	}
	afterDoc, err := docx.OpenBytes(beforeBytes)
	if err != nil {
		log.Fatalf("opening before file: %v", err)
	}

	totalCount := 0
	for _, r := range allReplacements() {
		n, err := afterDoc.ReplaceText(r.old, r.new)
		if err != nil {
			log.Fatalf("ReplaceText(%q, %q): %v", r.old, r.new, err)
		}
		log.Printf("  %-40s → %-30s  %d hits", regtest.TruncQuote(r.old), regtest.TruncQuote(r.new), n)
		totalCount += n
	}
	log.Printf("  total replacements: %d", totalCount)

	afterPath := filepath.Join(*outputDir, "02_after_replace.docx")
	if err := afterDoc.SaveFile(afterPath); err != nil {
		log.Fatalf("saving after document: %v", err)
	}
	results = append(results, regtest.FileResult{
		Name: "02_after_replace.docx", OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("OK   02_after_replace.docx (%s)", time.Since(start))

	// ---- Manifest ----
	if err := regtest.WriteManifest(filepath.Join(*outputDir, "manifest.json"), results); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}
	log.Printf("done: %d files, %d total replacements", len(results), totalCount)
}
