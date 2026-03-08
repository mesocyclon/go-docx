// replace-content generates .docx files to visually verify ReplaceWithContent:
//
//	out/01_template.docx          — template with highlighted placeholders
//	out/sources/src_*.docx        — source documents (one per content type)
//	out/02_result.docx            — template reopened, all placeholders filled
//
// Three-phase design:
//
//	Phase 1 — Build and save the template document to disk.
//	Phase 2 — Build and save each source document to disk.
//	Phase 3 — Reopen template and every source from disk (full serialization
//	          roundtrip), apply ReplaceWithContent for each tag, save result.
//
// Run:
//
//	go run ./visual-regtest/replace-content --output ./visual-regtest/replace-content/out
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

func main() {
	outputDir := flag.String("output", "", "directory for generated .docx files")
	flag.Parse()
	if *outputDir == "" {
		log.Fatal("--output is required")
	}
	srcDir := filepath.Join(*outputDir, "sources")
	if err := os.MkdirAll(srcDir, 0o755); err != nil {
		log.Fatalf("creating source dir: %v", err)
	}

	var results []regtest.FileResult

	// ── Phase 1: build and save template ──────────────────────────────
	start := time.Now()
	tplDoc, err := buildTemplate()
	if err != nil {
		log.Fatalf("building template: %v", err)
	}
	tplPath := filepath.Join(*outputDir, "01_template.docx")
	if err := tplDoc.SaveFile(tplPath); err != nil {
		log.Fatalf("saving template: %v", err)
	}
	results = append(results, regtest.FileResult{
		Name: "01_template.docx", OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("OK  01_template.docx (%s)", time.Since(start))

	// ── Phase 2: build and save source documents ─────────────────────
	start = time.Now()
	srcs := allSources()
	for _, s := range srcs {
		doc, err := s.builder()
		if err != nil {
			log.Fatalf("building source %s: %v", s.filename, err)
		}
		p := filepath.Join(srcDir, s.filename)
		if err := doc.SaveFile(p); err != nil {
			log.Fatalf("saving source %s: %v", s.filename, err)
		}
		log.Printf("  src  %s", s.filename)
	}
	results = append(results, regtest.FileResult{
		Name: fmt.Sprintf("sources/ (%d files)", len(srcs)), OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("Phase 2: %d sources (%s)", len(srcs), time.Since(start))

	// ── Phase 3: reopen template + sources, fill, save ───────────────
	start = time.Now()
	tplBytes, err := os.ReadFile(tplPath)
	if err != nil {
		log.Fatalf("reading template: %v", err)
	}
	result, err := docx.OpenBytes(tplBytes)
	if err != nil {
		log.Fatalf("opening template: %v", err)
	}

	// Open each source from disk (full roundtrip).
	openedSources := make(map[string]*docx.Document, len(srcs))
	for _, s := range srcs {
		b, err := os.ReadFile(filepath.Join(srcDir, s.filename))
		if err != nil {
			log.Fatalf("reading source %s: %v", s.filename, err)
		}
		doc, err := docx.OpenBytes(b)
		if err != nil {
			log.Fatalf("opening source %s: %v", s.filename, err)
		}
		openedSources[s.filename] = doc
	}

	// Apply replacements.
	totalCount := 0
	for _, r := range allReplacements(openedSources) {
		n, err := result.ReplaceWithContent(r.tag, docx.ContentData{
			Source:  r.source,
			Format:  r.format,
			Options: r.opts,
		})
		if err != nil {
			log.Fatalf("ReplaceWithContent(%q): %v", r.tag, err)
		}
		log.Printf("  %-50s → %d", r.tag, n)
		totalCount += n
	}

	// Self-reference (source == target).
	nSelf, err := result.ReplaceWithContent("(<SELF>)", docx.ContentData{Source: result})
	if err != nil {
		log.Fatalf("ReplaceWithContent self-reference: %v", err)
	}
	log.Printf("  %-50s → %d", "(<SELF>)", nSelf)
	totalCount += nSelf

	// Empty search string — must return 0 without error.
	nEmpty, err := result.ReplaceWithContent("", docx.ContentData{Source: openedSources["src_paragraph.docx"]})
	if err != nil {
		log.Fatalf("ReplaceWithContent empty string: %v", err)
	}
	if nEmpty != 0 {
		log.Fatalf("empty search string returned %d, want 0", nEmpty)
	}
	log.Printf("  %-50s → %d (expected 0)", `""`, nEmpty)

	log.Printf("  total replacements: %d", totalCount)

	resultPath := filepath.Join(*outputDir, "02_result.docx")
	if err := result.SaveFile(resultPath); err != nil {
		log.Fatalf("saving result: %v", err)
	}
	results = append(results, regtest.FileResult{
		Name: "02_result.docx", OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("OK  02_result.docx (%s)", time.Since(start))

	// ── Manifest ─────────────────────────────────────────────────────
	if err := regtest.WriteManifest(filepath.Join(*outputDir, "manifest.json"), results); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}
	log.Printf("done: %d files, %d replacements", len(results), totalCount)
}
