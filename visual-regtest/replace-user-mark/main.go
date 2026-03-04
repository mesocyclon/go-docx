// replace-user-mark fills a user-provided template with user-provided mark
// documents. Tag names are derived from filenames in the marks directory.
//
// Directory layout:
//
//	replace-user-mark/
//	  in/
//	    template.docx              ← user-created template with (<tag>) placeholders
//	    mark/
//	      tag1.docx                ← source content for (<tag1>)
//	      tag2.docx                ← source content for (<tag2>)
//	      ...
//	  out/
//	    result.docx                ← filled template (generated)
//
// Mapping rule:  mark/NAME.docx  →  replaces every occurrence of  (<NAME>)
//
// Run:
//
//	go run ./visual-regtest/replace-user-mark --input ./in --output ./out
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
)

func main() {
	inputDir := flag.String("input", "visual-regtest/replace-user-mark/in", "directory with template.docx and mark/")
	outputDir := flag.String("output", "visual-regtest/replace-user-mark/out", "directory for output .docx files")
	flag.Parse()

	// ── Validate input ───────────────────────────────────────────────
	templatePath := filepath.Join(*inputDir, "template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		log.Fatalf("ERROR: template not found: %s\n\n"+
			"  Create the file and place (<tag>) placeholders inside.\n"+
			"  Then add mark documents to %s/mark/ with matching names.\n"+
			"  Example: mark/invoice_header.docx → replaces (<invoice_header>)\n",
			templatePath, *inputDir)
	}

	markDir := filepath.Join(*inputDir, "mark")
	if _, err := os.Stat(markDir); os.IsNotExist(err) {
		log.Fatalf("ERROR: mark directory not found: %s\n\n"+
			"  Create the directory and place mark_name.docx files inside.\n",
			markDir)
	}

	// ── Discover marks ───────────────────────────────────────────────
	entries, err := os.ReadDir(markDir)
	if err != nil {
		log.Fatalf("reading mark directory: %v", err)
	}

	type markEntry struct {
		tag      string // e.g. "(<invoice_header>)"
		name     string // e.g. "invoice_header"
		filename string // e.g. "invoice_header.docx"
		path     string // full path
	}

	var marks []markEntry
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		fname := e.Name()
		if !strings.HasSuffix(strings.ToLower(fname), ".docx") {
			continue
		}
		name := strings.TrimSuffix(fname, filepath.Ext(fname))
		marks = append(marks, markEntry{
			tag:      "(<" + name + ">)",
			name:     name,
			filename: fname,
			path:     filepath.Join(markDir, fname),
		})
	}

	if len(marks) == 0 {
		log.Fatalf("ERROR: no .docx files found in %s\n\n"+
			"  Add mark documents: mark/tag_name.docx → replaces (<tag_name>)\n",
			markDir)
	}

	log.Printf("template: %s", templatePath)
	log.Printf("marks:    %d file(s) in %s/", len(marks), markDir)
	for _, m := range marks {
		log.Printf("  %s  →  %s", m.filename, m.tag)
	}

	// ── Open template ────────────────────────────────────────────────
	start := time.Now()

	tplBytes, err := os.ReadFile(templatePath)
	if err != nil {
		log.Fatalf("reading template: %v", err)
	}
	result, err := docx.OpenBytes(tplBytes)
	if err != nil {
		log.Fatalf("opening template: %v", err)
	}

	// ── Open and apply each mark ─────────────────────────────────────
	totalCount := 0
	for _, m := range marks {
		markBytes, err := os.ReadFile(m.path)
		if err != nil {
			log.Fatalf("reading mark %s: %v", m.filename, err)
		}
		source, err := docx.OpenBytes(markBytes)
		if err != nil {
			log.Fatalf("opening mark %s: %v", m.filename, err)
		}
		n, err := result.ReplaceWithContent(m.tag, docx.ContentData{Source: source})
		if err != nil {
			log.Fatalf("ReplaceWithContent(%s): %v", m.tag, err)
		}
		log.Printf("  %-40s → %d replacement(s)", m.tag, n)
		totalCount += n
	}

	// ── Save result ──────────────────────────────────────────────────
	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}
	resultPath := filepath.Join(*outputDir, "result.docx")
	if err := result.SaveFile(resultPath); err != nil {
		log.Fatalf("saving result: %v", err)
	}

	elapsed := time.Since(start)
	log.Printf("done: %d mark(s), %d replacement(s), %s", len(marks), totalCount, elapsed)
	log.Printf("result: %s", resultPath)

	// ── Manifest ─────────────────────────────────────────────────────
	type ManifestEntry struct {
		Tag          string `json:"tag"`
		MarkFile     string `json:"mark_file"`
		Replacements int    `json:"replacements"`
	}
	type Manifest struct {
		Template      string          `json:"template"`
		Result        string          `json:"result"`
		Elapsed       string          `json:"elapsed"`
		TotalMarks    int             `json:"total_marks"`
		TotalReplaced int             `json:"total_replaced"`
		Marks         []ManifestEntry `json:"marks"`
	}

	manifest := Manifest{
		Template:      templatePath,
		Result:        resultPath,
		Elapsed:       elapsed.String(),
		TotalMarks:    len(marks),
		TotalReplaced: totalCount,
	}
	for _, m := range marks {
		manifest.Marks = append(manifest.Marks, ManifestEntry{
			Tag:      m.tag,
			MarkFile: m.filename,
		})
	}
	data, _ := json.MarshalIndent(manifest, "", "  ")
	manifestPath := filepath.Join(*outputDir, "manifest.json")
	if err := os.WriteFile(manifestPath, data, 0o644); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}

	// ── Summary ──────────────────────────────────────────────────────
	fmt.Println()
	fmt.Println("  ┌─────────────────────────────────────────────────┐")
	fmt.Printf("  │  Template:  %-37s│\n", filepath.Base(templatePath))
	fmt.Printf("  │  Marks:     %-37s│\n", fmt.Sprintf("%d file(s)", len(marks)))
	fmt.Printf("  │  Replaced:  %-37s│\n", fmt.Sprintf("%d occurrence(s)", totalCount))
	fmt.Printf("  │  Result:    %-37s│\n", resultPath)
	fmt.Println("  └─────────────────────────────────────────────────┘")
}
