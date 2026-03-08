// replace-user-mark-batch fills a user-provided template once per mark
// document, producing a separate output file for each mark.
//
// Directory layout:
//
//	in/
//	  template.docx           ← template with (<mark>) placeholders
//	  mark/
//	    alpha.docx            ← content to substitute
//	    beta.docx
//	    ...
//	out/
//	    alpha.docx            ← filled template (name = mark name)
//	    beta.docx
//	    manifest.json         ← per-file result [{name, ok, error, elapsed}]
//
// For each mark/NAME.docx a fresh copy of the template is opened,
// every occurrence of the tag is replaced, and the result is saved
// as out/NAME.docx.
//
// Usage:
//
//	go run ./visual-regtest/replace-user-mark-batch \
//	    --input  visual-regtest/replace-user-mark-batch/in \
//	    --output visual-regtest/replace-user-mark-batch/out \
//	    --tag '(<mark>)' --workers 4
package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

func main() {
	inputDir := flag.String("input", "visual-regtest/replace-user-mark-batch/in", "directory with template.docx and mark/")
	outputDir := flag.String("output", "visual-regtest/replace-user-mark-batch/out", "directory for output .docx files")
	tag := flag.String("tag", "(<mark>)", "placeholder tag to replace in template")
	workers := flag.Int("workers", 4, "parallel workers")
	flag.Parse()

	if *workers < 1 {
		*workers = 1
	}

	templatePath := filepath.Join(*inputDir, "template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		log.Fatalf("template not found: %s", templatePath)
	}

	markDir := filepath.Join(*inputDir, "mark")
	if _, err := os.Stat(markDir); os.IsNotExist(err) {
		log.Fatalf("mark directory not found: %s", markDir)
	}

	entries, err := os.ReadDir(markDir)
	if err != nil {
		log.Fatalf("reading mark directory: %v", err)
	}

	type markEntry struct {
		name, filename, path string
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
			name:     name,
			filename: fname,
			path:     filepath.Join(markDir, fname),
		})
	}

	if len(marks) == 0 {
		log.Fatalf("no .docx files found in %s", markDir)
	}

	tplBytes, err := os.ReadFile(templatePath)
	if err != nil {
		log.Fatalf("reading template: %v", err)
	}

	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	log.Printf("template: %s", templatePath)
	log.Printf("tag:      %s", *tag)
	log.Printf("marks:    %d file(s) × %d variant(s)", len(marks), len(allVariants()))
	log.Printf("workers:  %d", *workers)

	jobs := make(chan markEntry, len(marks))
	for _, m := range marks {
		jobs <- m
	}
	close(jobs)

	var (
		mu      sync.Mutex
		results []regtest.FileResult
	)

	var wg sync.WaitGroup
	for i := 0; i < *workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for m := range jobs {
				for _, v := range allVariants() {
					r := processOneMark(tplBytes, m.name, m.filename, m.path, *tag, *outputDir, v)
					mu.Lock()
					results = append(results, r)
					mu.Unlock()
					if r.OK {
						log.Printf("  OK   %-40s %d replacement(s)  %s", r.Name, r.Replacements, r.Elapsed)
					} else {
						log.Printf("  FAIL %-40s %s", r.Name, r.Error)
					}
				}
			}
		}()
	}
	wg.Wait()

	if err := regtest.WriteManifest(filepath.Join(*outputDir, "manifest.json"), results); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}

	okCount := 0
	for _, r := range results {
		if r.OK {
			okCount++
		}
	}
	log.Printf("done: %d/%d succeeded", okCount, len(results))
}

// importVariant describes one ImportFormatMode + ImportFormatOptions combination
// to exercise during batch processing. Each mark file is processed once per
// variant, producing a separate output file with the variant's suffix.
type importVariant struct {
	suffix string
	mode   docx.ImportFormatMode
	opts   docx.ImportFormatOptions
}

func allVariants() []importVariant {
	return []importVariant{
		// Default: target styles win on conflict.
		{"", docx.UseDestinationStyles, docx.ImportFormatOptions{}},

		// KeepSourceFormatting: expand source style props into direct attrs.
		{"__keep_src", docx.KeepSourceFormatting, docx.ImportFormatOptions{}},

		// KeepSourceFormatting + ForceCopyStyles: copy conflicting styles
		// with unique suffix (_0, _1, ...) instead of expanding.
		{"__keep_src_force", docx.KeepSourceFormatting, docx.ImportFormatOptions{ForceCopyStyles: true}},

		// KeepDifferentStyles: hybrid — identical formatting → use target,
		// different formatting → expand to direct attrs.
		{"__keep_diff", docx.KeepDifferentStyles, docx.ImportFormatOptions{}},

		// KeepDifferentStyles + ForceCopyStyles: different → copy with suffix.
		{"__keep_diff_force", docx.KeepDifferentStyles, docx.ImportFormatOptions{ForceCopyStyles: true}},

		// KeepSourceNumbering: preserve source list numbering as separate
		// definitions instead of merging into matching target lists.
		{"__keep_num", docx.UseDestinationStyles, docx.ImportFormatOptions{KeepSourceNumbering: true}},
	}
}

func processOneMark(tplBytes []byte, markName, markFilename, markPath, tag, outputDir string, v importVariant) (result regtest.FileResult) {
	start := time.Now()
	outName := markName + v.suffix + ".docx"
	result.Name = outName

	defer func() {
		if r := recover(); r != nil {
			result.OK = false
			result.Error = fmt.Sprintf("panic: %v", r)
			result.Elapsed = time.Since(start).String()
		}
	}()

	tplDoc, err := docx.OpenBytes(tplBytes)
	if err != nil {
		return regtest.FileResult{Name: outName, OK: false, Error: fmt.Sprintf("open template: %v", err), Elapsed: time.Since(start).String()}
	}

	markBytes, err := os.ReadFile(markPath)
	if err != nil {
		return regtest.FileResult{Name: outName, OK: false, Error: fmt.Sprintf("read mark %s: %v", markFilename, err), Elapsed: time.Since(start).String()}
	}

	source, err := docx.OpenBytes(markBytes)
	if err != nil {
		return regtest.FileResult{Name: outName, OK: false, Error: fmt.Sprintf("open mark %s: %v", markFilename, err), Elapsed: time.Since(start).String()}
	}

	n, err := tplDoc.ReplaceWithContent(tag, docx.ContentData{
		Source:  source,
		Format:  v.mode,
		Options: v.opts,
	})
	if err != nil {
		return regtest.FileResult{Name: outName, OK: false, Error: fmt.Sprintf("replace %s: %v", tag, err), Elapsed: time.Since(start).String()}
	}

	dstPath := filepath.Join(outputDir, outName)
	if err := tplDoc.SaveFile(dstPath); err != nil {
		return regtest.FileResult{Name: outName, OK: false, Error: fmt.Sprintf("save: %v", err), Elapsed: time.Since(start).String()}
	}

	return regtest.FileResult{
		Name:         outName,
		OK:           true,
		Replacements: n,
		Elapsed:      time.Since(start).String(),
	}
}
