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
//	  use_dest/
//	    alpha.docx            ← filled template
//	    beta.docx
//	    manifest.json
//	  keep_src/
//	    ...
//
// Each variant is written to its own subdirectory under out/.
// Use --variant to run a single variant or omit for all.
//
// Usage:
//
//	go run ./visual-regtest/replace-user-mark-batch \
//	    --input  visual-regtest/replace-user-mark-batch/in \
//	    --output visual-regtest/replace-user-mark-batch/out \
//	    --tag '(<mark>)' --variant keep_src --workers 4
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

// importVariant describes one ImportFormatMode + ImportFormatOptions combination.
type importVariant struct {
	ID   string
	Mode docx.ImportFormatMode
	Opts docx.ImportFormatOptions
}

var variantRegistry = []importVariant{
	{"use_dest", docx.UseDestinationStyles, docx.ImportFormatOptions{}},
	{"keep_src", docx.KeepSourceFormatting, docx.ImportFormatOptions{}},
	{"keep_src_force", docx.KeepSourceFormatting, docx.ImportFormatOptions{ForceCopyStyles: true}},
	{"keep_diff", docx.KeepDifferentStyles, docx.ImportFormatOptions{}},
	{"keep_diff_force", docx.KeepDifferentStyles, docx.ImportFormatOptions{ForceCopyStyles: true}},
	{"keep_num", docx.UseDestinationStyles, docx.ImportFormatOptions{KeepSourceNumbering: true}},
}

func findVariant(id string) (importVariant, bool) {
	for _, v := range variantRegistry {
		if v.ID == id {
			return v, true
		}
	}
	return importVariant{}, false
}

func variantIDs() []string {
	ids := make([]string, len(variantRegistry))
	for i, v := range variantRegistry {
		ids[i] = v.ID
	}
	return ids
}

func main() {
	inputDir := flag.String("input", "visual-regtest/replace-user-mark-batch/in", "directory with template.docx and mark/")
	outputDir := flag.String("output", "visual-regtest/replace-user-mark-batch/out", "directory for output .docx files")
	tag := flag.String("tag", "(<mark>)", "placeholder tag to replace in template")
	workers := flag.Int("workers", 4, "parallel workers")
	variant := flag.String("variant", "", "variant ID to run (empty = all); available: "+strings.Join(variantIDs(), ", "))
	flag.Parse()

	if *workers < 1 {
		*workers = 1
	}

	var variants []importVariant
	if *variant == "" || *variant == "all" {
		variants = variantRegistry
	} else {
		v, ok := findVariant(*variant)
		if !ok {
			log.Fatalf("unknown variant %q; available: %s", *variant, strings.Join(variantIDs(), ", "))
		}
		variants = []importVariant{v}
	}

	templatePath := filepath.Join(*inputDir, "template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		log.Fatalf("template not found: %s", templatePath)
	}

	marks := loadMarks(filepath.Join(*inputDir, "mark"))

	tplBytes, err := os.ReadFile(templatePath)
	if err != nil {
		log.Fatalf("reading template: %v", err)
	}

	log.Printf("template: %s", templatePath)
	log.Printf("tag:      %s", *tag)
	log.Printf("marks:    %d file(s)", len(marks))
	log.Printf("variants: %s", joinVariantIDs(variants))
	log.Printf("workers:  %d", *workers)

	for _, v := range variants {
		varDir := filepath.Join(*outputDir, v.ID)
		if err := os.MkdirAll(varDir, 0o755); err != nil {
			log.Fatalf("creating output dir %s: %v", varDir, err)
		}

		results := runVariant(tplBytes, marks, *tag, varDir, v, *workers)

		if err := regtest.WriteManifest(filepath.Join(varDir, "manifest.json"), results); err != nil {
			log.Fatalf("writing manifest: %v", err)
		}

		okCount := 0
		for _, r := range results {
			if r.OK {
				okCount++
			}
		}
		log.Printf("variant %s: %d/%d succeeded", v.ID, okCount, len(results))
	}

	log.Printf("done")
}

func joinVariantIDs(vs []importVariant) string {
	ids := make([]string, len(vs))
	for i, v := range vs {
		ids[i] = v.ID
	}
	return strings.Join(ids, ", ")
}

type markEntry struct {
	name, filename, path string
}

func loadMarks(markDir string) []markEntry {
	if _, err := os.Stat(markDir); os.IsNotExist(err) {
		log.Fatalf("mark directory not found: %s", markDir)
	}

	entries, err := os.ReadDir(markDir)
	if err != nil {
		log.Fatalf("reading mark directory: %v", err)
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
	return marks
}

func runVariant(tplBytes []byte, marks []markEntry, tag, outputDir string, v importVariant, workers int) []regtest.FileResult {
	log.Printf("── variant: %s ──", v.ID)

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
	for i := 0; i < workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for m := range jobs {
				r := processOneMark(tplBytes, m.name, m.filename, m.path, tag, outputDir, v)
				mu.Lock()
				results = append(results, r)
				mu.Unlock()
				if r.OK {
					log.Printf("  OK   %-30s %d replacement(s)  %s", r.Name, r.Replacements, r.Elapsed)
				} else {
					log.Printf("  FAIL %-30s %s", r.Name, r.Error)
				}
			}
		}()
	}
	wg.Wait()
	return results
}

func processOneMark(tplBytes []byte, markName, markFilename, markPath, tag, outputDir string, v importVariant) (result regtest.FileResult) {
	start := time.Now()
	outName := markName + ".docx"
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
		Format:  v.Mode,
		Options: v.Opts,
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
