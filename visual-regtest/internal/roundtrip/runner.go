// Package roundtrip implements the parallel roundtrip pipeline shared by
// the OPC and DOCX visual regression tests.
package roundtrip

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

// Saver is implemented by both *opc.OpcPackage and *docx.Document.
type Saver interface {
	SaveFile(dst string) error
}

// opcSaverAdapter wraps OpcPackage.SaveToFile as SaveFile.
type opcSaverAdapter struct {
	saveFunc func(path string) error
}

func (a opcSaverAdapter) SaveFile(dst string) error { return a.saveFunc(dst) }

// WrapSaveToFile wraps a SaveToFile function into the Saver interface.
// Use for OPC layer where the method is named SaveToFile instead of SaveFile.
func WrapSaveToFile(saveToFile func(path string) error) Saver {
	return opcSaverAdapter{saveFunc: saveToFile}
}

// Opener opens a .docx file and returns a Saver.
type Opener func(srcPath string) (Saver, error)

// Run collects .docx files from inputDir, processes them in parallel using
// the provided opener, saves results to outputDir, and returns results.
func Run(inputDir, outputDir string, workers int, open Opener) []regtest.FileResult {
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	entries, err := os.ReadDir(inputDir)
	if err != nil {
		log.Fatalf("reading input dir: %v", err)
	}

	var files []string
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		if strings.HasSuffix(strings.ToLower(e.Name()), ".docx") {
			files = append(files, e.Name())
		}
	}
	log.Printf("found %d .docx files", len(files))

	jobs := make(chan string, len(files))
	for _, f := range files {
		jobs <- f
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
			for name := range jobs {
				r := processFile(name, inputDir, outputDir, open)
				mu.Lock()
				results = append(results, r)
				mu.Unlock()
				if !r.OK {
					log.Printf("FAIL %s: %s", name, r.Error)
				}
			}
		}()
	}
	wg.Wait()

	okCount := 0
	for _, r := range results {
		if r.OK {
			okCount++
		}
	}
	log.Printf("done: %d/%d succeeded", okCount, len(results))

	return results
}

func processFile(name, inputDir, outputDir string, open Opener) regtest.FileResult {
	start := time.Now()
	srcPath := filepath.Join(inputDir, name)
	dstPath := filepath.Join(outputDir, name)

	saver, err := open(srcPath)
	if err != nil {
		return regtest.FileResult{Name: name, OK: false, Error: fmt.Sprintf("open: %v", err), Elapsed: time.Since(start).String()}
	}

	if err := saver.SaveFile(dstPath); err != nil {
		return regtest.FileResult{Name: name, OK: false, Error: fmt.Sprintf("save: %v", err), Elapsed: time.Since(start).String()}
	}

	return regtest.FileResult{Name: name, OK: true, Elapsed: time.Since(start).String()}
}
