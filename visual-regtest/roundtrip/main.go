// roundtrip reads every .docx from --input, opens it via the specified
// layer (opc or docx), re-serialises it unchanged, and writes the result
// to --output.
//
// Two layers can be tested:
//
//	opc   – OPC packaging only  (opc.OpenFile → SaveToFile)
//	docx  – full stack          (docx.OpenFile → doc.SaveFile)
//
// Exit code 0  = all files processed (some may have had errors).
// A per-file JSON manifest is written to --output/manifest.json so
// downstream tools know which files succeeded and which failed.
package main

import (
	"flag"
	"log"
	"path/filepath"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/visual-regtest/internal/regtest"
	"github.com/vortex/go-docx/visual-regtest/internal/roundtrip"
)

func main() {
	inputDir := flag.String("input", "", "directory containing original .docx files")
	outputDir := flag.String("output", "", "directory for roundtripped .docx files")
	workers := flag.Int("workers", 8, "parallel workers")
	layer := flag.String("layer", "docx", "layer to test: opc or docx")
	flag.Parse()

	if *inputDir == "" || *outputDir == "" {
		log.Fatal("--input and --output are required")
	}

	var opener roundtrip.Opener
	switch *layer {
	case "opc":
		opener = func(src string) (roundtrip.Saver, error) {
			pkg, err := opc.OpenFile(src, nil)
			if err != nil {
				return nil, err
			}
			return roundtrip.WrapSaveToFile(pkg.SaveToFile), nil
		}
	case "docx":
		opener = func(src string) (roundtrip.Saver, error) {
			return docx.OpenFile(src)
		}
	default:
		log.Fatalf("unknown layer %q (must be opc or docx)", *layer)
	}

	results := roundtrip.Run(*inputDir, *outputDir, *workers, opener)

	manifestPath := filepath.Join(*outputDir, "manifest.json")
	if err := regtest.WriteManifest(manifestPath, results); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}
}
