// Command codegen generates Go source code from YAML schema files
// describing Office Open XML element types.
//
// Usage:
//
//	go run ./cmd/codegen -schema ./schema/ -out ./pkg/docx/oxml/
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/vortex/go-docx/internal/codegen"

	"gopkg.in/yaml.v3"
)

func main() {
	schemaDir := flag.String("schema", "", "Path to YAML schema directory")
	outDir := flag.String("out", "", "Output directory for generated .go files")
	flag.Parse()

	if *schemaDir == "" || *outDir == "" {
		fmt.Fprintf(os.Stderr, "Usage: codegen -schema <dir> -out <dir>\n")
		os.Exit(1)
	}

	entries, err := os.ReadDir(*schemaDir)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error reading schema directory %q: %v\n", *schemaDir, err)
		os.Exit(1)
	}

	count := 0
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if !strings.HasSuffix(name, ".yaml") && !strings.HasSuffix(name, ".yml") {
			continue
		}

		schemaPath := filepath.Join(*schemaDir, name)
		if err := processSchema(schemaPath, *outDir); err != nil {
			fmt.Fprintf(os.Stderr, "Error processing %q: %v\n", schemaPath, err)
			os.Exit(1)
		}

		baseName := strings.TrimSuffix(strings.TrimSuffix(name, ".yaml"), ".yml")
		fmt.Printf("Generated: zz_gen_%s.go\n", baseName)
		count++
	}

	fmt.Printf("Done. Generated %d file(s).\n", count)
}

func processSchema(schemaPath, outDir string) error {
	data, err := os.ReadFile(schemaPath)
	if err != nil {
		return fmt.Errorf("reading %s: %w", schemaPath, err)
	}

	var schema codegen.Schema
	if err := yaml.Unmarshal(data, &schema); err != nil {
		return fmt.Errorf("parsing YAML %s: %w", schemaPath, err)
	}

	gen, err := codegen.NewGenerator(schema)
	if err != nil {
		return fmt.Errorf("creating generator: %w", err)
	}

	output, err := gen.Generate()
	if err != nil {
		return fmt.Errorf("generating code: %w", err)
	}

	baseName := strings.TrimSuffix(filepath.Base(schemaPath), filepath.Ext(schemaPath))
	outPath := filepath.Join(outDir, "zz_gen_"+baseName+".go")

	if err := os.WriteFile(outPath, output, 0o644); err != nil {
		return fmt.Errorf("writing %s: %w", outPath, err)
	}

	return nil
}
