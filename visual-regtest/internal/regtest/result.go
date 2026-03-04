// Package regtest provides shared types and utilities for all
// visual-regtest programs.
package regtest

import (
	"encoding/json"
	"os"
)

// FileResult captures the outcome of one generation, roundtrip, or replacement.
// Compatible with compare_ssim.py manifest reader.
type FileResult struct {
	Name         string `json:"name"`
	OK           bool   `json:"ok"`
	Error        string `json:"error,omitempty"`
	Elapsed      string `json:"elapsed"`
	Replacements int    `json:"replacements,omitempty"`
}

// WriteManifest writes a JSON manifest file at the given path.
func WriteManifest(path string, results []FileResult) error {
	data, err := json.MarshalIndent(results, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0o644)
}
