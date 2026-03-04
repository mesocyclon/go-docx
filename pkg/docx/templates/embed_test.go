package templates

import "testing"

func TestEmbeddedFilesExist(t *testing.T) {
	t.Parallel()
	files := []string{
		"default.docx",
		"default-header.xml",
		"default-footer.xml",
		"default-settings.xml",
		"default-styles.xml",
		"default-comments.xml",
		"default-numbering.xml",
		"default-footnotes.xml",
		"default-endnotes.xml",
	}
	for _, name := range files {
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			data, err := FS.ReadFile(name)
			if err != nil {
				t.Fatalf("FS.ReadFile(%q) failed: %v", name, err)
			}
			if len(data) == 0 {
				t.Errorf("FS.ReadFile(%q) returned empty content", name)
			}
		})
	}
}

func TestEmbeddedFileCount(t *testing.T) {
	t.Parallel()
	entries, err := FS.ReadDir(".")
	if err != nil {
		t.Fatalf("FS.ReadDir(\".\") failed: %v", err)
	}
	if len(entries) != 9 {
		t.Errorf("expected 9 embedded files, got %d", len(entries))
		for _, e := range entries {
			t.Logf("  - %s", e.Name())
		}
	}
}
