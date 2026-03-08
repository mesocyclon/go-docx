package docx

import (
	"testing"
)

// FuzzRWC_TagDiscovery exercises tag discovery with arbitrary tag strings.
// Ensures no panics on unexpected input patterns.
func FuzzRWC_TagDiscovery(f *testing.F) {
	f.Add("[[test]]")
	f.Add("[<CONTENT>]")
	f.Add("{{tag}}")
	f.Add("$PLACEHOLDER$")
	f.Add("[<СОДЕРЖАНИЕ>]")

	f.Fuzz(func(t *testing.T, tag string) {
		if len(tag) == 0 || len(tag) > 100 {
			return
		}
		target, err := New()
		if err != nil {
			return
		}
		target.AddParagraph("prefix " + tag + " suffix")

		source, err := New()
		if err != nil {
			return
		}
		source.AddParagraph("inserted")

		// Must not panic regardless of tag content.
		target.ReplaceWithContent(tag, ContentData{Source: source})
	})
}
