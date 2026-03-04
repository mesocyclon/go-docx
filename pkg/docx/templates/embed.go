// Package templates provides embedded default template files for new documents.
package templates

import "embed"

// FS contains the embedded template files used when creating new documents.
//
//go:embed default.docx default-header.xml default-footer.xml default-settings.xml default-styles.xml default-comments.xml default-numbering.xml default-footnotes.xml default-endnotes.xml
var FS embed.FS
