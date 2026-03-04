// Package docx provides types and functions for creating, reading, and modifying
// Microsoft Word (.docx) documents.
//
// It implements the Office Open XML (OOXML) standard with a layered architecture:
// OPC (package) → parts → oxml (XML types) → docx (domain objects).
//
// # Concurrency
//
// Package docx is not safe for concurrent use. A single [Document] and all
// objects derived from it (paragraphs, runs, tables, sections, etc.) must be
// accessed from one goroutine at a time, or protected by an external mutex.
// Independent Document instances may be used concurrently.
package docx
