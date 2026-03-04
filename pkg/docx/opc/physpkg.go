package opc

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"
)

// ErrMemberNotFound is returned by BlobFor when the requested member
// does not exist in the ZIP archive.  Callers (e.g. RelsXmlFor) can
// test for it with errors.Is to distinguish "missing" from real I/O
// errors — exactly the way python-docx catches KeyError in
// _ZipPkgReader.rels_xml_for.
var ErrMemberNotFound = errors.New("opc: member not found in package")

// ErrNotZipPackage is returned when the input cannot be read as a ZIP
// archive.  Callers can test with errors.Is(err, ErrNotZipPackage).
var ErrNotZipPackage = errors.New("opc: not a ZIP-based OPC package")

// ErrEncryptedPackage is returned when the input appears to be an OLE2
// Compound Document (encrypted .docx).  Such files require decryption
// before they can be opened as OPC packages.
var ErrEncryptedPackage = errors.New("opc: file is encrypted (OLE2 Compound Document, not a ZIP-based package)")

// ErrPartTooLarge is returned by BlobFor when a decompressed part
// exceeds MaxPartSize. This protects against zip bombs and other
// decompression attacks that could cause out-of-memory conditions.
var ErrPartTooLarge = errors.New("opc: decompressed part exceeds size limit")

// ErrTooManyEntries is returned when the ZIP archive contains more
// entries than MaxEntries. This protects against ZIP bombs that use
// millions of tiny or empty entries to exhaust memory during central
// directory parsing.
var ErrTooManyEntries = errors.New("opc: too many entries in ZIP archive")

// DefaultMaxPartSize is the default maximum decompressed size for a single
// part within an OPC package (256 MB). Override per-reader via
// PhysPkgReader.MaxPartSize.
const DefaultMaxPartSize int64 = 256 << 20 // 256 MB

// DefaultMaxEntries is the default maximum number of entries allowed in
// a ZIP-based OPC package. A typical .docx contains 20–50 entries;
// 10 000 provides generous headroom while blocking million-entry bombs.
const DefaultMaxEntries = 10_000

// --------------------------------------------------------------------------
// PhysPkgReader — reads a ZIP-based OPC package
// --------------------------------------------------------------------------

// PhysPkgReader provides low-level access to a ZIP-based OPC package.
type PhysPkgReader struct {
	reader      *zip.Reader
	closer      io.Closer // non-nil when opened from a file
	files       map[string]*zip.File
	MaxPartSize int64 // maximum decompressed size per part; 0 means DefaultMaxPartSize
}

// NewPhysPkgReader creates a PhysPkgReader from an io.ReaderAt.
func NewPhysPkgReader(r io.ReaderAt, size int64) (*PhysPkgReader, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, wrapZipOpenError(err, r)
	}
	return newPhysPkgReaderFromZip(zr, nil)
}

// NewPhysPkgReaderFromFile opens a PhysPkgReader from a file path.
func NewPhysPkgReaderFromFile(path string) (*PhysPkgReader, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("opc: opening file %q: %w", path, err)
	}
	info, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("opc: stat file %q: %w", path, err)
	}
	zr, err := zip.NewReader(f, info.Size())
	if err != nil {
		wrapped := wrapZipOpenError(err, f)
		f.Close()
		return nil, fmt.Errorf("opening %q: %w", path, wrapped)
	}
	return newPhysPkgReaderFromZip(zr, f)
}

// NewPhysPkgReaderFromBytes creates a PhysPkgReader from in-memory bytes.
func NewPhysPkgReaderFromBytes(data []byte) (*PhysPkgReader, error) {
	r := bytes.NewReader(data)
	return NewPhysPkgReader(r, int64(len(data)))
}

func newPhysPkgReaderFromZip(zr *zip.Reader, closer io.Closer) (*PhysPkgReader, error) {
	if len(zr.File) > DefaultMaxEntries {
		if closer != nil {
			closer.Close()
		}
		return nil, fmt.Errorf("%w: archive contains %d entries (limit %d)",
			ErrTooManyEntries, len(zr.File), DefaultMaxEntries)
	}
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	return &PhysPkgReader{
		reader: zr,
		closer: closer,
		files:  files,
	}, nil
}

// BlobFor returns the contents of the part at the given PackURI.
// The decompressed size is capped at MaxPartSize (or DefaultMaxPartSize
// when MaxPartSize is 0) to guard against zip bombs.
func (p *PhysPkgReader) BlobFor(uri PackURI) ([]byte, error) {
	membername := uri.Membername()
	f, ok := p.files[membername]
	if !ok {
		return nil, fmt.Errorf("%w: %s", ErrMemberNotFound, membername)
	}
	rc, err := f.Open()
	if err != nil {
		return nil, fmt.Errorf("opc: opening member %q: %w", membername, err)
	}
	defer rc.Close()

	limit := p.MaxPartSize
	if limit <= 0 {
		limit = DefaultMaxPartSize
	}
	// Read up to limit+1 bytes: if we get more than limit, the part is too large.
	lr := io.LimitReader(rc, limit+1)
	data, err := io.ReadAll(lr)
	if err != nil {
		return nil, fmt.Errorf("opc: reading member %q: %w", membername, err)
	}
	if int64(len(data)) > limit {
		return nil, fmt.Errorf("%w: %s (%d bytes exceeds %d byte limit)",
			ErrPartTooLarge, membername, f.UncompressedSize64, limit)
	}
	return data, nil
}

// ContentTypesXml returns the [Content_Types].xml blob.
func (p *PhysPkgReader) ContentTypesXml() ([]byte, error) {
	return p.BlobFor(ContentTypesURI)
}

// RelsXmlFor returns the .rels XML for the given source URI, or nil if none exists.
func (p *PhysPkgReader) RelsXmlFor(sourceURI PackURI) ([]byte, error) {
	relsURI := sourceURI.RelsURI()
	blob, err := p.BlobFor(relsURI)
	if err != nil {
		// No .rels file is not an error — it simply means no relationships.
		// This mirrors python-docx's _ZipPkgReader.rels_xml_for which
		// catches KeyError from ZipFile.read and returns None.
		if errors.Is(err, ErrMemberNotFound) {
			return nil, nil
		}
		return nil, err
	}
	return blob, nil
}

// URIs returns a list of all member URIs in the package, excluding
// [Content_Types].xml and .rels files.
func (p *PhysPkgReader) URIs() []PackURI {
	var uris []PackURI
	for name := range p.files {
		uri := NewPackURI(name)
		// Skip [Content_Types].xml and .rels files
		if uri == ContentTypesURI {
			continue
		}
		if strings.Contains(name, "_rels/") && strings.HasSuffix(name, ".rels") {
			continue
		}
		uris = append(uris, uri)
	}
	return uris
}

// Close releases resources held by the reader.
func (p *PhysPkgReader) Close() error {
	if p.closer != nil {
		return p.closer.Close()
	}
	return nil
}

// ole2Magic is the signature of an OLE2 Compound Document (D0 CF 11 E0 A1 B1 1A E1).
// Encrypted .docx/.xlsx/.pptx files are wrapped in this format.
var ole2Magic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// wrapZipOpenError enriches a zip.NewReader failure with a more specific
// sentinel error by inspecting the first bytes of the input.
func wrapZipOpenError(zipErr error, r io.ReaderAt) error {
	var header [8]byte
	if n, _ := r.ReadAt(header[:], 0); n >= len(ole2Magic) {
		if bytes.Equal(header[:len(ole2Magic)], ole2Magic) {
			return fmt.Errorf("%w: %w", ErrEncryptedPackage, zipErr)
		}
	}
	return fmt.Errorf("%w: %w", ErrNotZipPackage, zipErr)
}

// --------------------------------------------------------------------------
// PhysPkgWriter — writes a ZIP-based OPC package
// --------------------------------------------------------------------------

// PhysPkgWriter provides low-level write access to a ZIP-based OPC package.
type PhysPkgWriter struct {
	writer *zip.Writer
}

// NewPhysPkgWriter creates a PhysPkgWriter backed by the given writer.
func NewPhysPkgWriter(w io.Writer) *PhysPkgWriter {
	return &PhysPkgWriter{writer: zip.NewWriter(w)}
}

// Write adds a member to the ZIP package.
func (p *PhysPkgWriter) Write(uri PackURI, blob []byte) error {
	membername := uri.Membername()
	w, err := p.writer.Create(membername)
	if err != nil {
		return fmt.Errorf("opc: creating zip member %q: %w", membername, err)
	}
	if _, err := w.Write(blob); err != nil {
		return fmt.Errorf("opc: writing zip member %q: %w", membername, err)
	}
	return nil
}

// Close finalizes the ZIP archive.
func (p *PhysPkgWriter) Close() error {
	return p.writer.Close()
}
