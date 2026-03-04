package opc

import (
	"fmt"
	"io"
)

// PackageWriter writes an OPC package to a ZIP stream.
type PackageWriter struct{}

// Write serializes the package relationships and parts to the writer.
func (pw *PackageWriter) Write(w io.Writer, pkgRels *Relationships, parts []Part) error {
	physWriter := NewPhysPkgWriter(w)

	// 1. Write [Content_Types].xml
	if err := pw.writeContentTypes(physWriter, parts); err != nil {
		return err
	}

	// 2. Write package-level .rels
	if err := pw.writeRels(physWriter, PackageURI, pkgRels); err != nil {
		return err
	}

	// 3. Write each part's blob and its .rels (if any)
	for _, part := range parts {
		blob, err := part.Blob()
		if err != nil {
			return fmt.Errorf("opc: serializing part %q: %w", part.PartName(), err)
		}
		if err := physWriter.Write(part.PartName(), blob); err != nil {
			return fmt.Errorf("opc: writing part %q: %w", part.PartName(), err)
		}
		if part.Rels() != nil && part.Rels().Len() > 0 {
			if err := pw.writeRels(physWriter, part.PartName(), part.Rels()); err != nil {
				return err
			}
		}
	}

	return physWriter.Close()
}

func (pw *PackageWriter) writeContentTypes(physWriter *PhysPkgWriter, parts []Part) error {
	var infos []PartInfo
	for _, p := range parts {
		infos = append(infos, PartInfo{
			PartName:    p.PartName(),
			ContentType: p.ContentType(),
		})
	}
	blob, err := SerializeContentTypes(infos)
	if err != nil {
		return fmt.Errorf("opc: writing content types: %w", err)
	}
	return physWriter.Write(ContentTypesURI, blob)
}

func (pw *PackageWriter) writeRels(physWriter *PhysPkgWriter, sourceURI PackURI, rels *Relationships) error {
	blob, err := SerializeRelationships(rels)
	if err != nil {
		return fmt.Errorf("opc: serializing rels for %q: %w", sourceURI, err)
	}
	relsURI := sourceURI.RelsURI()
	return physWriter.Write(relsURI, blob)
}
