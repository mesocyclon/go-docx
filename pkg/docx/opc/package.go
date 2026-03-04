package opc

import (
	"bytes"
	"fmt"
	"io"
	"os"
)

// OpcPackage is the root object representing an OPC package.
type OpcPackage struct {
	rels        *Relationships
	partFactory *PartFactory
	parts       map[PackURI]Part
	appPkg      any // application-level package (e.g. *parts.WmlPackage); mirrors Python Package(OpcPackage) inheritance
}

// NewOpcPackage creates an empty OpcPackage.
func NewOpcPackage(factory *PartFactory) *OpcPackage {
	if factory == nil {
		factory = NewPartFactory()
	}
	return &OpcPackage{
		rels:        NewRelationships("/"),
		partFactory: factory,
		parts:       make(map[PackURI]Part),
	}
}

// AppPackage returns the application-level package stored on this OpcPackage.
//
// In Python, Package(OpcPackage) uses inheritance — every Part._package IS
// the WML-level Package subclass. In Go we use composition: the WML-level
// WmlPackage is stored here, and any Part that needs WML services (e.g.
// image deduplication) retrieves it via Part.Package().AppPackage().
func (p *OpcPackage) AppPackage() any {
	return p.appPkg
}

// SetAppPackage stores the application-level package on this OpcPackage.
func (p *OpcPackage) SetAppPackage(app any) {
	p.appPkg = app
}

// --------------------------------------------------------------------------
// Open
// --------------------------------------------------------------------------

// Open reads an OPC package from an io.ReaderAt.
func Open(r io.ReaderAt, size int64, factory *PartFactory) (*OpcPackage, error) {
	physReader, err := NewPhysPkgReader(r, size)
	if err != nil {
		return nil, err
	}
	defer physReader.Close()
	return openFromPhysReader(physReader, factory)
}

// OpenFile opens an OPC package from a file path.
func OpenFile(path string, factory *PartFactory) (*OpcPackage, error) {
	physReader, err := NewPhysPkgReaderFromFile(path)
	if err != nil {
		return nil, err
	}
	defer physReader.Close()
	return openFromPhysReader(physReader, factory)
}

// OpenBytes opens an OPC package from in-memory bytes.
func OpenBytes(data []byte, factory *PartFactory) (*OpcPackage, error) {
	physReader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		return nil, err
	}
	defer physReader.Close()
	return openFromPhysReader(physReader, factory)
}

func openFromPhysReader(physReader *PhysPkgReader, factory *PartFactory) (*OpcPackage, error) {
	if factory == nil {
		factory = NewPartFactory()
	}
	pkg := NewOpcPackage(factory)

	reader := &PackageReader{}
	result, err := reader.Read(physReader)
	if err != nil {
		return nil, err
	}

	// Unmarshal: create parts
	parts := make(map[PackURI]Part, len(result.SParts))
	for _, sp := range result.SParts {
		part, err := factory.New(sp.Partname, sp.ContentType, sp.RelType, sp.Blob, pkg)
		if err != nil {
			return nil, fmt.Errorf("opc: creating part %q: %w", sp.Partname, err)
		}
		parts[sp.Partname] = part
	}

	// Wire up package-level relationships.
	// Mirrors python-docx Unmarshaller._unmarshal_relationships where
	// source is the package and target is parts[srel.target_partname].
	for _, srel := range result.PkgSRels {
		var targetPart Part
		if !srel.IsExternal() {
			pn := srel.TargetPartname()
			p, ok := parts[pn]
			if ok {
				targetPart = p
			}
			// If !ok, targetPart stays nil — the dangling relationship
			// is preserved with its original TargetRef so the .rels XML
			// round-trips faithfully.  The serializer already handles
			// nil TargetPart by falling back to TargetRef.
		}
		pkg.rels.Load(srel.RID, srel.RelType, srel.TargetRef, targetPart, srel.IsExternal())
	}

	// Wire up part-level relationships.
	// Mirrors the same Python loop with source = parts[source_uri].
	for _, sp := range result.SParts {
		part := parts[sp.Partname] // guaranteed to exist: built from same SParts slice
		rels := NewRelationships(sp.Partname.BaseURI())
		for _, srel := range sp.SRels {
			var targetPart Part
			if !srel.IsExternal() {
				pn := srel.TargetPartname()
				p, ok := parts[pn]
				if ok {
					targetPart = p
				}
				// Dangling rel preserved — see comment in pkg-level loop above.
			}
			rels.Load(srel.RID, srel.RelType, srel.TargetRef, targetPart, srel.IsExternal())
		}
		part.SetRels(rels)
	}

	pkg.parts = parts

	// Call AfterUnmarshal on all parts in load order (Python iterates
	// parts.values() which preserves insertion order from iter_sparts).
	for _, sp := range result.SParts {
		parts[sp.Partname].AfterUnmarshal()
	}

	return pkg, nil
}

// --------------------------------------------------------------------------
// Save
// --------------------------------------------------------------------------

// Save writes the package to an io.Writer.
func (p *OpcPackage) Save(w io.Writer) error {
	// Collect parts once via deterministic DFS traversal (mirrors Python
	// Package.save which calls self.parts → list(self.iter_parts()) for
	// both before_marshal and PackageWriter.write).
	parts := p.Parts()

	for _, part := range parts {
		part.BeforeMarshal()
	}

	pw := &PackageWriter{}
	return pw.Write(w, p.rels, parts)
}

// SaveToFile writes the package to a file.
func (p *OpcPackage) SaveToFile(path string) (err error) {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("opc: creating file %q: %w", path, err)
	}
	defer func() {
		if closeErr := f.Close(); err == nil {
			err = closeErr
		}
	}()
	return p.Save(f)
}

// SaveToBytes returns the package as a byte slice.
func (p *OpcPackage) SaveToBytes() ([]byte, error) {
	var buf bytes.Buffer
	if err := p.Save(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// --------------------------------------------------------------------------
// Accessors
// --------------------------------------------------------------------------

// Rels returns the package-level relationships.
func (p *OpcPackage) Rels() *Relationships {
	return p.rels
}

// Parts returns all parts reachable via the relationship graph.
// Order is deterministic: depth-first traversal of rels, matching Python's
// OpcPackage.parts property which returns list(self.iter_parts()).
func (p *OpcPackage) Parts() []Part {
	return p.IterParts()
}

// PartByName returns a part by its PackURI.
func (p *OpcPackage) PartByName(pn PackURI) (Part, bool) {
	part, ok := p.parts[pn]
	return part, ok
}

// RelatedPart returns the part that the package has a relationship of relType to.
func (p *OpcPackage) RelatedPart(relType string) (Part, error) {
	rel, err := p.rels.GetByRelType(relType)
	if err != nil {
		return nil, err
	}
	if rel.IsExternal || rel.TargetPart == nil {
		return nil, fmt.Errorf("opc: relationship %q is external or unresolved", relType)
	}
	return rel.TargetPart, nil
}

// MainDocumentPart returns the main document part (via RT.OFFICE_DOCUMENT relationship).
func (p *OpcPackage) MainDocumentPart() (Part, error) {
	return p.RelatedPart(RTOfficeDocument)
}

// RelateTo creates or returns an existing package-level relationship to the given part.
func (p *OpcPackage) RelateTo(part Part, relType string) string {
	rel := p.rels.GetOrAdd(relType, part)
	return rel.RID
}

// AddPart adds a part to the package.
func (p *OpcPackage) AddPart(part Part) {
	p.parts[part.PartName()] = part
}

// NextPartname returns the next available partname matching the template (printf-style).
// E.g. NextPartname("/word/header%d.xml") might return "/word/header1.xml".
func (p *OpcPackage) NextPartname(template string) PackURI {
	partnames := make(map[PackURI]bool, len(p.parts))
	for pn := range p.parts {
		partnames[pn] = true
	}
	for n := 1; n <= len(partnames)+2; n++ {
		candidate := PackURI(fmt.Sprintf(template, n))
		if !partnames[candidate] {
			return candidate
		}
	}
	return PackURI(fmt.Sprintf(template, len(partnames)+1))
}

// IterParts generates all parts reachable via the relationship graph.
// Uses iterative DFS to avoid unbounded call-stack growth on deep
// relationship chains.
func (p *OpcPackage) IterParts() []Part {
	var result []Part
	visited := make(map[Part]bool)
	// Explicit stack: each entry is a slice of relationships to process.
	// We push slices in reverse so the first rel is popped first,
	// preserving the original DFS order.
	stack := []([]*Relationship){p.rels.All()}

	for len(stack) > 0 {
		top := len(stack) - 1
		rels := stack[top]

		// Find next unvisited part in current rels slice.
		var advanced bool
		for len(rels) > 0 {
			rel := rels[0]
			rels = rels[1:]
			stack[top] = rels // consume

			if rel.IsExternal || rel.TargetPart == nil {
				continue
			}
			part := rel.TargetPart
			if visited[part] {
				continue
			}
			visited[part] = true
			result = append(result, part)
			// Push child rels — will be processed before remaining siblings.
			stack = append(stack, part.Rels().All())
			advanced = true
			break
		}
		if !advanced {
			// Current slice exhausted — pop it.
			stack = stack[:top]
		}
	}
	return result
}

// IterRels yields every relationship in the package exactly once via a
// depth-first traversal of the relationship graph. Mirrors Python
// OpcPackage.iter_rels.
// Uses iterative DFS to avoid unbounded call-stack growth.
func (p *OpcPackage) IterRels() []*Relationship {
	var result []*Relationship
	visited := make(map[Part]bool)
	stack := []([]*Relationship){p.rels.All()}

	for len(stack) > 0 {
		top := len(stack) - 1
		rels := stack[top]

		if len(rels) == 0 {
			stack = stack[:top]
			continue
		}

		rel := rels[0]
		stack[top] = rels[1:] // consume

		result = append(result, rel)

		if rel.IsExternal || rel.TargetPart == nil {
			continue
		}
		part := rel.TargetPart
		if visited[part] {
			continue
		}
		visited[part] = true
		stack = append(stack, part.Rels().All())
	}
	return result
}
