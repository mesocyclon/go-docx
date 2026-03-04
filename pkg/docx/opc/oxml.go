package opc

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// --------------------------------------------------------------------------
// [Content_Types].xml structures
// --------------------------------------------------------------------------

// xmlTypes is the root <Types> element in [Content_Types].xml.
type xmlTypes struct {
	XMLName   xml.Name      `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Defaults  []xmlDefault  `xml:"Default"`
	Overrides []xmlOverride `xml:"Override"`
}

// xmlDefault is a <Default> element mapping an extension to a content type.
type xmlDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// xmlOverride is an <Override> element mapping a partname to a content type.
type xmlOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// ParseContentTypes parses [Content_Types].xml bytes into a ContentTypeMap.
func ParseContentTypes(blob []byte) (*ContentTypeMap, error) {
	var types xmlTypes
	if err := xml.Unmarshal(blob, &types); err != nil {
		return nil, fmt.Errorf("opc: parsing [Content_Types].xml: %w", err)
	}
	ct := NewContentTypeMap()
	for _, d := range types.Defaults {
		ct.AddDefault(d.Extension, d.ContentType)
	}
	for _, o := range types.Overrides {
		ct.AddOverride(o.PartName, o.ContentType)
	}
	return ct, nil
}

// SerializeContentTypes builds [Content_Types].xml bytes from the given parts.
func SerializeContentTypes(parts []PartInfo) ([]byte, error) {
	types := xmlTypes{}

	// Always include rels and xml defaults
	defaults := newCaseInsensitiveMap()
	defaults.Set("rels", CTOpcRelationships)
	defaults.Set("xml", CTXml)

	overrides := make(map[string]string)

	for _, p := range parts {
		ext := p.PartName.Ext()
		ct := p.ContentType
		if IsDefaultContentType(ext, ct) {
			defaults.Set(ext, ct)
		} else {
			overrides[string(p.PartName)] = ct
		}
	}

	for _, key := range defaults.SortedKeys() {
		types.Defaults = append(types.Defaults, xmlDefault{
			Extension:   key,
			ContentType: defaults.Get(key),
		})
	}
	// Sort overrides by partname for determinism
	sortedPartnames := sortedStringKeys(overrides)
	for _, pn := range sortedPartnames {
		types.Overrides = append(types.Overrides, xmlOverride{
			PartName:    pn,
			ContentType: overrides[pn],
		})
	}

	output, err := xml.MarshalIndent(types, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("opc: serializing [Content_Types].xml: %w", err)
	}
	return append([]byte(xml.Header), output...), nil
}

// PartInfo is a minimal struct for SerializeContentTypes.
type PartInfo struct {
	PartName    PackURI
	ContentType string
}

// --------------------------------------------------------------------------
// .rels file structures
// --------------------------------------------------------------------------

// xmlRelationships is the root <Relationships> element in a .rels file.
type xmlRelationships struct {
	XMLName       xml.Name          `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xmlRelationship `xml:"Relationship"`
}

// xmlRelationship is a single <Relationship> element.
type xmlRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// ParseRelationships parses a .rels XML blob into a slice of SerializedRelationship.
// Returns an empty slice if blob is nil or empty.
func ParseRelationships(blob []byte, baseURI string) ([]SerializedRelationship, error) {
	if len(blob) == 0 {
		return nil, nil
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(blob, &rels); err != nil {
		return nil, fmt.Errorf("opc: parsing .rels: %w", err)
	}

	result := make([]SerializedRelationship, 0, len(rels.Relationships))
	for _, r := range rels.Relationships {
		targetMode := TargetModeInternal
		if r.TargetMode == TargetModeExternal {
			targetMode = TargetModeExternal
		}
		result = append(result, SerializedRelationship{
			BaseURI:    baseURI,
			RID:        r.ID,
			RelType:    NormalizeRelType(r.Type),
			TargetRef:  r.Target,
			TargetMode: targetMode,
		})
	}
	return result, nil
}

// SerializeRelationships builds .rels XML bytes from a Relationships collection.
// For internal relationships with a resolved TargetPart, the target reference is
// recomputed from the part's current partname — matching Python's dynamic
// _Relationship.target_ref property behavior.
func SerializeRelationships(rels *Relationships) ([]byte, error) {
	xrels := xmlRelationships{}

	for _, rel := range rels.All() {
		targetRef := rel.TargetRef
		// Recompute target_ref for internal rels with resolved parts,
		// matching Python: target.partname.relative_ref(self._baseURI)
		if !rel.IsExternal && rel.TargetPart != nil {
			targetRef = rel.TargetPart.PartName().RelativeRef(rels.BaseURI())
		}
		xr := xmlRelationship{
			ID:     rel.RID,
			Type:   rel.RelType,
			Target: targetRef,
		}
		if rel.IsExternal {
			xr.TargetMode = TargetModeExternal
		}
		xrels.Relationships = append(xrels.Relationships, xr)
	}

	output, err := xml.MarshalIndent(xrels, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("opc: serializing .rels: %w", err)
	}
	return append([]byte(xml.Header), output...), nil
}

// --------------------------------------------------------------------------
// ContentTypeMap
// --------------------------------------------------------------------------

// ContentTypeMap resolves content types for parts by partname or extension.
type ContentTypeMap struct {
	defaults  *caseInsensitiveMap
	overrides *caseInsensitiveMap
}

// NewContentTypeMap creates an empty ContentTypeMap.
func NewContentTypeMap() *ContentTypeMap {
	return &ContentTypeMap{
		defaults:  newCaseInsensitiveMap(),
		overrides: newCaseInsensitiveMap(),
	}
}

// AddDefault adds a default (extension-based) content type mapping.
func (m *ContentTypeMap) AddDefault(ext, contentType string) {
	m.defaults.Set(ext, contentType)
}

// AddOverride adds an override (partname-based) content type mapping.
func (m *ContentTypeMap) AddOverride(partname, contentType string) {
	m.overrides.Set(partname, contentType)
}

// ContentType returns the content type for the given PackURI.
// Override takes precedence over default (by extension).
// As a last resort, if the extension matches a well-known OPC default
// (from DefaultContentTypes), that content type is returned.  This handles
// malformed packages where [Content_Types].xml is incomplete — common in
// files from LibreOffice and Google Docs.
func (m *ContentTypeMap) ContentType(uri PackURI) (string, error) {
	if ct := m.overrides.Get(string(uri)); ct != "" {
		return ct, nil
	}
	ext := uri.Ext()
	if ct := m.defaults.Get(ext); ct != "" {
		return ct, nil
	}
	// Fallback: infer from well-known extension → content-type table.
	if ct := inferContentType(ext); ct != "" {
		return ct, nil
	}
	return "", fmt.Errorf("opc: no content type for partname %q in [Content_Types].xml", uri)
}

// inferContentType returns the content type for a well-known file extension,
// or "" if the extension is not recognized.  Only the first match in
// DefaultContentTypes is returned (sufficient for images and common types).
func inferContentType(ext string) string {
	lower := strings.ToLower(ext)
	for _, pair := range DefaultContentTypes {
		if pair.Ext == lower {
			return pair.ContentType
		}
	}
	return ""
}
