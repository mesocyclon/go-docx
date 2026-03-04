package opc

import (
	"fmt"
)

// --------------------------------------------------------------------------
// Relationship — a single relationship between parts
// --------------------------------------------------------------------------

// Relationship represents a single OPC relationship from a source part to a target.
type Relationship struct {
	RID        string
	RelType    string
	TargetRef  string // relative reference (for internal) or URL (for external)
	IsExternal bool
	TargetPart Part // resolved target (nil for external relationships)
}

// --------------------------------------------------------------------------
// Relationships — collection of relationships for a part or package
// --------------------------------------------------------------------------

// Relationships manages a collection of Relationship instances for one source.
type Relationships struct {
	baseURI string
	rels    []*Relationship
	byRID   map[string]*Relationship
	nextNum int
}

// NewRelationships creates an empty Relationships collection for the given base URI.
func NewRelationships(baseURI string) *Relationships {
	return &Relationships{
		baseURI: baseURI,
		byRID:   make(map[string]*Relationship),
		nextNum: 1,
	}
}

// Add adds a new relationship and returns it.
// For internal relationships, targetRef is the relative path and targetPart is the
// resolved Part. For external, targetRef is the URL and targetPart is nil.
func (rs *Relationships) Add(relType, targetRef string, targetPart Part, external bool) *Relationship {
	rID := rs.nextRID()
	rel := &Relationship{
		RID:        rID,
		RelType:    relType,
		TargetRef:  targetRef,
		IsExternal: external,
		TargetPart: targetPart,
	}
	rs.rels = append(rs.rels, rel)
	rs.byRID[rID] = rel
	return rel
}

// Load adds a relationship with a known rId (used during package reading).
func (rs *Relationships) Load(rID, relType, targetRef string, targetPart Part, external bool) *Relationship {
	rel := &Relationship{
		RID:        rID,
		RelType:    relType,
		TargetRef:  targetRef,
		IsExternal: external,
		TargetPart: targetPart,
	}
	rs.rels = append(rs.rels, rel)
	rs.byRID[rID] = rel
	// Advance nextNum past this rId if necessary
	if n := parseRIdNum(rID); n >= rs.nextNum {
		rs.nextNum = n + 1
	}
	return rel
}

// GetByRID looks up a relationship by its rId.
func (rs *Relationships) GetByRID(rID string) *Relationship {
	return rs.byRID[rID]
}

// GetByRelType returns the single relationship matching relType.
// Returns an error if none or more than one is found.
func (rs *Relationships) GetByRelType(relType string) (*Relationship, error) {
	var matches []*Relationship
	for _, rel := range rs.rels {
		if rel.RelType == relType {
			matches = append(matches, rel)
		}
	}
	if len(matches) == 0 {
		return nil, fmt.Errorf("opc: no relationship of type %q", relType)
	}
	if len(matches) > 1 {
		return nil, fmt.Errorf("opc: multiple relationships of type %q", relType)
	}
	return matches[0], nil
}

// AllByRelType returns all relationships with the given type.
func (rs *Relationships) AllByRelType(relType string) []*Relationship {
	var result []*Relationship
	for _, rel := range rs.rels {
		if rel.RelType == relType {
			result = append(result, rel)
		}
	}
	return result
}

// GetOrAdd returns an existing relationship of relType to targetPart, or creates one.
func (rs *Relationships) GetOrAdd(relType string, targetPart Part) *Relationship {
	for _, rel := range rs.rels {
		if rel.RelType == relType && !rel.IsExternal && rel.TargetPart == targetPart {
			return rel
		}
	}
	targetRef := targetPart.PartName().RelativeRef(rs.baseURI)
	return rs.Add(relType, targetRef, targetPart, false)
}

// GetOrAddExtRel returns the rId of an existing external rel, or creates one.
func (rs *Relationships) GetOrAddExtRel(relType, targetRef string) string {
	for _, rel := range rs.rels {
		if rel.RelType == relType && rel.IsExternal && rel.TargetRef == targetRef {
			return rel.RID
		}
	}
	rel := rs.Add(relType, targetRef, nil, true)
	return rel.RID
}

// All returns all relationships in order.
func (rs *Relationships) All() []*Relationship {
	return rs.rels
}

// Len returns the number of relationships.
func (rs *Relationships) Len() int {
	return len(rs.rels)
}

// RelatedParts returns a map of rId → Part for all internal relationships.
func (rs *Relationships) RelatedParts() map[string]Part {
	m := make(map[string]Part)
	for _, rel := range rs.rels {
		if !rel.IsExternal && rel.TargetPart != nil {
			m[rel.RID] = rel.TargetPart
		}
	}
	return m
}

// BaseURI returns the base URI of the relationships source.
func (rs *Relationships) BaseURI() string {
	return rs.baseURI
}

// Delete removes a relationship by rId.
func (rs *Relationships) Delete(rID string) {
	delete(rs.byRID, rID)
	for i, rel := range rs.rels {
		if rel.RID == rID {
			rs.rels = append(rs.rels[:i], rs.rels[i+1:]...)
			return
		}
	}
}

func (rs *Relationships) nextRID() string {
	for {
		candidate := fmt.Sprintf("rId%d", rs.nextNum)
		rs.nextNum++
		if _, exists := rs.byRID[candidate]; !exists {
			return candidate
		}
	}
}

// parseRIdNum extracts the numeric portion of an rId like "rId3" → 3.
// Returns 0 if the format doesn't match.
func parseRIdNum(rID string) int {
	n := 0
	if len(rID) > 3 && rID[:3] == "rId" {
		for _, c := range rID[3:] {
			if c >= '0' && c <= '9' {
				n = n*10 + int(c-'0')
			} else {
				return 0
			}
		}
	}
	return n
}
