// Package codegen generates Go source code from YAML schema definitions
// describing Office Open XML element types.
//
// It replicates the behavior of python-docx's xmlchemy.py descriptor system,
// generating typed accessor methods for child elements and attributes based on
// their cardinality and type.
package codegen

// Schema is the root object of a YAML schema file.
type Schema struct {
	Package  string    `yaml:"package"`
	Imports  []string  `yaml:"imports"`
	Elements []Element `yaml:"elements"`
}

// Element describes one CT_* element class.
type Element struct {
	Name         string        `yaml:"name"`          // Go struct name, e.g. "CT_P"
	Tag          string        `yaml:"tag"`           // XML tag, e.g. "w:p"
	Doc          string        `yaml:"doc"`           // documentation comment
	Children     []Child       `yaml:"children"`      // child elements
	Attributes   []Attribute   `yaml:"attributes"`    // XML attributes
	ChoiceGroups []ChoiceGroup `yaml:"choice_groups"` // ZeroOrOneChoice groups
}

// Cardinality describes the multiplicity of a child element.
type Cardinality string

const (
	ZeroOrOne     Cardinality = "zero_or_one"
	ZeroOrMore    Cardinality = "zero_or_more"
	OneAndOnlyOne Cardinality = "one_and_only_one"
	OneOrMore     Cardinality = "one_or_more"
)

func (c Cardinality) valid() bool {
	switch c {
	case ZeroOrOne, ZeroOrMore, OneAndOnlyOne, OneOrMore:
		return true
	}
	return false
}

// Child describes a child element with its cardinality.
//
// Cardinality determines which methods are generated:
//   - ZeroOrOne:       getter, GetOrAdd, Remove, add, new, insert (6 methods)
//   - ZeroOrMore:      List getter, Add, add, new, insert          (5 methods)
//   - OneAndOnlyOne:   getter only (returns error if absent)        (1 method)
//   - OneOrMore:       List getter, Add, add, new, insert          (5 methods)
type Child struct {
	Name        string      `yaml:"name"`        // Go property name, e.g. "PPr"
	Tag         string      `yaml:"tag"`         // XML tag, e.g. "w:pPr"
	Type        string      `yaml:"type"`        // Go type name, e.g. "CT_PPr"
	Cardinality Cardinality `yaml:"cardinality"` // see above
	Successors  []string    `yaml:"successors"`  // tags for InsertElementBefore ordering
}

// Attribute describes an XML attribute on an element.
type Attribute struct {
	Name     string  `yaml:"name"`      // Go property name, e.g. "Val"
	AttrName string  `yaml:"attr_name"` // XML attribute name, e.g. "w:val" or "val"
	Type     string  `yaml:"type"`      // "string", "int", "int64", "bool", or enum
	Required bool    `yaml:"required"`  // true = RequiredAttribute, false = OptionalAttribute
	Default  *string `yaml:"default"`   // default value expression (only for optional)
}

// ChoiceGroup describes a ZeroOrOneChoice element group (e.g. EG_ColorChoice).
type ChoiceGroup struct {
	Name       string   `yaml:"name"`       // group property name, e.g. "ColorChoice"
	Choices    []Choice `yaml:"choices"`    // member elements
	Successors []string `yaml:"successors"` // tags for InsertElementBefore ordering
}

// Choice describes one member of a ChoiceGroup.
type Choice struct {
	Name string `yaml:"name"` // Go property name, e.g. "SchemeClr"
	Tag  string `yaml:"tag"`  // XML tag, e.g. "a:schemeClr"
	Type string `yaml:"type"` // Go type name, e.g. "CT_SchemeClr"
}
