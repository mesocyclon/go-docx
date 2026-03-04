package codegen

import (
	"bytes"
	"embed"
	"fmt"
	"go/format"
	"strings"
	"text/template"
)

// go run ./cmd/codegen -schema ./schema/ -out ./pkg/docx/oxml/

//go:embed templates/element.go.tmpl
var templateFS embed.FS

// Generator generates Go source code from a Schema.
type Generator struct {
	schema Schema
	tmpl   *template.Template
}

// NewGenerator creates a new Generator for the given schema.
func NewGenerator(schema Schema) (*Generator, error) {
	if err := schema.Validate(); err != nil {
		return nil, err
	}

	tmplContent, err := templateFS.ReadFile("templates/element.go.tmpl")
	if err != nil {
		return nil, fmt.Errorf("codegen: reading template: %w", err)
	}

	tmpl, err := template.New("element").Parse(string(tmplContent))
	if err != nil {
		return nil, fmt.Errorf("codegen: parsing template: %w", err)
	}

	return &Generator{schema: schema, tmpl: tmpl}, nil
}

// Generate produces gofmt-formatted Go source code.
func (g *Generator) Generate() ([]byte, error) {
	data := g.buildTemplateData()

	var buf bytes.Buffer
	if err := g.tmpl.Execute(&buf, data); err != nil {
		return nil, fmt.Errorf("codegen: executing template: %w", err)
	}

	formatted, err := format.Source(buf.Bytes())
	if err != nil {
		return buf.Bytes(), fmt.Errorf("codegen: gofmt failed: %w\n--- raw output ---\n%s", err, buf.String())
	}

	return formatted, nil
}

// --- Template data types ---

type templateData struct {
	Package  string
	Imports  []string
	Elements []elementData
}

type elementData struct {
	Name                  string
	Tag                   string
	Doc                   string
	ZeroOrOneChildren     []childData
	ZeroOrMoreChildren    []childData
	OneAndOnlyOneChildren []childData
	OneOrMoreChildren     []childData
	OptionalAttributes    []attrData
	RequiredAttributes    []attrData
	ChoiceGroups          []choiceGroupData
}

type childData struct {
	GoName     string   // exported Go name
	Tag        string   // XML tag, e.g. "w:pPr"
	Type       string   // Go type, e.g. "CT_PPr"
	ParentType string   // parent struct name
	Successors []string // successor tags for insert ordering
}

type attrData struct {
	GoName      string // exported Go name
	AttrName    string // XML attribute name
	GoType      string // Go type string
	ParentType  string // parent struct name
	DefaultExpr string // default value expression (optional attrs)
	ZeroExpr    string // zero value expression (required attrs)
	ParseExpr   string // expression to parse string "val" → typed value
	FormatExpr  string // expression to format typed "v" → (string, error)
	Failable    bool   // true when ParseExpr returns (T, error) instead of T
	IsPointer   bool   // true when GoType is a pointer (e.g. *enum.Wd…); result needs &
}

type choiceGroupData struct {
	GoName     string       // group property name
	ParentType string       // parent struct name
	Tags       []string     // all member tags
	Choices    []choiceData // per-choice data
}

type choiceData struct {
	GoName     string   // choice member name
	Tag        string   // XML tag
	Type       string   // Go type
	ParentType string   // parent struct name
	GroupName  string   // group property name (for Remove call)
	Successors []string // successor tags
}

// --- Build template data ---

func (g *Generator) buildTemplateData() templateData {
	data := templateData{
		Package: g.schema.Package,
		Imports: g.schema.Imports,
	}

	for _, el := range g.schema.Elements {
		ed := elementData{
			Name: el.Name,
			Tag:  el.Tag,
			Doc:  el.Doc,
		}

		for _, ch := range el.Children {
			cd := childData{
				GoName:     ExportName(ch.Name),
				Tag:        ch.Tag,
				Type:       ch.Type,
				ParentType: el.Name,
				Successors: ch.Successors,
			}

			switch ch.Cardinality {
			case ZeroOrOne:
				ed.ZeroOrOneChildren = append(ed.ZeroOrOneChildren, cd)
			case ZeroOrMore:
				ed.ZeroOrMoreChildren = append(ed.ZeroOrMoreChildren, cd)
			case OneAndOnlyOne:
				ed.OneAndOnlyOneChildren = append(ed.OneAndOnlyOneChildren, cd)
			case OneOrMore:
				ed.OneOrMoreChildren = append(ed.OneOrMoreChildren, cd)
			default:
				panic(fmt.Sprintf("codegen: unhandled cardinality %q for %s.%s (should be caught by Validate)",
					ch.Cardinality, el.Name, ch.Name))
			}
		}

		for _, attr := range el.Attributes {
			rt := resolveAttrType(attr)

			ad := attrData{
				GoName:     ExportName(attr.Name),
				AttrName:   attr.AttrName,
				GoType:     rt.GoType,
				ParentType: el.Name,
				ParseExpr:  rt.ParseExpr,
				FormatExpr: rt.FormatExpr,
				Failable:   rt.Failable,
				IsPointer:  rt.IsPointer,
			}

			if attr.Required {
				ad.ZeroExpr = rt.ZeroExpr
				ed.RequiredAttributes = append(ed.RequiredAttributes, ad)
			} else {
				ad.DefaultExpr = rt.DefaultExpr
				if attr.Default != nil {
					ad.DefaultExpr = *attr.Default
				}
				ed.OptionalAttributes = append(ed.OptionalAttributes, ad)
			}
		}

		for _, cg := range el.ChoiceGroups {
			tags := make([]string, len(cg.Choices))
			choices := make([]choiceData, len(cg.Choices))
			for i, ch := range cg.Choices {
				tags[i] = ch.Tag
				choices[i] = choiceData{
					GoName:     ExportName(ch.Name),
					Tag:        ch.Tag,
					Type:       ch.Type,
					ParentType: el.Name,
					GroupName:  ExportName(cg.Name),
					Successors: cg.Successors,
				}
			}
			ed.ChoiceGroups = append(ed.ChoiceGroups, choiceGroupData{
				GoName:     ExportName(cg.Name),
				ParentType: el.Name,
				Tags:       tags,
				Choices:    choices,
			})
		}

		data.Elements = append(data.Elements, ed)
	}

	return data
}

// resolvedType bundles type-resolution results for an attribute.
type resolvedType struct {
	GoType      string
	ZeroExpr    string
	DefaultExpr string
	ParseExpr   string
	FormatExpr  string
	Failable    bool // ParseExpr returns (T, error) rather than T
	IsPointer   bool // GoType is a pointer; template must & the parsed value
}

// resolveAttrType determines Go type information for a schema attribute.
// parseExpr uses "val" as the string variable;
// formatExpr uses "v" as the typed variable and always returns (string, error).
func resolveAttrType(attr Attribute) resolvedType {
	switch attr.Type {
	case "string":
		return resolvedType{
			GoType: "string", ZeroExpr: `""`, DefaultExpr: `""`,
			ParseExpr: "val", FormatExpr: "formatStringAttr(v)",
			Failable: false,
		}

	case "int":
		if attr.Required {
			return resolvedType{
				GoType: "int", ZeroExpr: "0", DefaultExpr: "0",
				ParseExpr: "parseIntAttr(val)", FormatExpr: "formatIntAttr(v)",
				Failable: true,
			}
		}
		return resolvedType{
			GoType: "*int", ZeroExpr: "nil", DefaultExpr: "nil",
			ParseExpr: "parseIntAttr(val)", FormatExpr: "formatIntAttr(*v)",
			Failable: true, IsPointer: true,
		}

	case "int64":
		if attr.Required {
			return resolvedType{
				GoType: "int64", ZeroExpr: "0", DefaultExpr: "0",
				ParseExpr: "parseInt64Attr(val)", FormatExpr: "formatInt64Attr(v)",
				Failable: true,
			}
		}
		return resolvedType{
			GoType: "*int64", ZeroExpr: "nil", DefaultExpr: "nil",
			ParseExpr: "parseInt64Attr(val)", FormatExpr: "formatInt64Attr(*v)",
			Failable: true, IsPointer: true,
		}

	case "bool":
		return resolvedType{
			GoType: "bool", ZeroExpr: "false", DefaultExpr: "false",
			ParseExpr: "parseBoolAttr(val)", FormatExpr: "formatBoolAttr(v)",
			Failable: false,
		}

	default:
		// Enum or custom type, e.g. "enum.WdAlignParagraph"
		if strings.HasPrefix(attr.Type, "*") {
			// Optional pointer-to-enum
			inner := attr.Type[1:]
			fromFn := inner + "FromXml"
			return resolvedType{
				GoType: attr.Type, ZeroExpr: "nil", DefaultExpr: "nil",
				ParseExpr:  fmt.Sprintf("parseEnum(val, %s)", fromFn),
				FormatExpr: "(*v).ToXml()",
				Failable:   true, IsPointer: true,
			}
		}
		// Required or value enum
		fromFn := attr.Type + "FromXml"
		return resolvedType{
			GoType: attr.Type, ZeroExpr: attr.Type + "(0)", DefaultExpr: attr.Type + "(0)",
			ParseExpr:  fmt.Sprintf("parseEnum(val, %s)", fromFn),
			FormatExpr: "v.ToXml()",
			Failable:   true,
		}
	}
}

// ExportName ensures the first character is uppercase (Go exported).
func ExportName(name string) string {
	if name == "" {
		return name
	}
	if name[0] >= 'A' && name[0] <= 'Z' {
		return name
	}
	return strings.ToUpper(name[:1]) + name[1:]
}
