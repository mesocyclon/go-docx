package codegen

import (
	"strings"
	"testing"
)

func TestValidate_ValidCardinalities(t *testing.T) {
	t.Parallel()
	s := Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_P",
			Tag:  "w:p",
			Children: []Child{
				{Name: "A", Tag: "w:a", Type: "CT_A", Cardinality: ZeroOrOne},
				{Name: "B", Tag: "w:b", Type: "CT_B", Cardinality: ZeroOrMore},
				{Name: "C", Tag: "w:c", Type: "CT_C", Cardinality: OneAndOnlyOne},
				{Name: "D", Tag: "w:d", Type: "CT_D", Cardinality: OneOrMore},
			},
		}},
	}
	if err := s.Validate(); err != nil {
		t.Errorf("expected no error, got: %v", err)
	}
}

func TestValidate_InvalidCardinality(t *testing.T) {
	t.Parallel()
	s := Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_P",
			Tag:  "w:p",
			Children: []Child{
				{Name: "Bad", Tag: "w:bad", Type: "CT_Bad", Cardinality: "zero_or_onne"},
			},
		}},
	}
	err := s.Validate()
	if err == nil {
		t.Fatal("expected validation error, got nil")
	}
	if !strings.Contains(err.Error(), "zero_or_onne") {
		t.Errorf("error should mention the invalid value, got: %v", err)
	}
	if !strings.Contains(err.Error(), "CT_P") {
		t.Errorf("error should mention the parent element, got: %v", err)
	}
	if !strings.Contains(err.Error(), "Bad") {
		t.Errorf("error should mention the child name, got: %v", err)
	}
}

func TestValidate_EmptyCardinality(t *testing.T) {
	t.Parallel()
	s := Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_P",
			Tag:  "w:p",
			Children: []Child{
				{Name: "X", Tag: "w:x", Type: "CT_X", Cardinality: ""},
			},
		}},
	}
	err := s.Validate()
	if err == nil {
		t.Fatal("expected validation error for empty cardinality, got nil")
	}
}

func TestValidate_MultipleErrors(t *testing.T) {
	t.Parallel()
	s := Schema{
		Package: "oxml",
		Elements: []Element{
			{
				Name: "CT_A",
				Tag:  "w:a",
				Children: []Child{
					{Name: "X", Tag: "w:x", Type: "CT_X", Cardinality: "bogus1"},
				},
			},
			{
				Name: "CT_B",
				Tag:  "w:b",
				Children: []Child{
					{Name: "Y", Tag: "w:y", Type: "CT_Y", Cardinality: "bogus2"},
				},
			},
		},
	}
	err := s.Validate()
	if err == nil {
		t.Fatal("expected validation errors, got nil")
	}
	if !strings.Contains(err.Error(), "bogus1") || !strings.Contains(err.Error(), "bogus2") {
		t.Errorf("error should report both invalid values, got: %v", err)
	}
}

func TestNewGenerator_RejectsInvalidCardinality(t *testing.T) {
	t.Parallel()
	_, err := NewGenerator(Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_P",
			Tag:  "w:p",
			Children: []Child{
				{Name: "Bad", Tag: "w:bad", Type: "CT_Bad", Cardinality: "typo"},
			},
		}},
	})
	if err == nil {
		t.Fatal("NewGenerator should reject schema with invalid cardinality")
	}
}

func TestCardinality_Valid(t *testing.T) {
	t.Parallel()
	valid := []Cardinality{ZeroOrOne, ZeroOrMore, OneAndOnlyOne, OneOrMore}
	for _, c := range valid {
		if !c.valid() {
			t.Errorf("%q should be valid", c)
		}
	}

	invalid := []Cardinality{"", "many", "zero_or_onne", "ZERO_OR_ONE"}
	for _, c := range invalid {
		if c.valid() {
			t.Errorf("%q should be invalid", c)
		}
	}
}
