package codegen

import (
	"fmt"
	"strings"
)

// Validate checks the schema for errors and returns a combined error
// listing every problem found. It is called by NewGenerator before
// any code is emitted.
func (s *Schema) Validate() error {
	var errs []string

	for _, el := range s.Elements {
		for _, ch := range el.Children {
			if !ch.Cardinality.valid() {
				errs = append(errs, fmt.Sprintf(
					"element %s, child %s: invalid cardinality %q",
					el.Name, ch.Name, ch.Cardinality))
			}
		}
	}

	if len(errs) > 0 {
		return fmt.Errorf("schema validation:\n  %s", strings.Join(errs, "\n  "))
	}
	return nil
}
