package oxml

import "fmt"

// ParseAttrError indicates that an XML attribute value could not be parsed
// into the expected Go type. It carries enough context for diagnostics:
// which element, which attribute, what the raw value was, and the underlying
// parsing error.
//
// Callers can match this type with errors.As:
//
//	var pe *oxml.ParseAttrError
//	if errors.As(err, &pe) {
//	    log.Printf("bad attribute %s=%q on <%s>: %v", pe.Attr, pe.RawValue, pe.Element, pe.Err)
//	}
type ParseAttrError struct {
	Element  string // XML tag of the element, e.g. "w:pgMar"
	Attr     string // attribute name, e.g. "w:top"
	RawValue string // the raw string value that failed to parse
	Err      error  // underlying error (e.g. strconv.ErrSyntax)
}

func (e *ParseAttrError) Error() string {
	return fmt.Sprintf("oxml: cannot parse attribute %s=%q on <%s>: %v",
		e.Attr, e.RawValue, e.Element, e.Err)
}

func (e *ParseAttrError) Unwrap() error {
	return e.Err
}
