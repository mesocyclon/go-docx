package docx

import "fmt"

// DocxError is the base error type for all go-docx errors.
// It implements Unwrap() so errors.Is / errors.As traverse the chain.
type DocxError struct {
	msg   string
	cause error
}

func (e *DocxError) Error() string { return e.msg }
func (e *DocxError) Unwrap() error { return e.cause }

// NewDocxError creates a DocxError. cause may be nil.
func NewDocxError(cause error, msg string, args ...any) *DocxError {
	return &DocxError{msg: fmt.Sprintf(msg, args...), cause: cause}
}

// InvalidXmlError indicates invalid or non-conformant XML.
type InvalidXmlError struct{ DocxError }

func NewInvalidXmlError(cause error, msg string, args ...any) *InvalidXmlError {
	return &InvalidXmlError{DocxError{msg: fmt.Sprintf(msg, args...), cause: cause}}
}

// PackageNotFoundError indicates a missing package file.
type PackageNotFoundError struct{ DocxError }

func NewPackageNotFoundError(cause error, msg string, args ...any) *PackageNotFoundError {
	return &PackageNotFoundError{DocxError{msg: fmt.Sprintf(msg, args...), cause: cause}}
}

// InvalidSpanError indicates an invalid table cell span.
type InvalidSpanError struct{ DocxError }

func NewInvalidSpanError(cause error, msg string, args ...any) *InvalidSpanError {
	return &InvalidSpanError{DocxError{msg: fmt.Sprintf(msg, args...), cause: cause}}
}

// errIndexOutOfRange returns an IndexError-equivalent for collections.
func errIndexOutOfRange(collection string, idx, length int) error {
	return fmt.Errorf("docx: %s index [%d] out of range (len=%d)", collection, idx, length)
}
