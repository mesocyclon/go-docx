package oxml

import (
	"strconv"
	"strings"
)

// --- Attribute conversion helpers used by generated and custom code ---

// parseIntAttr parses a string attribute value into an int.
// Returns an error if the string is not a valid integer.
func parseIntAttr(s string) (int, error) {
	return strconv.Atoi(strings.TrimSpace(s))
}

// parseInt64Attr parses a string attribute value into an int64.
// Returns an error if the string is not a valid int64.
func parseInt64Attr(s string) (int64, error) {
	return strconv.ParseInt(strings.TrimSpace(s), 10, 64)
}

// parseBoolAttr parses an XML boolean attribute value.
// Accepts "true", "1", "on" as true; everything else is false.
//
// This function is intentionally infallible: the xsd:boolean value space
// is small and a non-matching string mapping to false is a reasonable default.
func parseBoolAttr(s string) bool {
	s = strings.TrimSpace(strings.ToLower(s))
	return s == "true" || s == "1" || s == "on"
}

// parseEnum parses an XML attribute value using the provided fromXml function.
// Returns the parsed enum value or the error from fromXml.
func parseEnum[T any](s string, fromXml func(string) (T, error)) (T, error) {
	return fromXml(s)
}

// --- Format helpers (unchanged — formatting is infallible) ---

// formatStringAttr formats a string as an attribute value.
func formatStringAttr(v string) (string, error) {
	return v, nil
}

// formatIntAttr formats an int as a string attribute value.
func formatIntAttr(v int) (string, error) {
	return strconv.Itoa(v), nil
}

// formatInt64Attr formats an int64 as a string attribute value.
func formatInt64Attr(v int64) (string, error) {
	return strconv.FormatInt(v, 10), nil
}

// formatBoolAttr formats a bool as an XML attribute value.
func formatBoolAttr(v bool) (string, error) {
	if v {
		return "true", nil
	}
	return "false", nil
}
