// Package enum provides enumerations used throughout the go-docx library,
// corresponding to the MS Office API enumerations.
package enum

import "fmt"

// FromXml looks up an enum member by its XML attribute value using the provided mapping.
// Returns an error if the XML value is not found in the mapping.
func FromXml[T comparable](mapping map[string]T, xmlVal string) (T, error) {
	if v, ok := mapping[xmlVal]; ok {
		return v, nil
	}
	var zero T
	return zero, fmt.Errorf("no enum member mapped to XML value %q", xmlVal)
}

// ToXml looks up the XML attribute value for the given enum member.
// Returns an error if the member has no XML representation.
func ToXml[T comparable](mapping map[T]string, val T) (string, error) {
	if s, ok := mapping[val]; ok {
		return s, nil
	}
	var zero T
	_ = zero
	return "", fmt.Errorf("enum member %v has no XML representation", val)
}

// invertMap creates a reverse mapping from value→key to key→value.
func invertMap[K, V comparable](m map[K]V) map[V]K {
	inv := make(map[V]K, len(m))
	for k, v := range m {
		inv[v] = k
	}
	return inv
}
