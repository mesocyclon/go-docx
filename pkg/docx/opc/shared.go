package opc

import (
	"sort"
	"strings"
)

// caseInsensitiveMap is a simple map that stores and looks up keys in lowercase.
type caseInsensitiveMap struct {
	data map[string]string
	// keys stores the original-case key for iteration order
	keys []string
}

func newCaseInsensitiveMap() *caseInsensitiveMap {
	return &caseInsensitiveMap{data: make(map[string]string)}
}

func (m *caseInsensitiveMap) Set(key, value string) {
	lower := strings.ToLower(key)
	if _, exists := m.data[lower]; !exists {
		m.keys = append(m.keys, key)
	}
	m.data[lower] = value
}

func (m *caseInsensitiveMap) Get(key string) string {
	return m.data[strings.ToLower(key)]
}

func (m *caseInsensitiveMap) Has(key string) bool {
	_, ok := m.data[strings.ToLower(key)]
	return ok
}

func (m *caseInsensitiveMap) SortedKeys() []string {
	sorted := make([]string, len(m.keys))
	copy(sorted, m.keys)
	sort.Strings(sorted)
	return sorted
}

// sortedStringKeys returns the keys of a map[string]string sorted alphabetically.
func sortedStringKeys(m map[string]string) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	return keys
}
