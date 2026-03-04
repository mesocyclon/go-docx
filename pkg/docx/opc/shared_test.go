package opc

import (
	"testing"
)

func TestCaseInsensitiveMap_SetGet(t *testing.T) {
	t.Parallel()

	m := newCaseInsensitiveMap()
	m.Set("Content-Type", "text/xml")

	if got := m.Get("content-type"); got != "text/xml" {
		t.Errorf("Get lowercase: got %q, want %q", got, "text/xml")
	}
	if got := m.Get("CONTENT-TYPE"); got != "text/xml" {
		t.Errorf("Get uppercase: got %q, want %q", got, "text/xml")
	}
	if got := m.Get("Content-Type"); got != "text/xml" {
		t.Errorf("Get original case: got %q, want %q", got, "text/xml")
	}
	if got := m.Get("nonexistent"); got != "" {
		t.Errorf("Get nonexistent: got %q, want empty", got)
	}
}

func TestCaseInsensitiveMap_Has(t *testing.T) {
	t.Parallel()

	m := newCaseInsensitiveMap()
	m.Set("Accept", "application/json")

	if !m.Has("accept") {
		t.Error("Has should be case-insensitive")
	}
	if !m.Has("ACCEPT") {
		t.Error("Has should match uppercase")
	}
	if m.Has("missing") {
		t.Error("Has should return false for missing keys")
	}
}

func TestCaseInsensitiveMap_Overwrite(t *testing.T) {
	t.Parallel()

	m := newCaseInsensitiveMap()
	m.Set("Key", "value1")
	m.Set("key", "value2") // same key, different case

	if got := m.Get("key"); got != "value2" {
		t.Errorf("after overwrite: got %q, want %q", got, "value2")
	}

	// Should not duplicate keys
	keys := m.SortedKeys()
	if len(keys) != 1 {
		t.Errorf("expected 1 key after overwrite, got %d: %v", len(keys), keys)
	}
}

func TestCaseInsensitiveMap_SortedKeys(t *testing.T) {
	t.Parallel()

	m := newCaseInsensitiveMap()
	m.Set("Zebra", "z")
	m.Set("Apple", "a")
	m.Set("Mango", "m")

	keys := m.SortedKeys()
	if len(keys) != 3 {
		t.Fatalf("expected 3 keys, got %d", len(keys))
	}
	if keys[0] != "Apple" || keys[1] != "Mango" || keys[2] != "Zebra" {
		t.Errorf("expected sorted keys [Apple Mango Zebra], got %v", keys)
	}
}

func TestSortedStringKeys(t *testing.T) {
	t.Parallel()

	m := map[string]string{
		"c": "3",
		"a": "1",
		"b": "2",
	}
	keys := sortedStringKeys(m)
	if len(keys) != 3 {
		t.Fatalf("expected 3 keys, got %d", len(keys))
	}
	if keys[0] != "a" || keys[1] != "b" || keys[2] != "c" {
		t.Errorf("expected sorted keys [a b c], got %v", keys)
	}
}

func TestSortedStringKeys_Empty(t *testing.T) {
	t.Parallel()

	m := map[string]string{}
	keys := sortedStringKeys(m)
	if len(keys) != 0 {
		t.Errorf("expected 0 keys for empty map, got %d", len(keys))
	}
}
