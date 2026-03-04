package oxml

import (
	"errors"
	"fmt"
	"strconv"
	"testing"
)

func TestParseIntAttr(t *testing.T) {
	t.Parallel()
	t.Run("valid values", func(t *testing.T) {
		t.Parallel()
		tests := []struct {
			input string
			want  int
		}{
			{"42", 42},
			{" 100 ", 100},
			{"0", 0},
			{"-5", -5},
		}
		for _, tc := range tests {
			t.Run(tc.input, func(t *testing.T) {
				t.Parallel()
				got, err := parseIntAttr(tc.input)
				if err != nil {
					t.Fatalf("parseIntAttr(%q) unexpected error: %v", tc.input, err)
				}
				if got != tc.want {
					t.Errorf("parseIntAttr(%q) = %d, want %d", tc.input, got, tc.want)
				}
			})
		}
	})

	t.Run("invalid values return error", func(t *testing.T) {
		t.Parallel()
		for _, input := range []string{"abc", "", "12.5", "1e3"} {
			t.Run(input, func(t *testing.T) {
				t.Parallel()
				_, err := parseIntAttr(input)
				if err == nil {
					t.Errorf("parseIntAttr(%q) expected error, got nil", input)
				}
				var numErr *strconv.NumError
				if !errors.As(err, &numErr) {
					t.Errorf("parseIntAttr(%q) error is %T, want *strconv.NumError", input, err)
				}
			})
		}
	})
}

func TestParseInt64Attr(t *testing.T) {
	t.Parallel()
	t.Run("valid", func(t *testing.T) {
		t.Parallel()
		got, err := parseInt64Attr("914400")
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != 914400 {
			t.Errorf("parseInt64Attr(\"914400\") = %d, want 914400", got)
		}
	})

	t.Run("negative", func(t *testing.T) {
		t.Parallel()
		got, err := parseInt64Attr("-360000")
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != -360000 {
			t.Errorf("parseInt64Attr(\"-360000\") = %d, want -360000", got)
		}
	})

	t.Run("invalid returns error", func(t *testing.T) {
		t.Parallel()
		_, err := parseInt64Attr("not_a_number")
		if err == nil {
			t.Error("parseInt64Attr(\"not_a_number\") expected error, got nil")
		}
	})
}

func TestParseBoolAttr(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input string
		want  bool
	}{
		{"true", true},
		{"True", true},
		{"TRUE", true},
		{"1", true},
		{"on", true},
		{"false", false},
		{"0", false},
		{"", false},
		{"off", false},
		{"garbage", false},
	}
	for _, tc := range tests {
		t.Run(tc.input, func(t *testing.T) {
			t.Parallel()
			if got := parseBoolAttr(tc.input); got != tc.want {
				t.Errorf("parseBoolAttr(%q) = %v, want %v", tc.input, got, tc.want)
			}
		})
	}
}

func TestParseEnum(t *testing.T) {
	t.Parallel()
	fromXml := func(s string) (int, error) {
		m := map[string]int{"a": 1, "b": 2}
		if v, ok := m[s]; ok {
			return v, nil
		}
		return 0, fmt.Errorf("unknown value: %s", s)
	}

	t.Run("valid", func(t *testing.T) {
		t.Parallel()
		got, err := parseEnum("a", fromXml)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != 1 {
			t.Errorf("parseEnum(\"a\") = %d, want 1", got)
		}
	})

	t.Run("invalid returns error", func(t *testing.T) {
		t.Parallel()
		_, err := parseEnum("unknown", fromXml)
		if err == nil {
			t.Error("parseEnum(\"unknown\") expected error, got nil")
		}
	})
}

func TestFormatIntAttr(t *testing.T) {
	t.Parallel()
	got, err := formatIntAttr(42)
	if err != nil {
		t.Fatalf("formatIntAttr(42) error: %v", err)
	}
	if got != "42" {
		t.Errorf("formatIntAttr(42) = %q, want %q", got, "42")
	}
}

func TestFormatInt64Attr(t *testing.T) {
	t.Parallel()
	got, err := formatInt64Attr(914400)
	if err != nil {
		t.Fatalf("formatInt64Attr(914400) error: %v", err)
	}
	if got != "914400" {
		t.Errorf("formatInt64Attr(914400) = %q", got)
	}
}

func TestFormatBoolAttr(t *testing.T) {
	t.Parallel()
	got, err := formatBoolAttr(true)
	if err != nil {
		t.Fatalf("formatBoolAttr(true) error: %v", err)
	}
	if got != "true" {
		t.Errorf("formatBoolAttr(true) = %q", got)
	}
	got, err = formatBoolAttr(false)
	if err != nil {
		t.Fatalf("formatBoolAttr(false) error: %v", err)
	}
	if got != "false" {
		t.Errorf("formatBoolAttr(false) = %q", got)
	}
}

// --- ParseAttrError tests ---

func TestParseAttrError_Error(t *testing.T) {
	t.Parallel()
	inner := fmt.Errorf("strconv.Atoi: parsing \"abc\": invalid syntax")
	pe := &ParseAttrError{
		Element:  "w:pgMar",
		Attr:     "w:top",
		RawValue: "abc",
		Err:      inner,
	}
	msg := pe.Error()
	if msg == "" {
		t.Fatal("Error() returned empty string")
	}
	// Verify key parts are present
	for _, want := range []string{"w:pgMar", "w:top", "abc"} {
		if !contains(msg, want) {
			t.Errorf("Error() = %q, missing %q", msg, want)
		}
	}
}

func TestParseAttrError_Unwrap(t *testing.T) {
	t.Parallel()
	inner := &strconv.NumError{Func: "Atoi", Num: "abc", Err: strconv.ErrSyntax}
	pe := &ParseAttrError{
		Element:  "w:pgMar",
		Attr:     "w:top",
		RawValue: "abc",
		Err:      inner,
	}

	// errors.Is should see through to the strconv error
	if !errors.Is(pe, strconv.ErrSyntax) {
		t.Error("errors.Is(pe, strconv.ErrSyntax) = false, want true")
	}

	// errors.As should match the inner type
	var numErr *strconv.NumError
	if !errors.As(pe, &numErr) {
		t.Error("errors.As(pe, *strconv.NumError) = false, want true")
	}
}

func contains(s, substr string) bool {
	return len(s) >= len(substr) && searchString(s, substr)
}

func searchString(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}
