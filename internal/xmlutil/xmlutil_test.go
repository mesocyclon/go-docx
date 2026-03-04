package xmlutil

import "testing"

func TestIsDigits(t *testing.T) {
	tests := []struct {
		in   string
		want bool
	}{
		{"", false},
		{"0", true},
		{"42", true},
		{"007", true},
		{"123456789", true},
		{" 1", false},
		{"1 ", false},
		{"12.3", false},
		{"-1", false},
		{"abc", false},
		{"1a2", false},
	}
	for _, tt := range tests {
		if got := IsDigits(tt.in); got != tt.want {
			t.Errorf("IsDigits(%q) = %v, want %v", tt.in, got, tt.want)
		}
	}
}
