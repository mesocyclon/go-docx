package regtest

// BoolPtr returns a pointer to v.
func BoolPtr(v bool) *bool { return &v }

// IntPtr returns a pointer to v.
func IntPtr(v int) *int { return &v }

// StrPtr returns a pointer to v.
func StrPtr(v string) *string { return &v }

// Int64Ptr returns a pointer to v.
func Int64Ptr(v int64) *int64 { return &v }
