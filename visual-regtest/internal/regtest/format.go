package regtest

import "fmt"

// TruncQuote quotes a string, truncating if longer than 30 chars.
func TruncQuote(s string) string {
	if len(s) > 30 {
		return fmt.Sprintf("%q...", s[:27])
	}
	return fmt.Sprintf("%q", s)
}
