package regtest

import "log"

// Must panics via log.Fatalf if err != nil, otherwise returns v.
func Must[T any](v T, err error) T {
	if err != nil {
		log.Fatalf("fatal: %v", err)
	}
	return v
}

// Must0 panics via log.Fatalf if err != nil.
func Must0(err error) {
	if err != nil {
		log.Fatalf("fatal: %v", err)
	}
}
