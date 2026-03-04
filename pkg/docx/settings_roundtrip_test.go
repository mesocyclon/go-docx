package docx

import (
	"bytes"
	"testing"
)

// -----------------------------------------------------------------------
// Settings round-trip tests (MR-13)
// -----------------------------------------------------------------------

func TestSettings_OddAndEvenPages_RoundTrip_True(t *testing.T) {
	doc := mustNewDoc(t)
	settings, err := doc.Settings()
	if err != nil {
		t.Fatalf("Settings(): %v", err)
	}

	if err := settings.SetOddAndEvenPagesHeaderFooter(true); err != nil {
		t.Fatalf("SetOddAndEvenPagesHeaderFooter(true): %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	settings2, err := doc2.Settings()
	if err != nil {
		t.Fatalf("Settings(): %v", err)
	}
	if !settings2.OddAndEvenPagesHeaderFooter() {
		t.Error("expected OddAndEvenPagesHeaderFooter()=true after round-trip")
	}
}

func TestSettings_OddAndEvenPages_RoundTrip_False(t *testing.T) {
	doc := mustNewDoc(t)
	settings, err := doc.Settings()
	if err != nil {
		t.Fatalf("Settings(): %v", err)
	}

	// Set true then false to ensure the element is toggled.
	if err := settings.SetOddAndEvenPagesHeaderFooter(true); err != nil {
		t.Fatalf("SetOddAndEvenPagesHeaderFooter(true): %v", err)
	}
	if err := settings.SetOddAndEvenPagesHeaderFooter(false); err != nil {
		t.Fatalf("SetOddAndEvenPagesHeaderFooter(false): %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	settings2, err := doc2.Settings()
	if err != nil {
		t.Fatalf("Settings(): %v", err)
	}
	if settings2.OddAndEvenPagesHeaderFooter() {
		t.Error("expected OddAndEvenPagesHeaderFooter()=false after round-trip")
	}
}

func TestSettings_OddAndEvenPages_Toggle(t *testing.T) {
	doc := mustNewDoc(t)
	settings, err := doc.Settings()
	if err != nil {
		t.Fatalf("Settings(): %v", err)
	}

	// Default should be false.
	if settings.OddAndEvenPagesHeaderFooter() {
		t.Error("expected default OddAndEvenPagesHeaderFooter()=false")
	}

	// Toggle true → false → true.
	if err := settings.SetOddAndEvenPagesHeaderFooter(true); err != nil {
		t.Fatalf("SetOddAndEvenPagesHeaderFooter(true): %v", err)
	}
	if !settings.OddAndEvenPagesHeaderFooter() {
		t.Error("expected true after setting true")
	}

	if err := settings.SetOddAndEvenPagesHeaderFooter(false); err != nil {
		t.Fatalf("SetOddAndEvenPagesHeaderFooter(false): %v", err)
	}
	if settings.OddAndEvenPagesHeaderFooter() {
		t.Error("expected false after setting false")
	}

	if err := settings.SetOddAndEvenPagesHeaderFooter(true); err != nil {
		t.Fatalf("SetOddAndEvenPagesHeaderFooter(true): %v", err)
	}
	if !settings.OddAndEvenPagesHeaderFooter() {
		t.Error("expected true after re-setting true")
	}
}

func TestSettings_MultipleRoundTrips(t *testing.T) {
	// Create → set → save → open → verify → save → open → verify.
	doc := mustNewDoc(t)
	settings, _ := doc.Settings()
	settings.SetOddAndEvenPagesHeaderFooter(true)

	var buf1 bytes.Buffer
	doc.Save(&buf1)
	doc2, _ := OpenBytes(buf1.Bytes())

	settings2, _ := doc2.Settings()
	if !settings2.OddAndEvenPagesHeaderFooter() {
		t.Error("round-trip 1: expected true")
	}

	var buf2 bytes.Buffer
	doc2.Save(&buf2)
	doc3, _ := OpenBytes(buf2.Bytes())

	settings3, _ := doc3.Settings()
	if !settings3.OddAndEvenPagesHeaderFooter() {
		t.Error("round-trip 2: expected true")
	}
}
