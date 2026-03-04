package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

func TestCT_RPr_BoldVal_TriState(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	// Initially nil (not set)
	if rPr.BoldVal() != nil {
		t.Error("expected nil bold for new rPr")
	}

	// Set true → <w:b/> (no val attr)
	bTrue := true
	if err := rPr.SetBoldVal(&bTrue); err != nil {
		t.Fatalf("SetBoldVal: %v", err)
	}
	got := rPr.BoldVal()
	if got == nil || !*got {
		t.Error("expected *true after SetBoldVal(true)")
	}

	// Set false → <w:b w:val="false"/>
	bFalse := false
	if err := rPr.SetBoldVal(&bFalse); err != nil {
		t.Fatalf("SetBoldVal: %v", err)
	}
	got = rPr.BoldVal()
	if got == nil || *got {
		t.Error("expected *false after SetBoldVal(false)")
	}

	// Set nil → remove element
	if err := rPr.SetBoldVal(nil); err != nil {
		t.Fatalf("SetBoldVal: %v", err)
	}
	if rPr.BoldVal() != nil {
		t.Error("expected nil after SetBoldVal(nil)")
	}
}

func TestCT_RPr_ItalicVal_TriState(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	v := true
	if err := rPr.SetItalicVal(&v); err != nil {
		t.Fatalf("SetItalicVal: %v", err)
	}
	got := rPr.ItalicVal()
	if got == nil || !*got {
		t.Error("expected *true for italic")
	}
}

func TestCT_RPr_ColorVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	if cv, err := rPr.ColorVal(); err != nil {
		t.Fatalf("ColorVal: %v", err)
	} else if cv != nil {
		t.Error("expected nil color for new rPr")
	}

	c := "FF0000"
	if err := rPr.SetColorVal(&c); err != nil {
		t.Fatalf("SetColorVal: %v", err)
	}
	got, err := rPr.ColorVal()
	if err != nil {
		t.Fatalf("ColorVal: %v", err)
	}
	if got == nil || *got != "FF0000" {
		t.Errorf("expected FF0000, got %v", got)
	}

	if err := rPr.SetColorVal(nil); err != nil {
		t.Fatalf("SetColorVal: %v", err)
	}
	if cv, err := rPr.ColorVal(); err != nil {
		t.Fatalf("ColorVal: %v", err)
	} else if cv != nil {
		t.Error("expected nil after removing color")
	}
}

func TestCT_RPr_SzVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	if sv, err := rPr.SzVal(); err != nil {
		t.Fatalf("SzVal: %v", err)
	} else if sv != nil {
		t.Error("expected nil sz for new rPr")
	}

	var sz int64 = 24 // 12pt in half-points
	if err := rPr.SetSzVal(&sz); err != nil {
		t.Fatalf("SetSzVal: %v", err)
	}
	got, err := rPr.SzVal()
	if err != nil {
		t.Fatalf("SzVal: %v", err)
	}
	if got == nil || *got != 24 {
		t.Errorf("expected 24, got %v", got)
	}

	if err := rPr.SetSzVal(nil); err != nil {
		t.Fatalf("SetSzVal: %v", err)
	}
	if sv, err := rPr.SzVal(); err != nil {
		t.Fatalf("SzVal: %v", err)
	} else if sv != nil {
		t.Error("expected nil after removing sz")
	}
}

func TestCT_RPr_RFontsAscii(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	if rPr.RFontsAscii() != nil {
		t.Error("expected nil font for new rPr")
	}

	f := "Arial"
	if err := rPr.SetRFontsAscii(&f); err != nil {
		t.Fatalf("SetRFontsAscii: %v", err)
	}
	got := rPr.RFontsAscii()
	if got == nil || *got != "Arial" {
		t.Errorf("expected Arial, got %v", got)
	}

	if err := rPr.SetRFontsAscii(nil); err != nil {
		t.Fatalf("SetRFontsAscii: %v", err)
	}
	if rPr.RFontsAscii() != nil {
		t.Error("expected nil after removing font")
	}
}

func TestCT_RPr_StyleVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	s := "CommentReference"
	if err := rPr.SetStyleVal(&s); err != nil {
		t.Fatalf("SetStyleVal: %v", err)
	}
	got, err := rPr.StyleVal()
	if err != nil {
		t.Fatalf("StyleVal: %v", err)
	}
	if got == nil || *got != "CommentReference" {
		t.Errorf("expected CommentReference, got %v", got)
	}

	if err := rPr.SetStyleVal(nil); err != nil {
		t.Fatalf("SetStyleVal: %v", err)
	}
	if sv, err := rPr.StyleVal(); err != nil {
		t.Fatalf("StyleVal: %v", err)
	} else if sv != nil {
		t.Error("expected nil after removing style")
	}
}

func TestCT_RPr_UVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	if rPr.UVal() != nil {
		t.Error("expected nil underline for new rPr")
	}

	u := "single"
	if err := rPr.SetUVal(&u); err != nil {
		t.Fatalf("SetUVal: %v", err)
	}
	got := rPr.UVal()
	if got == nil || *got != "single" {
		t.Errorf("expected single, got %v", got)
	}

	if err := rPr.SetUVal(nil); err != nil {
		t.Fatalf("SetUVal: %v", err)
	}
	if rPr.UVal() != nil {
		t.Error("expected nil after removing underline")
	}
}

func TestCT_RPr_Subscript(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{e: rPrEl}}

	if sub, err := rPr.Subscript(); err != nil {
		t.Fatalf("Subscript: %v", err)
	} else if sub != nil {
		t.Error("expected nil subscript for new rPr")
	}

	bTrue := true
	if err := rPr.SetSubscript(&bTrue); err != nil {
		t.Fatalf("SetSubscript: %v", err)
	}
	got, err := rPr.Subscript()
	if err != nil {
		t.Fatalf("Subscript: %v", err)
	}
	if got == nil || !*got {
		t.Error("expected *true for subscript")
	}

	bFalse := false
	if err := rPr.SetSubscript(&bFalse); err != nil {
		t.Fatalf("SetSubscript: %v", err)
	}
	// Should remove since current is subscript and setting to false
	if sub, err := rPr.Subscript(); err != nil {
		t.Fatalf("Subscript: %v", err)
	} else if sub != nil {
		t.Error("expected nil after setting subscript to false (was subscript)")
	}
}

// ===========================================================================
// CT_RPr tri-state boolean properties — table-driven test
// ===========================================================================

// triStateProp describes a getter/setter pair for testing.
type triStateProp struct {
	name   string
	get    func(rPr *CT_RPr) *bool
	set    func(rPr *CT_RPr, v *bool) error
}

func TestCT_RPr_TriStateBooleans(t *testing.T) {
	t.Parallel()

	props := []triStateProp{
		{"Caps", (*CT_RPr).CapsVal, (*CT_RPr).SetCapsVal},
		{"SmallCaps", (*CT_RPr).SmallCapsVal, (*CT_RPr).SetSmallCapsVal},
		{"Strike", (*CT_RPr).StrikeVal, (*CT_RPr).SetStrikeVal},
		{"Dstrike", (*CT_RPr).DstrikeVal, (*CT_RPr).SetDstrikeVal},
		{"Outline", (*CT_RPr).OutlineVal, (*CT_RPr).SetOutlineVal},
		{"Shadow", (*CT_RPr).ShadowVal, (*CT_RPr).SetShadowVal},
		{"Emboss", (*CT_RPr).EmbossVal, (*CT_RPr).SetEmbossVal},
		{"Imprint", (*CT_RPr).ImprintVal, (*CT_RPr).SetImprintVal},
		{"NoProof", (*CT_RPr).NoProofVal, (*CT_RPr).SetNoProofVal},
		{"SnapToGrid", (*CT_RPr).SnapToGridVal, (*CT_RPr).SetSnapToGridVal},
		{"Vanish", (*CT_RPr).VanishVal, (*CT_RPr).SetVanishVal},
		{"WebHidden", (*CT_RPr).WebHiddenVal, (*CT_RPr).SetWebHiddenVal},
		{"SpecVanish", (*CT_RPr).SpecVanishVal, (*CT_RPr).SetSpecVanishVal},
		{"OMath", (*CT_RPr).OMathVal, (*CT_RPr).SetOMathVal},
	}

	for _, p := range props {
		p := p
		t.Run(p.name, func(t *testing.T) {
			t.Parallel()
			rPr := &CT_RPr{Element{e: OxmlElement("w:rPr")}}

			// Initially nil
			if p.get(rPr) != nil {
				t.Errorf("%s: expected nil initially", p.name)
			}

			// Set true
			bTrue := true
			if err := p.set(rPr, &bTrue); err != nil {
				t.Fatalf("%s: set true: %v", p.name, err)
			}
			got := p.get(rPr)
			if got == nil || !*got {
				t.Errorf("%s: expected *true, got %v", p.name, got)
			}

			// Set false
			bFalse := false
			if err := p.set(rPr, &bFalse); err != nil {
				t.Fatalf("%s: set false: %v", p.name, err)
			}
			got = p.get(rPr)
			if got == nil || *got {
				t.Errorf("%s: expected *false, got %v", p.name, got)
			}

			// Set nil (remove)
			if err := p.set(rPr, nil); err != nil {
				t.Fatalf("%s: set nil: %v", p.name, err)
			}
			if p.get(rPr) != nil {
				t.Errorf("%s: expected nil after removal", p.name)
			}
		})
	}
}

// ===========================================================================
// CT_RPr additional property tests
// ===========================================================================

func TestCT_RPr_ColorTheme_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{e: OxmlElement("w:rPr")}}

	// Initially nil
	if ct, err := rPr.ColorTheme(); err != nil {
		t.Fatalf("ColorTheme: %v", err)
	} else if ct != nil {
		t.Error("expected nil color theme initially")
	}

	// Set theme
	tc := enum.MsoThemeColorIndexAccent1
	if err := rPr.SetColorTheme(&tc); err != nil {
		t.Fatalf("SetColorTheme: %v", err)
	}
	got, err := rPr.ColorTheme()
	if err != nil {
		t.Fatalf("ColorTheme: %v", err)
	}
	if got == nil || *got != enum.MsoThemeColorIndexAccent1 {
		t.Errorf("expected Accent1, got %v", got)
	}

	// Remove
	if err := rPr.SetColorTheme(nil); err != nil {
		t.Fatalf("SetColorTheme nil: %v", err)
	}
	got, _ = rPr.ColorTheme()
	if got != nil {
		t.Error("expected nil after removing theme color")
	}
}

func TestCT_RPr_HighlightVal_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{e: OxmlElement("w:rPr")}}

	if hv, err := rPr.HighlightVal(); err != nil {
		t.Fatalf("HighlightVal: %v", err)
	} else if hv != nil {
		t.Error("expected nil highlight initially")
	}

	h := "yellow"
	if err := rPr.SetHighlightVal(&h); err != nil {
		t.Fatalf("SetHighlightVal: %v", err)
	}
	got, err := rPr.HighlightVal()
	if err != nil {
		t.Fatalf("HighlightVal: %v", err)
	}
	if got == nil || *got != "yellow" {
		t.Errorf("expected yellow, got %v", got)
	}

	if err := rPr.SetHighlightVal(nil); err != nil {
		t.Fatalf("SetHighlightVal nil: %v", err)
	}
	if hv, _ := rPr.HighlightVal(); hv != nil {
		t.Error("expected nil after removing highlight")
	}
}

func TestCT_RPr_Superscript_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{e: OxmlElement("w:rPr")}}

	// Initially nil
	if sup, err := rPr.Superscript(); err != nil {
		t.Fatalf("Superscript: %v", err)
	} else if sup != nil {
		t.Error("expected nil initially")
	}

	// Set true
	bTrue := true
	if err := rPr.SetSuperscript(&bTrue); err != nil {
		t.Fatalf("SetSuperscript true: %v", err)
	}
	got, err := rPr.Superscript()
	if err != nil {
		t.Fatalf("Superscript: %v", err)
	}
	if got == nil || !*got {
		t.Error("expected *true for superscript")
	}

	// Set false clears only if currently superscript
	bFalse := false
	if err := rPr.SetSuperscript(&bFalse); err != nil {
		t.Fatalf("SetSuperscript false: %v", err)
	}
	if sup, _ := rPr.Superscript(); sup != nil {
		t.Error("expected nil after false (was superscript)")
	}
}

func TestCT_RPr_RFontsHAnsi_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{e: OxmlElement("w:rPr")}}

	if rPr.RFontsHAnsi() != nil {
		t.Error("expected nil hAnsi initially")
	}

	f := "Times New Roman"
	if err := rPr.SetRFontsHAnsi(&f); err != nil {
		t.Fatalf("SetRFontsHAnsi: %v", err)
	}
	got := rPr.RFontsHAnsi()
	if got == nil || *got != "Times New Roman" {
		t.Errorf("expected Times New Roman, got %v", got)
	}

	// Set nil
	if err := rPr.SetRFontsHAnsi(nil); err != nil {
		t.Fatalf("SetRFontsHAnsi nil: %v", err)
	}
}
