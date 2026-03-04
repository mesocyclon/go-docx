package oxml

import (
	"testing"
)

func TestCT_LastRenderedPageBreak_PrecedesAllContent(t *testing.T) {
	// Build: <w:p><w:r><w:lastRenderedPageBreak/><w:t>text</w:t></w:r></w:p>
	pEl := OxmlElement("w:p")
	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)
	tEl := OxmlElement("w:t")
	tEl.SetText("text")
	rEl.AddChild(tEl)

	lrpb := &CT_LastRenderedPageBreak{Element{e: lrpbEl}}

	if !lrpb.PrecedesAllContent() {
		t.Error("expected PrecedesAllContent to be true when lrpb is first in first run")
	}
}

func TestCT_LastRenderedPageBreak_FollowsAllContent(t *testing.T) {
	// Build: <w:p><w:r><w:t>text</w:t><w:lastRenderedPageBreak/></w:r></w:p>
	pEl := OxmlElement("w:p")
	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)
	tEl := OxmlElement("w:t")
	tEl.SetText("text")
	rEl.AddChild(tEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	lrpb := &CT_LastRenderedPageBreak{Element{e: lrpbEl}}

	if !lrpb.FollowsAllContent() {
		t.Error("expected FollowsAllContent to be true when lrpb is last in last run")
	}
}

func TestCT_LastRenderedPageBreak_IsInHyperlink(t *testing.T) {
	// Build: <w:p><w:hyperlink><w:r><w:lastRenderedPageBreak/></w:r></w:hyperlink></w:p>
	pEl := OxmlElement("w:p")
	hEl := OxmlElement("w:hyperlink")
	pEl.AddChild(hEl)
	rEl := OxmlElement("w:r")
	hEl.AddChild(rEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	lrpb := &CT_LastRenderedPageBreak{Element{e: lrpbEl}}

	if !lrpb.IsInHyperlink() {
		t.Error("expected IsInHyperlink to be true")
	}
}

func TestCT_LastRenderedPageBreak_EnclosingP(t *testing.T) {
	pEl := OxmlElement("w:p")
	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	lrpb := &CT_LastRenderedPageBreak{Element{e: lrpbEl}}
	p := lrpb.EnclosingP()
	if p == nil || p.e != pEl {
		t.Error("EnclosingP should return the parent w:p")
	}
}

// ===========================================================================
// CT_LastRenderedPageBreak fragmentation tests
// ===========================================================================

// buildParagraphWithBreakInRun builds:
// <w:p><w:pPr/><w:r><w:t>before</w:t><w:lastRenderedPageBreak/><w:t>after</w:t></w:r></w:p>
func buildParagraphWithBreakInRun() (*CT_P, *CT_LastRenderedPageBreak) {
	pEl := OxmlElement("w:p")
	pPrEl := OxmlElement("w:pPr")
	pEl.AddChild(pPrEl)

	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)

	t1 := OxmlElement("w:t")
	t1.SetText("before")
	rEl.AddChild(t1)

	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	t2 := OxmlElement("w:t")
	t2.SetText("after")
	rEl.AddChild(t2)

	return &CT_P{Element{e: pEl}}, &CT_LastRenderedPageBreak{Element{e: lrpbEl}}
}

func TestCT_LastRenderedPageBreak_PrecedingFragment_InRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInRun()

	frag, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP: %v", err)
	}

	// The preceding fragment should contain "before" but not "after"
	text := frag.ParagraphText()
	if text != "before" {
		t.Errorf("PrecedingFragmentP text: got %q, want %q", text, "before")
	}
}

func TestCT_LastRenderedPageBreak_FollowingFragment_InRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInRun()

	frag, err := lrpb.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP: %v", err)
	}

	text := frag.ParagraphText()
	if text != "after" {
		t.Errorf("FollowingFragmentP text: got %q, want %q", text, "after")
	}
}

// buildParagraphWithBreakInHyperlink builds:
// <w:p><w:pPr/><w:r><w:t>pre</w:t></w:r>
//   <w:hyperlink><w:r><w:lastRenderedPageBreak/><w:t>link</w:t></w:r></w:hyperlink>
//   <w:r><w:t>post</w:t></w:r></w:p>
func buildParagraphWithBreakInHyperlink() (*CT_P, *CT_LastRenderedPageBreak) {
	pEl := OxmlElement("w:p")
	pPrEl := OxmlElement("w:pPr")
	pEl.AddChild(pPrEl)

	r1 := OxmlElement("w:r")
	t1 := OxmlElement("w:t")
	t1.SetText("pre")
	r1.AddChild(t1)
	pEl.AddChild(r1)

	hEl := OxmlElement("w:hyperlink")
	hr := OxmlElement("w:r")
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	hr.AddChild(lrpbEl)
	tLink := OxmlElement("w:t")
	tLink.SetText("link")
	hr.AddChild(tLink)
	hEl.AddChild(hr)
	pEl.AddChild(hEl)

	r2 := OxmlElement("w:r")
	t2 := OxmlElement("w:t")
	t2.SetText("post")
	r2.AddChild(t2)
	pEl.AddChild(r2)

	return &CT_P{Element{e: pEl}}, &CT_LastRenderedPageBreak{Element{e: lrpbEl}}
}

func TestCT_LastRenderedPageBreak_PrecedingFragment_InHyperlink(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInHyperlink()

	frag, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP (hyperlink): %v", err)
	}

	// Preceding should include pre-run and the hyperlink (without lrpb),
	// but not the post-run
	text := frag.ParagraphText()
	if text != "prelink" {
		t.Errorf("PrecedingFragmentP text: got %q, want %q", text, "prelink")
	}
}

func TestCT_LastRenderedPageBreak_FollowingFragment_InHyperlink(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInHyperlink()

	frag, err := lrpb.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP (hyperlink): %v", err)
	}

	// Following should include content after the hyperlink
	text := frag.ParagraphText()
	if text != "post" {
		t.Errorf("FollowingFragmentP text: got %q, want %q", text, "post")
	}
}

func TestCT_LastRenderedPageBreak_Fragment_PreservesProperties(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInRun()

	// pPr should survive in both fragments
	preceding, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP: %v", err)
	}
	if preceding.e.FindElement("w:pPr") == nil {
		t.Error("pPr should be preserved in preceding fragment")
	}

	_, lrpb2 := buildParagraphWithBreakInRun()
	following, err := lrpb2.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP: %v", err)
	}
	if following.e.FindElement("w:pPr") == nil {
		t.Error("pPr should be preserved in following fragment")
	}
}

// buildMultiRunParagraphWithBreak builds:
// <w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t><w:lastRenderedPageBreak/><w:t>C</w:t></w:r><w:r><w:t>D</w:t></w:r></w:p>
func buildMultiRunParagraphWithBreak() (*CT_P, *CT_LastRenderedPageBreak) {
	pEl := OxmlElement("w:p")

	r1 := OxmlElement("w:r")
	t1 := OxmlElement("w:t")
	t1.SetText("A")
	r1.AddChild(t1)
	pEl.AddChild(r1)

	r2 := OxmlElement("w:r")
	t2 := OxmlElement("w:t")
	t2.SetText("B")
	r2.AddChild(t2)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	r2.AddChild(lrpbEl)
	t3 := OxmlElement("w:t")
	t3.SetText("C")
	r2.AddChild(t3)
	pEl.AddChild(r2)

	r3 := OxmlElement("w:r")
	t4 := OxmlElement("w:t")
	t4.SetText("D")
	r3.AddChild(t4)
	pEl.AddChild(r3)

	return &CT_P{Element{e: pEl}}, &CT_LastRenderedPageBreak{Element{e: lrpbEl}}
}

func TestCT_LastRenderedPageBreak_PrecedingFragment_MultiRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildMultiRunParagraphWithBreak()

	frag, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP: %v", err)
	}
	text := frag.ParagraphText()
	if text != "AB" {
		t.Errorf("PrecedingFragmentP (multi-run): got %q, want %q", text, "AB")
	}
}

func TestCT_LastRenderedPageBreak_FollowingFragment_MultiRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildMultiRunParagraphWithBreak()

	frag, err := lrpb.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP: %v", err)
	}
	text := frag.ParagraphText()
	if text != "CD" {
		t.Errorf("FollowingFragmentP (multi-run): got %q, want %q", text, "CD")
	}
}
