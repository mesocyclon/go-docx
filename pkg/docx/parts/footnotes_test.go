package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

func TestFootnotesPart_Existing(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireFootnotesPart(t, dp, pkg)

	got, err := dp.FootnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("FootnotesPart should return the wired part")
	}
}

func TestFootnotesPart_NotFound(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	_, err := dp.FootnotesPart()
	if err == nil {
		t.Error("expected error when no footnotes part exists")
	}
}

func TestGetOrAddFootnotesPart_CreatesWhenMissing(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	fp, err := dp.GetOrAddFootnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if fp == nil {
		t.Fatal("GetOrAddFootnotesPart returned nil")
	}
	if fp.Element() == nil {
		t.Error("default footnotes part element is nil")
	}
}

func TestGetOrAddFootnotesPart_ReturnsExisting(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireFootnotesPart(t, dp, pkg)

	got, err := dp.GetOrAddFootnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("GetOrAddFootnotesPart should return existing part")
	}
}

func TestEndnotesPart_Existing(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireEndnotesPart(t, dp, pkg)

	got, err := dp.EndnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("EndnotesPart should return the wired part")
	}
}

func TestEndnotesPart_NotFound(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	_, err := dp.EndnotesPart()
	if err == nil {
		t.Error("expected error when no endnotes part exists")
	}
}

func TestGetOrAddEndnotesPart_CreatesWhenMissing(t *testing.T) {
	dp, _ := newDocPartWithBody(t)
	ep, err := dp.GetOrAddEndnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if ep == nil {
		t.Fatal("GetOrAddEndnotesPart returned nil")
	}
	if ep.Element() == nil {
		t.Error("default endnotes part element is nil")
	}
}

func TestGetOrAddEndnotesPart_ReturnsExisting(t *testing.T) {
	dp, pkg := newDocPartWithBody(t)
	wired := wireEndnotesPart(t, dp, pkg)

	got, err := dp.GetOrAddEndnotesPart()
	if err != nil {
		t.Fatal(err)
	}
	if got != wired {
		t.Error("GetOrAddEndnotesPart should return existing part")
	}
}

func TestDefaultFootnotesPart_Creates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	fp, err := DefaultFootnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if fp.Element() == nil {
		t.Fatal("default footnotes part element is nil")
	}
	if fp.PartName() != "/word/footnotes.xml" {
		t.Errorf("partname = %q, want /word/footnotes.xml", fp.PartName())
	}
}

func TestDefaultEndnotesPart_Creates(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	ep, err := DefaultEndnotesPart(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if ep.Element() == nil {
		t.Fatal("default endnotes part element is nil")
	}
	if ep.PartName() != "/word/endnotes.xml" {
		t.Errorf("partname = %q, want /word/endnotes.xml", ep.PartName())
	}
}
