package parts

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// Mirrors Python: it_provides_access_to_its_settings
func TestSettingsPart_SettingsElement(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`)
	pkg := opc.NewOpcPackage(nil)
	xp, err := opc.NewXmlPart("/word/settings.xml", opc.CTWmlSettings, blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	sp := NewSettingsPart(xp)

	settings, err := sp.SettingsElement()
	if err != nil {
		t.Fatalf("SettingsElement(): %v", err)
	}
	if settings == nil {
		t.Fatal("SettingsElement() returned nil")
	}
}

// Mirrors Python: it_constructs_a_default_settings_part_to_help
func TestSettingsPart_Default(t *testing.T) {
	pkg := opc.NewOpcPackage(nil)
	sp, err := DefaultSettingsPart(pkg)
	if err != nil {
		t.Fatalf("DefaultSettingsPart: %v", err)
	}
	if sp == nil {
		t.Fatal("DefaultSettingsPart returned nil")
	}
	if sp.Element() == nil {
		t.Error("default SettingsPart has nil element")
	}
	settings, err := sp.SettingsElement()
	if err != nil {
		t.Fatalf("default SettingsElement(): %v", err)
	}
	if settings == nil {
		t.Fatal("default SettingsElement() returned nil")
	}
}

func TestLoadSettingsPart_ReturnsSettingsPart(t *testing.T) {
	blob := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg := opc.NewOpcPackage(nil)
	part, err := LoadSettingsPart("/word/settings.xml", opc.CTWmlSettings, "", blob, pkg)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := part.(*SettingsPart); !ok {
		t.Errorf("LoadSettingsPart returned %T, want *SettingsPart", part)
	}
}
