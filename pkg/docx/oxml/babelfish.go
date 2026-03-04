package oxml

// --------------------------------------------------------------------------
// BabelFish — style name UI↔internal translation
// --------------------------------------------------------------------------
//
// Translates special-case style names between their UI form (e.g. "Heading 1")
// and the internal/styles.xml form (e.g. "heading 1").
//
// Only the built-in styles listed below have aliases. User-defined (custom)
// styles are not translated and must be referenced by their exact name as it
// appears in styles.xml.
//
// Lives in oxml because it's a static mapping table used by CT_Styles methods.
// Mirrors Python docx.styles.BabelFish exactly.

var babelFishAliases = [][2]string{
	{"Caption", "caption"},
	{"Footer", "footer"},
	{"Header", "header"},
	{"Heading 1", "heading 1"},
	{"Heading 2", "heading 2"},
	{"Heading 3", "heading 3"},
	{"Heading 4", "heading 4"},
	{"Heading 5", "heading 5"},
	{"Heading 6", "heading 6"},
	{"Heading 7", "heading 7"},
	{"Heading 8", "heading 8"},
	{"Heading 9", "heading 9"},
}

var (
	ui2internalMap = buildUI2InternalMap()
	internal2uiMap = buildInternal2UIMap()
)

func buildUI2InternalMap() map[string]string {
	m := make(map[string]string, len(babelFishAliases))
	for _, a := range babelFishAliases {
		m[a[0]] = a[1]
	}
	return m
}

func buildInternal2UIMap() map[string]string {
	m := make(map[string]string, len(babelFishAliases))
	for _, a := range babelFishAliases {
		m[a[1]] = a[0]
	}
	return m
}

// UI2Internal converts a UI style name to its internal/styles.xml form.
// Only built-in styles (Heading 1–9, Caption, Header, Footer) have mappings;
// all other names are returned unchanged.
//
// Mirrors Python BabelFish.ui2internal.
func UI2Internal(name string) string {
	if v, ok := ui2internalMap[name]; ok {
		return v
	}
	return name
}

// Internal2UI converts an internal/styles.xml name to its UI form.
// Only built-in styles (Heading 1–9, Caption, Header, Footer) have mappings;
// all other names are returned unchanged.
//
// Mirrors Python BabelFish.internal2ui.
func Internal2UI(name string) string {
	if v, ok := internal2uiMap[name]; ok {
		return v
	}
	return name
}
