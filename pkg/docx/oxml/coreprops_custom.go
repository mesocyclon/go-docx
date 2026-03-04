package oxml

import (
	"fmt"
	"regexp"
	"strconv"
	"time"
)

// ===========================================================================
// CT_CoreProperties — custom methods
// ===========================================================================

// NewCoreProperties creates a new empty <cp:coreProperties> element.
func NewCoreProperties() (*CT_CoreProperties, error) {
	xml := `<cp:coreProperties ` +
		`xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
		`xmlns:dcterms="http://purl.org/dc/terms/"/>`
	el, err := ParseXml([]byte(xml))
	if err != nil {
		return nil, fmt.Errorf("oxml: failed to parse coreProperties XML: %w", err)
	}
	return &CT_CoreProperties{Element{e: el}}, nil
}

// --- Text property helpers ---

// textOfElement returns the text content of the child element found by the generated
// accessor method (e.g. cp.Title(), cp.Creator()). Returns "" if element is absent or empty.
func (cp *CT_CoreProperties) textOfElement(el *CT_CorePropText) string {
	if el == nil {
		return ""
	}
	t := el.e.Text()
	if t == "" {
		return ""
	}
	return t
}

// setElementText sets the text of a child element via get-or-add, enforcing the 255 char limit.
func (cp *CT_CoreProperties) setElementText(getOrAdd func() *CT_CorePropText, value string) error {
	if len(value) > 255 {
		return fmt.Errorf("exceeded 255 char limit for core property, got %d chars", len(value))
	}
	el := getOrAdd()
	el.e.SetText(value)
	return nil
}

// --- Text properties ---

// TitleText returns the title (dc:title) or "".
func (cp *CT_CoreProperties) TitleText() string {
	return cp.textOfElement(cp.Title())
}

// SetTitleText sets the title (dc:title).
func (cp *CT_CoreProperties) SetTitleText(v string) error {
	return cp.setElementText(cp.GetOrAddTitle, v)
}

// AuthorText returns the author (dc:creator) or "".
func (cp *CT_CoreProperties) AuthorText() string {
	return cp.textOfElement(cp.Creator())
}

// SetAuthorText sets the author (dc:creator).
func (cp *CT_CoreProperties) SetAuthorText(v string) error {
	return cp.setElementText(cp.GetOrAddCreator, v)
}

// SubjectText returns the subject (dc:subject) or "".
func (cp *CT_CoreProperties) SubjectText() string {
	return cp.textOfElement(cp.Subject())
}

// SetSubjectText sets the subject (dc:subject).
func (cp *CT_CoreProperties) SetSubjectText(v string) error {
	return cp.setElementText(cp.GetOrAddSubject, v)
}

// CategoryText returns the category (cp:category) or "".
func (cp *CT_CoreProperties) CategoryText() string {
	return cp.textOfElement(cp.Category())
}

// SetCategoryText sets the category (cp:category).
func (cp *CT_CoreProperties) SetCategoryText(v string) error {
	return cp.setElementText(cp.GetOrAddCategory, v)
}

// KeywordsText returns the keywords (cp:keywords) or "".
func (cp *CT_CoreProperties) KeywordsText() string {
	return cp.textOfElement(cp.Keywords())
}

// SetKeywordsText sets the keywords (cp:keywords).
func (cp *CT_CoreProperties) SetKeywordsText(v string) error {
	return cp.setElementText(cp.GetOrAddKeywords, v)
}

// CommentsText returns the comments/description (dc:description) or "".
func (cp *CT_CoreProperties) CommentsText() string {
	return cp.textOfElement(cp.Description())
}

// SetCommentsText sets the comments/description (dc:description).
func (cp *CT_CoreProperties) SetCommentsText(v string) error {
	return cp.setElementText(cp.GetOrAddDescription, v)
}

// LastModifiedByText returns the last modified by (cp:lastModifiedBy) or "".
func (cp *CT_CoreProperties) LastModifiedByText() string {
	return cp.textOfElement(cp.LastModifiedBy())
}

// SetLastModifiedByText sets the last modified by (cp:lastModifiedBy).
func (cp *CT_CoreProperties) SetLastModifiedByText(v string) error {
	return cp.setElementText(cp.GetOrAddLastModifiedBy, v)
}

// ContentStatusText returns the content status (cp:contentStatus) or "".
func (cp *CT_CoreProperties) ContentStatusText() string {
	return cp.textOfElement(cp.ContentStatus())
}

// SetContentStatusText sets the content status (cp:contentStatus).
func (cp *CT_CoreProperties) SetContentStatusText(v string) error {
	return cp.setElementText(cp.GetOrAddContentStatus, v)
}

// IdentifierText returns the identifier (dc:identifier) or "".
func (cp *CT_CoreProperties) IdentifierText() string {
	return cp.textOfElement(cp.Identifier())
}

// SetIdentifierText sets the identifier (dc:identifier).
func (cp *CT_CoreProperties) SetIdentifierText(v string) error {
	return cp.setElementText(cp.GetOrAddIdentifier, v)
}

// LanguageText returns the language (dc:language) or "".
func (cp *CT_CoreProperties) LanguageText() string {
	return cp.textOfElement(cp.Language())
}

// SetLanguageText sets the language (dc:language).
func (cp *CT_CoreProperties) SetLanguageText(v string) error {
	return cp.setElementText(cp.GetOrAddLanguage, v)
}

// VersionText returns the version (cp:version) or "".
func (cp *CT_CoreProperties) VersionText() string {
	return cp.textOfElement(cp.Version())
}

// SetVersionText sets the version (cp:version).
func (cp *CT_CoreProperties) SetVersionText(v string) error {
	return cp.setElementText(cp.GetOrAddVersion, v)
}

// --- Datetime properties ---

var w3cdtfTemplates = []string{
	"2006-01-02T15:04:05",
	"2006-01-02",
	"2006-01",
	"2006",
}

var offsetPattern = regexp.MustCompile(`([+-])(\d{2}):(\d{2})`)

// parseW3CDTF parses a W3CDTF datetime string into a time.Time in UTC.
// Supports formats: yyyy, yyyy-mm, yyyy-mm-dd, yyyy-mm-ddThh:mm:ssZ,
// yyyy-mm-ddThh:mm:ss±hh:mm
func parseW3CDTF(s string) (*time.Time, error) {
	if s == "" {
		return nil, fmt.Errorf("empty datetime string")
	}

	// Extract parseable part (up to 19 chars) and timezone offset
	parseablePart := s
	offsetStr := ""
	if len(s) > 19 {
		parseablePart = s[:19]
		offsetStr = s[19:]
	}

	var dt time.Time
	var parsed bool
	for _, tmpl := range w3cdtfTemplates {
		t, err := time.Parse(tmpl, parseablePart)
		if err == nil {
			dt = t
			parsed = true
			break
		}
	}
	if !parsed {
		return nil, fmt.Errorf("could not parse W3CDTF datetime string '%s'", s)
	}

	// Apply timezone offset if present
	if len(offsetStr) == 6 {
		match := offsetPattern.FindStringSubmatch(offsetStr)
		if match == nil {
			return nil, fmt.Errorf("'%s' is not a valid offset string", offsetStr)
		}
		sign := match[1]
		hours, _ := strconv.Atoi(match[2])
		minutes, _ := strconv.Atoi(match[3])

		// Reverse the sign: +07:00 means UTC-7, so we subtract
		signFactor := -1
		if sign == "-" {
			signFactor = 1
		}

		dt = dt.Add(time.Duration(signFactor*hours)*time.Hour + time.Duration(signFactor*minutes)*time.Minute)
	}

	utc := dt.UTC()
	return &utc, nil
}

// datetimeOfElement reads a datetime from a child element, or nil.
func (cp *CT_CoreProperties) datetimeOfElement(el *CT_CorePropText) (*time.Time, error) {
	if el == nil {
		return nil, nil
	}
	text := el.e.Text()
	if text == "" {
		return nil, nil
	}
	dt, err := parseW3CDTF(text)
	if err != nil {
		return nil, fmt.Errorf("oxml: parsing datetime %q: %w", text, err)
	}
	return dt, nil
}

// setElementDatetime sets a datetime on a child element, formatting as W3CDTF.
// For "created" and "modified", also sets the xsi:type attribute.
func (cp *CT_CoreProperties) setElementDatetime(getOrAdd func() *CT_CorePropText, t time.Time, isDateTerms bool) {
	el := getOrAdd()
	dtStr := t.UTC().Format("2006-01-02T15:04:05Z")
	el.e.SetText(dtStr)

	if isDateTerms {
		// dcterms:created and dcterms:modified require xsi:type="dcterms:W3CDTF"
		cp.ensureXsiNamespace()
		el.e.CreateAttr("xsi:type", "dcterms:W3CDTF")
	}
}

// ensureXsiNamespace ensures the xsi namespace declaration is present on the
// core properties root element. This is needed for dcterms:created and dcterms:modified
// which require xsi:type="dcterms:W3CDTF".
func (cp *CT_CoreProperties) ensureXsiNamespace() {
	const xsiURI = "http://www.w3.org/2001/XMLSchema-instance"
	// Check if xmlns:xsi is already declared
	if _, ok := HasNsDecl(cp.e, "xsi"); !ok {
		cp.e.CreateAttr("xmlns:xsi", xsiURI)
	}
}

// CreatedDatetime returns the created datetime (dcterms:created) or nil.
func (cp *CT_CoreProperties) CreatedDatetime() (*time.Time, error) {
	return cp.datetimeOfElement(cp.Created())
}

// SetCreatedDatetime sets the created datetime (dcterms:created).
func (cp *CT_CoreProperties) SetCreatedDatetime(t time.Time) {
	cp.setElementDatetime(cp.GetOrAddCreated, t, true)
}

// ModifiedDatetime returns the modified datetime (dcterms:modified) or nil.
func (cp *CT_CoreProperties) ModifiedDatetime() (*time.Time, error) {
	return cp.datetimeOfElement(cp.Modified())
}

// SetModifiedDatetime sets the modified datetime (dcterms:modified).
func (cp *CT_CoreProperties) SetModifiedDatetime(t time.Time) {
	cp.setElementDatetime(cp.GetOrAddModified, t, true)
}

// LastPrintedDatetime returns the last printed datetime (cp:lastPrinted) or nil.
func (cp *CT_CoreProperties) LastPrintedDatetime() (*time.Time, error) {
	return cp.datetimeOfElement(cp.LastPrinted())
}

// SetLastPrintedDatetime sets the last printed datetime (cp:lastPrinted).
func (cp *CT_CoreProperties) SetLastPrintedDatetime(t time.Time) {
	cp.setElementDatetime(cp.GetOrAddLastPrinted, t, false)
}

// --- Revision property ---

// RevisionNumber returns the integer value of the revision (cp:revision), or 0
// if not present, not a valid integer, or negative.
func (cp *CT_CoreProperties) RevisionNumber() int {
	rev := cp.Revision()
	if rev == nil {
		return 0
	}
	text := rev.e.Text()
	if text == "" {
		return 0
	}
	v, err := strconv.Atoi(text)
	if err != nil || v < 0 {
		return 0
	}
	return v
}

// SetRevisionNumber sets the revision number (cp:revision).
// Value must be a positive integer (>= 1).
func (cp *CT_CoreProperties) SetRevisionNumber(v int) error {
	if v < 1 {
		return fmt.Errorf("revision property requires positive int, got %d", v)
	}
	rev := cp.GetOrAddRevision()
	rev.e.SetText(strconv.Itoa(v))
	return nil
}
