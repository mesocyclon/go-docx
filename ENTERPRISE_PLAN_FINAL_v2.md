# ReplaceWithContent — Enterprise Plan (Final v4)

**Версия:** 4.0 (после code review и верификации)
**Метод:** `func (d *Document) ReplaceWithContent(old string, cd ContentData) (int, error)`
**Расположение:** `pkg/docx/document.go:558`
**Задача:** Привести метод к enterprise-уровню по модели Aspose.Words

---

## Содержание

1. [Текущее состояние и оценка](#1-текущее-состояние-и-оценка)
2. [Gap Analysis vs Aspose.Words](#2-gap-analysis)
3. [Факты об Aspose.Words (верифицированные)](#3-факты-об-aspose)
4. [Step 1 — Защита от порчи документа](#step-1)
5. [Step 2 — ContentData API (ImportFormatMode + Options)](#step-2)
6. [Step 3 — ImportFormatMode: 3 стратегии слияния стилей](#step-3)
7. [Step 4 — Expand to Direct Attributes](#step-4)
8. [Step 5 — ForceCopyStyles (переименование стилей)](#step-5)
9. [Step 6 — Deep Generic Part Import](#step-6)
10. [Step 7 — ImportFormatOptions (нумерация)](#step-7)
11. [Step 8 — Тесты (60+)](#step-8)
12. [Что НЕ нужно делать](#что-не-нужно-делать)
13. [Файловая карта изменений](#файловая-карта-изменений)
14. [Критерии приёмки](#критерии-приёмки)

---

<a id="1-текущее-состояние-и-оценка"></a>
## 1. Текущее состояние и оценка

### Архитектура (4 слоя)

```
Domain (pkg/docx)  →  Parts (parts)  →  OXML (oxml)  →  OPC (opc)
    Document              StoryPart        CT_P, CT_Tbl     Package
    Body                  DocumentPart     CT_Style          Part
    Paragraph             WmlPackage       CT_Numbering      Relationships
    Table                 ImagePart        CT_SectPr         PackURI
```

Зависимости строго однонаправлены. Все 4 слоя стабильны.

### Текущий pipeline ReplaceWithContent

```
ReplaceWithContent(old, cd)
│
├─ 1. newResourceImporter(source, target)         resourceimport.go:56
│
├─ 2. Phase 1 — import resources (once)
│  ├─ ri.importNumbering()                         resourceimport_num.go:25
│  │   └─ collect numIds → copy abstractNum (new ID + nsid) → copy num → numIdMap
│  ├─ ri.importStyles()                            resourceimport_styles.go:29
│  │   └─ collect styleIds → BFS closure (basedOn/next/link) → UseDestinationStyles merge → styleMap
│  ├─ ri.importFootnotes()                         resourceimport_notes.go:27
│  │   └─ collect refs → copy note (fresh ID) → sanitize → import styles → import rIds → renumber IDs
│  └─ ri.importEndnotes()                          resourceimport_notes.go:64
│
├─ 3. prepareContentElements(source, targetPart)   contentdata.go:265
│  ├─ extract body children (skip sectPr)
│  ├─ deep copy (el.Copy())
│  ├─ sanitizeForInsertion (17 annotation markers + paragraph sectPr)
│  ├─ materializeImplicitStyles (default style mismatch)
│  ├─ remapAll (numId, pStyle/rStyle/tblStyle, footnoteRef, endnoteRef)
│  ├─ collectReferencedRIds → importRelationship (3 cases) → remapRIds
│  └─ return preparedContent{elements}
│
├─ 4. Body replacement                             blkcntnr.go:375
│  └─ replaceTagWithElements(old, elementBuilder)
│      ├─ SplitParagraphAtTags → fragments
│      ├─ elementBuilder() → fresh Copy() + renumberDrawingIDs per placeholder
│      ├─ spliceElement (replace original paragraph)
│      └─ recurse into pre-existing table cells
│
├─ 5. Headers/Footers (6 per section, dedup by StoryPart pointer)
│  └─ per unique part: prepareContentElements → replaceWithContent
│
└─ 6. Comments
   └─ replaceWithContentInComments → same pipeline
```

### Что реализовано хорошо

| Возможность | Файл:строка | Зрелость |
|---|---|---|
| Style import с BFS closure (basedOn/next/link) | `resourceimport_styles.go:112-154` | Production |
| UseDestinationStyles стратегия | `resourceimport_styles.go:169-206` | Production |
| Материализация implicit default styles | `resourceimport_styles.go:352-408` | Production |
| Numbering import (fresh ID + nsid regen) | `resourceimport_num.go:25-103` | Production |
| Footnotes/endnotes с полным sub-pipeline | `resourceimport_notes.go:100-164` | Production |
| Image dedup (SHA-256) | `contentdata.go:162-192` | Production |
| External hyperlink import | `contentdata.go:150-154` | Production |
| Annotation sanitization (17 types) | `contentdata.go:359-441` | Production |
| Paragraph-level sectPr removal | `contentdata.go:422-425` | Production |
| Drawing ID renumbering (docPr/cNvPr) | `contentdata.go:459-473` | Production |
| Header/footer dedup by StoryPart* | `document.go:607-626` | Production |
| Table cell recursion | `blkcntnr.go:252-264` | Production |
| Cross-run tag detection | `oxml/replacetext.go` | Production |
| Source immutability (deep copy) | `contentdata.go:287-289` | Production |
| Fresh Copy() per placeholder | `blkcntnr.go:379-382` | Production |
| Idempotent import phases | `resourceimport.go:34,39,50,51` | Production |
| 25 integration tests | `replacecontent_test.go` | Good |

---

<a id="2-gap-analysis"></a>
## 2. Gap Analysis vs Aspose.Words

| # | Что отсутствует | Критичность | Aspose аналог |
|---|---|---|---|
| **G1** | ImportFormatMode — только UseDestinationStyles | HIGH | 3 режима + `ImportFormatOptions` |
| **G2** | Generic parts (charts, VML) — shallow copy, sub-rels теряются | HIGH | `NodeImporter` deep copy |
| **G3** | При ошибке document partially modified | HIGH | `Document.Clone()` before op |
| **G4** | Нет KeepSourceNumbering / MergePastedLists config | MEDIUM | `ImportFormatOptions` |
| **G5** | Нет ForceCopyStyles (rename with `_0` suffix) | MEDIUM | `ImportFormatOptions.ForceCopyStyles` |

---

<a id="3-факты-об-aspose"></a>
## 3. Верифицированные факты об Aspose.Words

Проверены по официальной документации и форумам Aspose. Критически важно для корректности плана.

### Aspose ДЕЛАЕТ

| Факт | Источник |
|---|---|
| 3 ImportFormatMode: UseDestinationStyles, KeepSourceFormatting, KeepDifferentStyles | [reference.aspose.com/words/net/aspose.words/importformatmode](https://reference.aspose.com/words/net/aspose.words/importformatmode/) |
| KeepSourceFormatting при конфликте: раскрывает formatting в **direct node attributes**, стиль меняется на Normal | Официальная документация: *"the source style formatting is expanded into direct Node attributes and the style is changed to Normal"* |
| ForceCopyStyles=true — только тогда копирует с суффиксами `_0`, `_1` | [ImportFormatOptions.ForceCopyStyles](https://reference.aspose.com/words/net/aspose.words/importformatoptions/forcecopystyles/) |
| KeepDifferentStyles: одинаковые стили → use target; разные → как KeepSourceFormatting | Документация |
| KeepSourceNumbering: true = отдельный список, false = сливается с target | [ImportFormatOptions.KeepSourceNumbering](https://reference.aspose.com/words/net/aspose.words/importformatoptions/keepsourcenumbering/) |
| Fields — сохраняются as-is, Word пересчитывает | *"all fields in a document are preserved during open/save and conversions"* |
| SDT — сохраняются as-is, data bindings могут ломаться | Документация по SDT |
| OLE — сохраняет blob as-is, не рекурсирует в embedded OLE | Документация по OLE |

### Aspose НЕ ДЕЛАЕТ

| Факт | Почему важно |
|---|---|
| **НЕ делает rollback/snapshot** | Рекомендует `Document.Clone()` до операции |
| **НЕ мержит theme** между документами | Один theme на документ |
| **НЕ мержит font table** при import | Resolution — задача рендера |
| **НЕ имеет Document.Validate()** | Post-import validation не существует |
| **НЕ ребиндит SDT data bindings** | Пользователь клонирует CustomXmlParts сам |
| **НЕ имеет progress callback** для import | Нет такого API |

---

<a id="step-1"></a>
## Step 1 — Защита от порчи документа при ошибке

**Приоритет:** CRITICAL
**Проблема:** `document.go:556-557` — *"On error, the document may be partially modified."*

### 1.1 Что именно может сломаться

При ошибке на любом этапе:
- **Phase 1 (import resources):** Стили/numbering уже добавлены в target `styles.xml` / `numbering.xml`, но body не модифицирован. Состояние: лишние стили (безвредно для Word).
- **Body replacement:** Часть параграфов заменена, часть нет. Состояние: **документ повреждён семантически**.
- **Headers/footers:** Body уже модифицирован, ошибка на 3-м header. Состояние: body OK, часть headers сломана.

Самый опасный случай — ошибка во время body replacement.

### 1.2 Решение: body snapshot через etree.Copy()

**Проверено:** `SetBlob()` существует на `BasePart` (`opc/part.go:63`). `etree.Element.Copy()` используется как deep copy по всему проекту.

**Верифицировано:** `CT_Document` embedding `oxml.Element`, который имеет `RawElement() *etree.Element` (`oxml/element.go:21`). `CT_Body` аналогично embedding `oxml.Element` с тем же `RawElement()`.

```go
// Файл: pkg/docx/document.go

// bodySnapshot captures the body element state before mutation.
type bodySnapshot struct {
    bodyEl *etree.Element // deep copy of <w:body>
}

// snapshotBody creates a deep copy of the current body element.
func (d *Document) snapshotBody() (*bodySnapshot, error) {
    b, err := d.getBody()
    if err != nil {
        return nil, err
    }
    return &bodySnapshot{bodyEl: b.Element().Copy()}, nil
}

// restoreBody replaces the current body with the snapshot.
//
// Verified API chain:
//   d.element → *oxml.CT_Document (embeds oxml.Element)
//   d.element.RawElement() → *etree.Element (<w:document>)
//   d.element.Body() → *oxml.CT_Body
//   body.RawElement() → *etree.Element (<w:body>)
func (d *Document) restoreBody(snap *bodySnapshot) {
    docEl := d.element.RawElement() // <w:document> *etree.Element
    oldBody := d.element.Body().RawElement()
    docEl.RemoveChild(oldBody)
    docEl.AddChild(snap.bodyEl)
    d.body = nil // invalidate cached Body proxy
}
```

### 1.3 Интеграция в ReplaceWithContent

```go
func (d *Document) ReplaceWithContent(old string, cd ContentData) (count int, err error) {
    if old == "" {
        return 0, nil
    }
    if cd.Source == nil {
        return 0, fmt.Errorf("docx: ContentData.Source is nil")
    }

    // Snapshot body before any mutation.
    snap, snapErr := d.snapshotBody()
    if snapErr != nil {
        return 0, fmt.Errorf("docx: creating body snapshot: %w", snapErr)
    }
    defer func() {
        if err != nil && snap != nil {
            d.restoreBody(snap)
        }
    }()

    // ... existing pipeline (phases 1-3, body, headers, comments) ...
}
```

### 1.4 Ограничения (документируем в godoc)

- **Orphan parts:** Если ошибка после `AddPart()` для image/generic part, part остаётся в `OpcPackage` — `RemovePart()` не существует. Orphan parts безвредны: Word их игнорирует.
- **Styles/numbering rollback:** Стили и numbering definitions добавленные в Phase 1 не откатываются. Это тоже безвредно — лишние стили не видны пользователю если не используются.
- **Полный rollback:** Для критических сценариев документируем паттерн save/reopen:

```go
// Для полного rollback (включая orphan parts):
buf, _ := doc.SaveToBytes()
count, err := doc.ReplaceWithContent(tag, cd)
if err != nil {
    doc, _ = OpenBytes(buf)
}
```

### 1.5 Файлы для изменения

| Файл | Что добавить |
|---|---|
| `pkg/docx/document.go` | `bodySnapshot`, `snapshotBody()`, `restoreBody()`, defer в ReplaceWithContent |

---

<a id="step-2"></a>
## Step 2 — ContentData API (ImportFormatMode + Options)

**Приоритет:** HIGH
**Принцип:** Zero-value = текущее поведение. Полная обратная совместимость.

### 2.1 Новые типы

```go
// Файл: pkg/docx/contentdata.go

// ImportFormatMode controls how styles are handled during content import.
type ImportFormatMode int

const (
    // UseDestinationStyles uses the destination document's styles when
    // a style with the same ID exists. Missing styles are copied from source.
    // This is the current (default) behavior.
    UseDestinationStyles ImportFormatMode = iota

    // KeepSourceFormatting preserves the source document's formatting.
    // On style ID conflict, source style properties are expanded into
    // direct paragraph/run attributes and the style reference is changed
    // to the target's default paragraph style.
    KeepSourceFormatting

    // KeepDifferentStyles is a hybrid: if conflicting styles have identical
    // formatting, uses the destination style. If formatting differs,
    // behaves like KeepSourceFormatting (expands to direct attributes).
    KeepDifferentStyles
)

// ImportFormatOptions provides fine-grained control over content import.
type ImportFormatOptions struct {
    // ForceCopyStyles forces conflicting styles to be copied with a
    // unique suffix (_0, _1, ...) instead of expanding to direct attributes.
    // Only effective with KeepSourceFormatting.
    // Mirrors Aspose.Words ImportFormatOptions.ForceCopyStyles.
    ForceCopyStyles bool

    // KeepSourceNumbering preserves source list numbering as a separate
    // list definition. When false (default), source lists merge into
    // matching target lists and continue their numbering.
    // Current project behavior is equivalent to KeepSourceNumbering=true.
    KeepSourceNumbering bool
}
```

### 2.2 Расширение ContentData

```go
type ContentData struct {
    // Source is the opened document whose body content will be inserted.
    Source *Document

    // Format controls how style conflicts are resolved.
    // Default (zero value): UseDestinationStyles.
    Format ImportFormatMode

    // Options provides fine-grained import control.
    // Default (zero value): all false.
    Options ImportFormatOptions
}
```

**Обратная совместимость:** `ContentData{Source: doc}` — Format=0=UseDestinationStyles, Options=zero — идентично текущему поведению.

### 2.3 Прокидывание в ResourceImporter

```go
// Файл: pkg/docx/resourceimport.go

type ResourceImporter struct {
    // ... existing fields ...

    // NEW: import configuration
    importFormatMode ImportFormatMode
    opts             ImportFormatOptions

    // NEW: styles marked for expansion to direct attributes.
    // Populated by mergeOneStyle when KeepSourceFormatting/KeepDifferentStyles
    // encounters a conflict. Key = source styleId, Value = source CT_Style.
    expandStyles map[string]*oxml.CT_Style
}

func newResourceImporter(sourceDoc, targetDoc *Document, targetPkg *parts.WmlPackage,
    mode ImportFormatMode, opts ImportFormatOptions) *ResourceImporter {
    return &ResourceImporter{
        // ... existing ...
        importFormatMode: mode,
        opts:             opts,
        expandStyles:     make(map[string]*oxml.CT_Style),
    }
}
```

### 2.4 Изменение вызова в document.go

```go
// document.go:569
ri := newResourceImporter(cd.Source, d, d.wmlPkg, cd.Format, cd.Options)
```

### 2.5 Файлы для изменения

| Файл | Что |
|---|---|
| `pkg/docx/contentdata.go` | Добавить `ImportFormatMode`, `ImportFormatOptions`, расширить `ContentData` |
| `pkg/docx/resourceimport.go` | Добавить поля `importFormatMode`, `opts`, `expandStyles`; изменить `newResourceImporter` |
| `pkg/docx/document.go` | Передать `cd.Format`, `cd.Options` в `newResourceImporter` |

---

<a id="step-3"></a>
## Step 3 — ImportFormatMode: 3 стратегии в mergeOneStyle

**Приоритет:** HIGH
**Ключевой файл:** `pkg/docx/resourceimport_styles.go:169`

### 3.1 Текущая mergeOneStyle (только UseDestinationStyles)

```go
// Текущий код (строки 169-206):
func (ri *ResourceImporter) mergeOneStyle(srcStyle *oxml.CT_Style) error {
    id := srcStyle.StyleId()
    if id == "" { return nil }
    if _, done := ri.styleMap[id]; done { return nil }

    tgtStyles, _ := ri.targetStyles()
    if tgtStyles.GetByID(id) != nil {
        ri.styleMap[id] = id     // exists in target → use target
        return nil
    }
    // missing → deep-copy from source
    clone := srcStyle.RawElement().Copy()
    ri.remapNumIdsInElement(clone)
    ri.remapStyleRefsInElement(clone)
    tgtStyles.RawElement().AddChild(clone)
    ri.styleMap[id] = id
    return nil
}
```

### 3.2 Новая mergeOneStyle (3 режима)

```go
func (ri *ResourceImporter) mergeOneStyle(srcStyle *oxml.CT_Style) error {
    id := srcStyle.StyleId()
    if id == "" {
        return nil
    }
    if _, done := ri.styleMap[id]; done {
        return nil
    }

    tgtStyles, err := ri.targetStyles()
    if err != nil {
        return fmt.Errorf("docx: accessing target styles: %w", err)
    }

    existing := tgtStyles.GetByID(id)

    // --- Style NOT in target: always copy (all 3 modes agree) ---
    if existing == nil {
        return ri.copyStyleToTarget(srcStyle, id)
    }

    // --- Style EXISTS in target: behavior depends on mode ---
    switch ri.importFormatMode {

    case UseDestinationStyles:
        // Conflict → use target definition. Current behavior.
        ri.styleMap[id] = id

    case KeepSourceFormatting:
        if ri.opts.ForceCopyStyles {
            // Copy with unique suffix: Heading1 → Heading1_0
            newId := ri.uniqueStyleId(id)
            return ri.copyStyleToTarget(srcStyle, newId)
        }
        // Default: mark for expansion to direct attributes
        ri.expandStyles[id] = srcStyle
        ri.styleMap[id] = ri.targetDefaultParaStyleId()

    case KeepDifferentStyles:
        if stylesContentEqual(srcStyle, existing) {
            // Same formatting → use target (like UseDestination)
            ri.styleMap[id] = id
        } else {
            // Different formatting → expand (like KeepSourceFormatting)
            if ri.opts.ForceCopyStyles {
                newId := ri.uniqueStyleId(id)
                return ri.copyStyleToTarget(srcStyle, newId)
            }
            ri.expandStyles[id] = srcStyle
            ri.styleMap[id] = ri.targetDefaultParaStyleId()
        }
    }
    return nil
}
```

### 3.3 Новые helper-методы

```go
// copyStyleToTarget deep-copies srcStyle into target styles.xml under targetId.
func (ri *ResourceImporter) copyStyleToTarget(srcStyle *oxml.CT_Style, targetId string) error {
    tgtStyles, err := ri.targetStyles()
    if err != nil {
        return err
    }
    clone := srcStyle.RawElement().Copy()
    if targetId != srcStyle.StyleId() {
        clone.CreateAttr("w:styleId", targetId)
        // Also rename w:name to avoid confusion in Word style gallery.
        if nameEl := findChild(clone, "w", "name"); nameEl != nil {
            if v := nameEl.SelectAttrValue("w:val", ""); v != "" {
                nameEl.CreateAttr("w:val", v+" (imported)")
            }
        }
    }
    ri.remapNumIdsInElement(clone)
    ri.remapStyleRefsInElement(clone)
    tgtStyles.RawElement().AddChild(clone)
    ri.styleMap[srcStyle.StyleId()] = targetId
    return nil
}

// targetDefaultParaStyleId returns the styleId of the target document's
// default paragraph style. Used as the replacement when expanding
// source styles to direct attributes.
func (ri *ResourceImporter) targetDefaultParaStyleId() string {
    tgtStyles, err := ri.targetStyles()
    if err != nil {
        return "Normal" // fallback
    }
    def, err := tgtStyles.DefaultFor(enum.WdStyleTypeParagraph)
    if err != nil || def == nil {
        return "Normal"
    }
    return def.StyleId()
}

// uniqueStyleId generates a unique styleId by appending _0, _1, etc.
func (ri *ResourceImporter) uniqueStyleId(base string) string {
    tgtStyles, _ := ri.targetStyles()
    for i := 0; ; i++ {
        candidate := fmt.Sprintf("%s_%d", base, i)
        if tgtStyles.GetByID(candidate) == nil {
            if _, used := ri.styleMap[candidate]; !used {
                return candidate
            }
        }
    }
}

// stylesContentEqual compares two styles by formatting content,
// ignoring w:name and w:rsid* (which don't affect appearance).
//
// Верифицировано: etree.WriteSettings.CanonicalText существует (v1.6.0).
// Element.WriteTo не существует — нужно использовать Document.WriteTo.
func stylesContentEqual(a, b *oxml.CT_Style) bool {
    ac := a.RawElement().Copy()
    bc := b.RawElement().Copy()
    stripNonFormattingChildren(ac)
    stripNonFormattingChildren(bc)

    var bufA, bufB bytes.Buffer

    docA := etree.NewDocument()
    docA.SetRoot(ac)
    docA.WriteSettings = etree.WriteSettings{CanonicalText: true}
    docA.WriteTo(&bufA)

    docB := etree.NewDocument()
    docB.SetRoot(bc)
    docB.WriteSettings = etree.WriteSettings{CanonicalText: true}
    docB.WriteTo(&bufB)

    return bytes.Equal(bufA.Bytes(), bufB.Bytes())
}

// stripNonFormattingChildren removes w:name and w:rsid* elements/attrs
// from a cloned style element before comparison.
func stripNonFormattingChildren(el *etree.Element) {
    var toRemove []*etree.Element
    for _, child := range el.ChildElements() {
        if child.Space == "w" && child.Tag == "name" {
            toRemove = append(toRemove, child)
        }
    }
    for _, rm := range toRemove {
        el.RemoveChild(rm)
    }
    // Remove rsid attributes (revision session IDs).
    filtered := el.Attr[:0]
    for _, a := range el.Attr {
        if !strings.HasPrefix(a.Key, "rsid") {
            filtered = append(filtered, a)
        }
    }
    el.Attr = filtered
}
```

### 3.4 Файлы для изменения

| Файл | Что |
|---|---|
| `pkg/docx/resourceimport_styles.go` | Переписать `mergeOneStyle`, добавить `copyStyleToTarget`, `targetDefaultParaStyleId`, `uniqueStyleId`, `stylesContentEqual`, `stripNonFormattingChildren` |

---

<a id="step-4"></a>
## Step 4 — Expand to Direct Attributes

**Приоритет:** HIGH
**Суть:** Ключевая логика KeepSourceFormatting. Когда стиль конфликтует и `ForceCopyStyles=false`, formatting source стиля раскрывается в direct attributes на каждом элементе, ссылающемся на этот стиль.

### 4.1 Логика

Для каждого source стиля, помеченного для expansion (в `ri.expandStyles`):

1. **Resolve formatting chain:** style → basedOn → basedOn → ... до конца цепочки
2. **Merge properties:** Собрать итоговые pPr и rPr (дочерние перезаписывают родительские)
3. **Apply to paragraphs:** Для каждого `<w:p>`, чей pStyle == этот styleId, слить resolved pPr + rPr
4. **Apply to runs:** Для каждого `<w:r>`, чей rStyle == этот styleId, слить resolved rPr
5. **Style reference** уже изменена на Normal/default через `remapAll` + `styleMap`

### 4.2 Pipeline restructuring

Текущий pipeline в `prepareContentElements` (`contentdata.go:265-337`):

```
Step 2:  deep copy
Step 2b: sanitize
Step 2c: materialize implicit styles
Step 2d: remapAll (styles, numId, footnotes)    ← style refs changed here
Step 3:  collect rIds
Step 4:  import relationships
Step 5:  remap rIds
```

Для expand-to-direct нужно:

```
Step 2:  deep copy
Step 2b: sanitize
Step 2c: materialize implicit styles
Step 2d: expandDirectFormatting(elements, ri)   ← NEW: BEFORE remapAll
Step 2e: remapAll (styles, numId, footnotes)    ← style refs changed here
Step 3:  collect rIds
Step 4:  import relationships
Step 5:  remap rIds
```

Порядок expandDirectFormatting → remapAll правильный потому что:
- `expandDirectFormatting` читает **оригинальные** styleId (ещё не ремапленные)
- Находит elements с pStyle/rStyle, которые есть в `ri.expandStyles`
- Применяет resolved properties к элементу
- Затем `remapAll` меняет pStyle на target default

### 4.3 Реализация

```go
// Файл: pkg/docx/resourceimport_styles.go (новые функции)

// expandDirectFormatting walks elements and for each paragraph/run
// whose style is in expandStyles, merges resolved source formatting
// into direct attributes.
//
// Handles BOTH paragraph styles (pStyle → pPr + rPr) and character
// styles (rStyle → rPr). This mirrors Aspose.Words which expands
// both style types to direct attributes.
//
// Pipeline position:
//   sanitize → materialize → expandDirectFormatting → remapAll → import rIds
func (ri *ResourceImporter) expandDirectFormatting(elements []*etree.Element) {
    if len(ri.expandStyles) == 0 {
        return
    }

    for _, root := range elements {
        stack := []*etree.Element{root}
        for len(stack) > 0 {
            el := stack[len(stack)-1]
            stack = stack[:len(stack)-1]

            if el.Space == "w" {
                switch el.Tag {
                case "p":
                    ri.expandParagraphStyle(el)
                    // Also process runs within this paragraph.
                    ri.expandRunStylesInParagraph(el)
                case "r":
                    // Runs at top level (outside paragraph) — rare but possible.
                    ri.expandRunStyle(el)
                }
            }
            stack = append(stack, el.ChildElements()...)
        }
    }
}

// expandParagraphStyle checks if the paragraph's pStyle is in expandStyles.
// If so, merges the resolved source pPr and rPr (default run formatting)
// into the paragraph's existing pPr.
func (ri *ResourceImporter) expandParagraphStyle(pEl *etree.Element) {
    pPr := findChild(pEl, "w", "pPr")
    if pPr == nil {
        return
    }
    pStyle := findChild(pPr, "w", "pStyle")
    if pStyle == nil {
        return
    }
    styleId := pStyle.SelectAttrValue("w:val", "")
    srcStyle, needsExpand := ri.expandStyles[styleId]
    if !needsExpand {
        return
    }

    // Resolve full formatting from source style chain.
    resolvedPPr, resolvedRPr := ri.resolveStyleChain(srcStyle)

    // Merge resolved pPr into existing pPr.
    // Existing direct attrs take precedence (user formatting > style).
    if resolvedPPr != nil {
        mergePropertiesDeep(pPr, resolvedPPr)
    }

    // Merge resolved rPr into the paragraph-level rPr (default run props).
    if resolvedRPr != nil {
        existingRPr := findChild(pPr, "w", "rPr")
        if existingRPr == nil {
            existingRPr = etree.NewElement("w:rPr")
            pPr.AddChild(existingRPr)
        }
        mergePropertiesDeep(existingRPr, resolvedRPr)
    }
}

// expandRunStylesInParagraph processes all <w:r> children of a paragraph,
// expanding character styles (rStyle) that are in expandStyles.
func (ri *ResourceImporter) expandRunStylesInParagraph(pEl *etree.Element) {
    for _, child := range pEl.ChildElements() {
        if child.Space == "w" && child.Tag == "r" {
            ri.expandRunStyle(child)
        }
    }
}

// expandRunStyle checks if the run's rStyle is in expandStyles.
// If so, merges the resolved source rPr into the run's existing rPr.
func (ri *ResourceImporter) expandRunStyle(rEl *etree.Element) {
    rPr := findChild(rEl, "w", "rPr")
    if rPr == nil {
        return
    }
    rStyle := findChild(rPr, "w", "rStyle")
    if rStyle == nil {
        return
    }
    styleId := rStyle.SelectAttrValue("w:val", "")
    srcStyle, needsExpand := ri.expandStyles[styleId]
    if !needsExpand {
        return
    }

    // For character styles, only rPr is relevant.
    _, resolvedRPr := ri.resolveStyleChain(srcStyle)
    if resolvedRPr != nil {
        mergePropertiesDeep(rPr, resolvedRPr)
    }
}

// resolveStyleChain walks the basedOn chain in the source styles part
// and merges pPr/rPr properties (child overrides parent).
//
// Returns copies — callers may safely modify the returned elements.
func (ri *ResourceImporter) resolveStyleChain(style *oxml.CT_Style) (pPr, rPr *etree.Element) {
    srcStyles, err := ri.sourceStyles()
    if err != nil {
        return nil, nil
    }
    // Build chain: [style, parent, grandparent, ...]
    var chain []*oxml.CT_Style
    visited := map[string]bool{}
    current := style
    for current != nil {
        id := current.StyleId()
        if visited[id] {
            break // cycle protection
        }
        visited[id] = true
        chain = append(chain, current)
        if basedOn, _ := current.BasedOnVal(); basedOn != "" {
            current = srcStyles.GetByID(basedOn)
        } else {
            current = nil
        }
    }
    // Merge from base to derived (so derived overrides base).
    for i := len(chain) - 1; i >= 0; i-- {
        raw := chain[i].RawElement()
        if p := findChild(raw, "w", "pPr"); p != nil {
            if pPr == nil {
                pPr = p.Copy()
            } else {
                mergePropertiesDeep(pPr, p)
            }
        }
        if r := findChild(raw, "w", "rPr"); r != nil {
            if rPr == nil {
                rPr = r.Copy()
            } else {
                mergePropertiesDeep(rPr, r)
            }
        }
    }
    return
}

// mergePropertiesDeep merges children of src into dst with attribute-level
// granularity. For each child in src:
//   - If dst has no child with the same space:tag → copy entire child from src
//   - If dst has child with the same space:tag → merge attributes from src
//     child into dst child (dst attributes take precedence)
//
// This produces correct results for complex properties like <w:rFonts>
// where src might have w:ascii and dst might have w:hAnsi — the result
// should contain both attributes.
func mergePropertiesDeep(dst, src *etree.Element) {
    for _, srcChild := range src.ChildElements() {
        dstChild := findChild(dst, srcChild.Space, srcChild.Tag)
        if dstChild == nil {
            // Not present in dst — copy entire element.
            dst.AddChild(srcChild.Copy())
        } else {
            // Both have this property — merge attributes.
            // dst attributes take precedence (direct formatting > style).
            mergeAttrs(dstChild, srcChild)
        }
    }
}

// mergeAttrs copies attributes from src to dst that don't already exist in dst.
// dst attributes take precedence — existing attributes are never overwritten.
func mergeAttrs(dst, src *etree.Element) {
    dstAttrKeys := make(map[string]bool, len(dst.Attr))
    for _, a := range dst.Attr {
        key := a.Space + ":" + a.Key
        dstAttrKeys[key] = true
    }
    for _, a := range src.Attr {
        key := a.Space + ":" + a.Key
        if !dstAttrKeys[key] {
            dst.Attr = append(dst.Attr, a)
        }
    }
}

// findChild finds first child element with given space:tag.
func findChild(el *etree.Element, space, tag string) *etree.Element {
    for _, child := range el.ChildElements() {
        if child.Space == space && child.Tag == tag {
            return child
        }
    }
    return nil
}
```

### 4.4 Файлы для изменения

| Файл | Что |
|---|---|
| `pkg/docx/resourceimport_styles.go` | Добавить `expandDirectFormatting`, `expandParagraphStyle`, `expandRunStylesInParagraph`, `expandRunStyle`, `resolveStyleChain`, `mergePropertiesDeep`, `mergeAttrs`, `findChild` |
| `pkg/docx/contentdata.go` | Вставить вызов `ri.expandDirectFormatting(elements)` перед `ri.remapAll(elements)` в `prepareContentElements` (после строки 303, перед строкой 309) |

---

<a id="step-5"></a>
## Step 5 — ForceCopyStyles (переименование стилей)

**Приоритет:** MEDIUM
**Зависимость:** Step 3 (уже реализует ветку `ForceCopyStyles` в `mergeOneStyle`)

Этот step уже полностью покрыт Step 3:
- `mergeOneStyle` при `ForceCopyStyles=true` вызывает `uniqueStyleId(id)` → `copyStyleToTarget(srcStyle, newId)`
- `styleMap[originalId] = newId` → `remapAll` перепишет все ссылки в body elements
- `remapStyleRefsInElement` перепишет basedOn/next/link в скопированном стиле

**Дополнительная работа:** При `ForceCopyStyles` скопированный стиль должен быть помечен как скрытый, чтобы не мусорить в Style Gallery Word:

```go
func (ri *ResourceImporter) copyStyleToTarget(srcStyle *oxml.CT_Style, targetId string) error {
    // ... existing code (see Step 3.3) ...

    // If this is a renamed copy (ForceCopyStyles), mark as semi-hidden.
    if targetId != srcStyle.StyleId() {
        // semiHidden — скрывает стиль из Style Gallery.
        sh := etree.NewElement("w:semiHidden")
        clone.AddChild(sh)
        // unhideWhenUsed — стиль появляется в gallery при использовании.
        uwu := etree.NewElement("w:unhideWhenUsed")
        clone.AddChild(uwu)
    }

    // ... rest of existing code ...
}
```

### 5.1 Файлы для изменения

Уже покрыт в Step 3. Единственное дополнение — `semiHidden`/`unhideWhenUsed` в `copyStyleToTarget`.

---

<a id="step-6"></a>
## Step 6 — Deep Generic Part Import

**Приоритет:** HIGH
**Текущая проблема:** `contentdata.go:235-236`:
> *"Sub-relationships of the source part are NOT imported (shallow copy)"*

Это означает charts, SmartArt, VML drawings теряют данные (embedded Excel, style/color XML).

### 6.1 Типичная структура chart в OOXML

```
/word/charts/chart1.xml           ← main chart (XML, has sub-rels)
  └→ rel to /word/charts/style1.xml     ← chart style
  └→ rel to /word/charts/colors1.xml    ← chart colors
  └→ rel to /word/embeddings/wb1.xlsx   ← data (binary, no sub-rels)
```

### 6.2 Проверка: какие методы уже есть

| Метод | Существует | Файл:строка |
|---|---|---|
| `BasePart.Blob()` | Да | `opc/part.go:50` |
| `BasePart.SetBlob()` | Да | `opc/part.go:63` |
| `BasePart.Rels()` | Да | `opc/part.go:51` |
| `BasePart.SetRels()` | Да | `opc/part.go:52` |
| `OpcPackage.AddPart()` | Да | `opc/package.go:239` |
| `Relationships.All()` | Да | `opc/rel.go:133` |
| `Relationships.Add()` | Да | — |
| `Relationships.GetOrAddExtRel()` | Да | — |
| `Part` interface: `Rels()`, `Blob()` | Да | `opc/part.go:14-20` |

Все нужные методы существуют. `RemovePart` не нужен.

### 6.3 Архитектура deep import

**Дизайн:** Единая рекурсивная функция `importPartDeep` обрабатывает
любой internal part на любой глубине. `importGenericPart` делегирует
в `importPartDeep` для первого уровня. `importPartDeep` рекурсивно
вызывает себя для sub-relationships.

Ключевое отличие от shallow версии: после копирования blob'а мы
обходим sub-relationships source part'а и рекурсивно импортируем
каждый sub-part. Затем, если part — XML, перезаписываем rId
ссылки в blob'е.

**Защита от бесконечной рекурсии:** `maxDepth = 10` (в реальных OOXML
документах глубина не превышает 3–4).

```go
// Файл: pkg/docx/contentdata.go (замена функции importGenericPart)

// importGenericPart copies a non-image internal part from source to target.
// Delegates to importPartDeep for recursive sub-relationship import.
func importGenericPart(
    srcRel *opc.Relationship,
    targetPart *parts.StoryPart,
    targetPkg *parts.WmlPackage,
    importedParts map[opc.PackURI]opc.Part,
) (string, error) {
    srcPart := srcRel.TargetPart
    srcPN := srcPart.PartName()

    // Dedup: same source part already imported.
    if existing, ok := importedParts[srcPN]; ok {
        rel := targetPart.Rels().GetOrAdd(srcRel.RelType, existing)
        return rel.RID, nil
    }

    // Deep import the part and its sub-relationships.
    newPart, err := importPartDeep(srcPart, targetPkg, importedParts, 0)
    if err != nil {
        return "", err
    }

    // Create relationship from caller's part to the new part.
    targetRef := newPart.PartName().RelativeRef(targetPart.Rels().BaseURI())
    rel := targetPart.Rels().Add(srcRel.RelType, targetRef, newPart, false)
    return rel.RID, nil
}

// importPartDeep recursively copies an internal part and all its
// sub-relationships into the target package.
//
// Pipeline per part:
//  1. Copy blob into target with fresh partname
//  2. For each sub-relationship of source part:
//     - External → GetOrAddExtRel on new part
//     - Image → SHA-256 dedup (GetOrAddImagePart) + relate on new part
//     - Internal → RECURSE (importPartDeep with depth+1) + relate on new part
//  3. If part is XML and has remapped sub-rIds → rewrite rIds in blob
//
// importedParts is shared across the entire ReplaceWithContent call
// (passed from ResourceImporter) so parts are never duplicated.
func importPartDeep(
    srcPart opc.Part,
    targetPkg *parts.WmlPackage,
    importedParts map[opc.PackURI]opc.Part,
    depth int,
) (opc.Part, error) {
    const maxDepth = 10
    if depth > maxDepth {
        return nil, fmt.Errorf("docx: part import depth exceeds %d", maxDepth)
    }

    srcPN := srcPart.PartName()

    // Dedup check.
    if existing, ok := importedParts[srcPN]; ok {
        return existing, nil
    }

    // 1. Copy blob.
    blob, err := srcPart.Blob()
    if err != nil {
        return nil, fmt.Errorf("reading part blob %s: %w", srcPN, err)
    }

    template := partNameTemplate(srcPN)
    newPN := targetPkg.OpcPackage.NextPartname(template)
    newPart := opc.NewBasePart(newPN, srcPart.ContentType(), blob, targetPkg.OpcPackage)
    newPart.SetRels(opc.NewRelationships(newPN.BaseURI()))
    targetPkg.OpcPackage.AddPart(newPart)
    importedParts[srcPN] = newPart

    // 2. Import sub-relationships.
    srcSubRels := srcPart.Rels()
    if srcSubRels == nil {
        return newPart, nil
    }
    allSubs := srcSubRels.All()
    if len(allSubs) == 0 {
        return newPart, nil
    }

    subRidMap := make(map[string]string, len(allSubs))

    for _, subRel := range allSubs {
        var newRId string

        if subRel.IsExternal {
            // External sub-relationship (e.g. hyperlink from chart).
            newRId = newPart.Rels().GetOrAddExtRel(subRel.RelType, subRel.TargetRef)
        } else if subRel.TargetPart == nil {
            continue
        } else if subRel.RelType == opc.RTImage {
            // Image sub-rel — SHA-256 dedup via WmlPackage.
            rId, err := importImageSubRel(subRel, newPart, targetPkg)
            if err != nil {
                continue // graceful degradation
            }
            newRId = rId
        } else {
            // Generic sub-part → RECURSE.
            subPart, err := importPartDeep(
                subRel.TargetPart, targetPkg, importedParts, depth+1,
            )
            if err != nil {
                continue // graceful degradation
            }
            subRef := subPart.PartName().RelativeRef(newPart.Rels().BaseURI())
            rel := newPart.Rels().Add(subRel.RelType, subRef, subPart, false)
            newRId = rel.RID
        }

        if newRId != "" {
            subRidMap[subRel.RID] = newRId
        }
    }

    // 3. Rewrite rIds in blob if it's XML and has remappings.
    if len(subRidMap) > 0 && isXmlContentType(newPart.ContentType()) {
        if rewritten, err := rewriteRIdsInBlob(blob, subRidMap); err == nil {
            newPart.SetBlob(rewritten)
        }
        // Parse failure → keep original blob (graceful degradation).
    }

    return newPart, nil
}

// importImageSubRel imports an image sub-relationship onto newPart.
// Extracted to keep importPartDeep readable.
func importImageSubRel(
    subRel *opc.Relationship,
    newPart *opc.BasePart,
    targetPkg *parts.WmlPackage,
) (string, error) {
    srcIP, ok := subRel.TargetPart.(*parts.ImagePart)
    if !ok {
        return "", fmt.Errorf("image sub-rel target is %T, want *ImagePart", subRel.TargetPart)
    }
    imgBlob, err := srcIP.Blob()
    if err != nil {
        return "", err
    }
    cloneIP := parts.NewImagePart(srcIP.PartName(), srcIP.ContentType(), imgBlob, nil)
    cloneIP.SetFilename(srcIP.Filename())
    // Copy image metadata if available.
    if w, err := srcIP.PxWidth(); err == nil {
        if h, errH := srcIP.PxHeight(); errH == nil {
            if hDpi, errD := srcIP.HorzDpi(); errD == nil {
                if vDpi, errV := srcIP.VertDpi(); errV == nil {
                    cloneIP.SetImageMeta(w, h, hDpi, vDpi)
                }
            }
        }
    }
    dedupIP, err := targetPkg.GetOrAddImagePart(cloneIP)
    if err != nil {
        return "", err
    }
    rel := newPart.Rels().GetOrAdd(opc.RTImage, dedupIP)
    return rel.RID, nil
}

// isXmlContentType checks if the content type indicates XML data.
func isXmlContentType(ct string) bool {
    return strings.HasSuffix(ct, "+xml") || strings.HasSuffix(ct, "/xml")
}

// rewriteRIdsInBlob parses XML, remaps relationship attributes, re-serializes.
// Uses the same remapRIds function used for body elements — it does
// recursive DFS internally so all nested rId refs are covered.
func rewriteRIdsInBlob(blob []byte, ridMap map[string]string) ([]byte, error) {
    doc := etree.NewDocument()
    if err := doc.ReadFromBytes(blob); err != nil {
        return nil, err
    }
    root := doc.Root()
    if root == nil {
        return nil, fmt.Errorf("empty XML document")
    }
    // remapRIds does iterative DFS — handles all descendants.
    remapRIds([]*etree.Element{root}, ridMap)
    return doc.WriteToBytes()
}
```

### 6.4 Файлы для изменения

| Файл | Что |
|---|---|
| `pkg/docx/contentdata.go` | Заменить `importGenericPart` на `importGenericPart` + `importPartDeep` + `importImageSubRel` + `isXmlContentType` + `rewriteRIdsInBlob` |

---

<a id="step-7"></a>
## Step 7 — ImportFormatOptions: нумерация

**Приоритет:** MEDIUM
**Ключевой файл:** `pkg/docx/resourceimport_num.go`

### 7.1 Текущее поведение

Текущий код **всегда** создаёт новый `abstractNum` с новым nsid — эквивалент `KeepSourceNumbering = true`. Это правильный default для enterprise (безопасный, не ломает существующую нумерацию target).

### 7.2 Необходимые методы в oxml (НЕ существуют — нужно добавить)

**Верифицировано:** `NumList()` существует в `zz_gen_numbering.go:20`. Но `AllAbstractNums()` и `AllNums()` — НЕ существуют. Нужно добавить в `numbering_custom.go`.

```go
// Файл: pkg/docx/oxml/numbering_custom.go (ДОБАВИТЬ)

// AllAbstractNums returns all <w:abstractNum> child elements as raw etree elements.
// Used by KeepSourceNumbering=false to find matching list definitions.
func (n *CT_Numbering) AllAbstractNums() []*etree.Element {
    var result []*etree.Element
    for _, child := range n.e.ChildElements() {
        if child.Space == "w" && child.Tag == "abstractNum" {
            result = append(result, child)
        }
    }
    return result
}

// AbstractNumId extracts the w:abstractNumId attribute value from a raw
// <w:abstractNum> element. Returns -1 if not present or invalid.
func AbstractNumIdOf(el *etree.Element) int {
    v := el.SelectAttrValue("w:abstractNumId", "")
    if v == "" {
        return -1
    }
    id, err := strconv.Atoi(v)
    if err != nil {
        return -1
    }
    return id
}
```

### 7.3 Добавить KeepSourceNumbering=false

При `KeepSourceNumbering = false` (Aspose default): если в target есть abstractNum с аналогичным list style, source list вливается в target list:

```go
// Файл: pkg/docx/resourceimport_num.go (модифицировать importNumbering)

// В importNumbering(), после нахождения srcNum и srcAbsId, ПЕРЕД созданием нового abstractNum:
if !ri.opts.KeepSourceNumbering {
    // Try to find matching target num by comparing
    // list style (numFmt + pStyle pattern in first level).
    tgtNumId := ri.findMatchingTargetNum(srcAbsId, srcNumbering, tgtNumbering)
    if tgtNumId > 0 {
        ri.numIdMap[srcNumId] = tgtNumId
        continue // skip creating new abstractNum
    }
}
// ... existing: create new abstractNum ...
```

```go
// findMatchingTargetNum finds a target num whose abstractNum has
// a compatible list style, based on the first level's numFmt.
//
// Matching heuristic: compare numFmt of the first <w:lvl w:ilvl="0">
// in source vs target abstractNums. If both use the same numFmt
// (bullet, decimal, etc.), they are considered a match.
func (ri *ResourceImporter) findMatchingTargetNum(
    srcAbsId int,
    srcNumbering, tgtNumbering *oxml.CT_Numbering,
) int {
    srcAbsNum := srcNumbering.FindAbstractNum(srcAbsId)
    if srcAbsNum == nil {
        return 0
    }
    srcFmt := firstLevelNumFmt(srcAbsNum)
    if srcFmt == "" {
        return 0
    }

    // Scan target abstractNums for matching first-level numFmt.
    for _, tgtAbsNum := range tgtNumbering.AllAbstractNums() {
        if firstLevelNumFmt(tgtAbsNum) == srcFmt {
            tgtAbsId := oxml.AbstractNumIdOf(tgtAbsNum)
            if tgtAbsId < 0 {
                continue
            }
            // Find num referencing this abstractNum.
            for _, tgtNum := range tgtNumbering.NumList() {
                if absId, err := abstractNumIdOf(tgtNum); err == nil && absId == tgtAbsId {
                    numId, err := tgtNum.NumId()
                    if err == nil {
                        return numId
                    }
                }
            }
        }
    }
    return 0
}

// firstLevelNumFmt returns the w:numFmt/@w:val of the first level
// (w:ilvl="0") in an abstractNum element. Returns "" if not found.
func firstLevelNumFmt(absNum *etree.Element) string {
    for _, child := range absNum.ChildElements() {
        if child.Space == "w" && child.Tag == "lvl" {
            if child.SelectAttrValue("w:ilvl", "") == "0" {
                for _, lvlChild := range child.ChildElements() {
                    if lvlChild.Space == "w" && lvlChild.Tag == "numFmt" {
                        return lvlChild.SelectAttrValue("w:val", "")
                    }
                }
            }
        }
    }
    return ""
}
```

### 7.4 Файлы для изменения

| Файл | Что |
|---|---|
| `pkg/docx/oxml/numbering_custom.go` | Добавить `AllAbstractNums()`, `AbstractNumIdOf()` |
| `pkg/docx/resourceimport_num.go` | Добавить `findMatchingTargetNum`, `firstLevelNumFmt`; добавить `if !ri.opts.KeepSourceNumbering` ветку в `importNumbering` |

---

<a id="step-8"></a>
## Step 8 — Тесты (60+)

**Приоритет:** CRITICAL

### 8.1 Тесты для backward compatibility (существующие 25 тестов должны остаться зелёными)

Все существующие тесты в `replacecontent_test.go` используют `ContentData{Source: doc}` — zero value Format и Options. Они НЕ ДОЛЖНЫ быть изменены.

### 8.2 Новые тесты: ImportFormatMode

```
TestRWC_UseDestination_ExistingStyleKept         — стиль target не перезаписывается
TestRWC_UseDestination_MissingStyleCopied         — отсутствующий стиль копируется
TestRWC_KeepSource_NoConflict_Copied              — стиль не в target → копируется
TestRWC_KeepSource_Conflict_ExpandedToDirect      — конфликт → pPr/rPr прямые, стиль=Normal
TestRWC_KeepSource_Conflict_DirectPreservesExisting — existing direct attrs не перезаписываются
TestRWC_KeepSource_BasedOnChain_Resolved          — basedOn цепочка resolved корректно
TestRWC_KeepSource_ForceRename_SuffixGenerated    — ForceCopyStyles → _0 суффикс
TestRWC_KeepSource_ForceRename_ChainRenamed       — basedOn цепочка → оба переименованы
TestRWC_KeepSource_ForceRename_SemiHidden         — renamed стиль помечен semiHidden
TestRWC_KeepDifferent_IdenticalUseTarget          — одинаковые стили → target
TestRWC_KeepDifferent_DifferentExpanded           — разные стили → expanded
TestRWC_BackwardCompat_ZeroValue                  — ContentData{Source:doc} = UseDestination
TestRWC_BackwardCompat_AllExistingTestsGreen      — meta-test
```

### 8.3 Новые тесты: Expand to Direct Attributes

```
TestRWC_Expand_ParagraphStyle_PPrMerged           — pPr из стиля слит в параграф
TestRWC_Expand_ParagraphStyle_RPrMerged           — rPr из paragraph style слит
TestRWC_Expand_RunStyle_Expanded                  — rStyle на run раскрыт в direct rPr
TestRWC_Expand_MixedStyles_BothExpanded           — paragraph + character style оба expanded
TestRWC_Expand_ExistingDirectWins                 — direct attrs не перезаписываются style attrs
TestRWC_Expand_DeepMerge_RFonts                   — <w:rFonts> attrs мержатся на уровне атрибутов
TestRWC_Expand_BasedOnChain_3Levels               — цепочка basedOn из 3 уровней resolved
TestRWC_Expand_CyclicBasedOn_NoPanic              — cycle в basedOn не вызывает infinite loop
```

### 8.4 Новые тесты: Deep Generic Part Import

```
TestRWC_DeepImport_SubRelsImported                — sub-rels скопированы
TestRWC_DeepImport_SubRidsRemapped                — rIds в XML blob ремаплены
TestRWC_DeepImport_BinarySubPart                  — non-XML blob → copy without rewrite
TestRWC_DeepImport_DepthLimit                     — depth > 10 → ошибка
TestRWC_DeepImport_Dedup                          — один source part → один target part
TestRWC_DeepImport_RoundTrip                      — save → reopen → sub-parts intact
TestRWC_DeepImport_NestedSubRels                  — sub-sub-relationships (3 levels)
```

### 8.5 Новые тесты: Error Recovery

```
TestRWC_Snapshot_BodyRestoredOnError              — ошибка → body восстановлен
TestRWC_Snapshot_BodyUnchangedOnSuccess            — без ошибки → snapshot не применяется
```

### 8.6 Новые тесты: Edge Cases

```
TestRWC_SDT_Preserved                             — SDT elements intact после Copy
TestRWC_FieldCode_Preserved                       — field codes intact
TestRWC_OLE_PreviewAndBlob                        — OLE preview image + blob imported
TestRWC_100Placeholders_SameTag                   — scaling test
TestRWC_DeeplyNestedTables_5Levels                — recursion depth
TestRWC_UnicodeTag_Cyrillic                       — кириллица в теге
TestRWC_KeepSourceNumbering_SeparateList          — source list отдельный
TestRWC_KeepSourceNumbering_False_Merged          — source list вливается
```

### 8.7 Benchmarks

```go
// Файл: pkg/docx/replacecontent_bench_test.go (НОВЫЙ)

func BenchmarkRWC_Simple(b *testing.B)                     // baseline
func BenchmarkRWC_WithImages(b *testing.B)                 // image dedup
func BenchmarkRWC_100Replacements(b *testing.B)            // scaling
func BenchmarkRWC_KeepSourceFormatting(b *testing.B)       // expand overhead
func BenchmarkRWC_Snapshot(b *testing.B)                   // snapshot overhead
```

### 8.8 Fuzz

```go
// Файл: pkg/docx/replacecontent_fuzz_test.go (НОВЫЙ)

func FuzzRWC_TagDiscovery(f *testing.F) {
    f.Add("[[test]]")
    f.Add("[<CONTENT>]")
    f.Add("{{tag}}")
    f.Fuzz(func(t *testing.T, tag string) {
        if len(tag) == 0 || len(tag) > 100 { return }
        target := mustNewDoc(t)
        target.AddParagraph("prefix " + tag + " suffix")
        source := mustNewDoc(t)
        source.AddParagraph("inserted")
        target.ReplaceWithContent(tag, ContentData{Source: source})
    })
}
```

### 8.9 Файлы

| Файл | Тип |
|---|---|
| `pkg/docx/replacecontent_test.go` | Расширить (~50 новых тестов) |
| `pkg/docx/replacecontent_bench_test.go` | НОВЫЙ |
| `pkg/docx/replacecontent_fuzz_test.go` | НОВЫЙ |

---

<a id="что-не-нужно-делать"></a>
## Что НЕ нужно делать

Верифицировано против реального поведения Aspose.Words:

| Фича | Почему НЕ нужна |
|---|---|
| Theme merge / Theme color resolution | Aspose НЕ мержит theme. Один theme на документ. Expand-to-direct уже решает проблему конфликтов |
| Font table merge | Aspose НЕ мержит. Resolution — задача рендера, не import |
| Document.Validate() | У Aspose НЕТ такого API |
| SDT data binding rebind | Aspose сохраняет as-is. Наш Copy() уже корректен |
| Field code remap / dirty flag | Aspose сохраняет as-is. Word пересчитывает |
| CustomXml parts import | Aspose требует ручного клонирования |
| Progress callbacks | У Aspose нет для import |
| Glossary document merge | Крайне редкий use case |
| OpcPackage.RemovePart() | Метод не существует. Orphan parts безвредны |
| Полный Document.Clone() для rollback | Слишком дорого. Body snapshot достаточен |

---

<a id="файловая-карта-изменений"></a>
## Файловая карта изменений

### Изменяемые файлы (6)

| Файл | Steps | Объём изменений |
|---|---|---|
| `pkg/docx/contentdata.go` | 2, 4, 6 | +120 строк (типы API, deep import, rId blob rewrite, expandDirectFormatting call) |
| `pkg/docx/resourceimport.go` | 2 | +15 строк (новые поля в struct, параметры в constructor) |
| `pkg/docx/resourceimport_styles.go` | 3, 4, 5 | +250 строк (3 ImportFormatMode, expand-to-direct с rStyle support, style comparison, deep merge) |
| `pkg/docx/resourceimport_num.go` | 7 | +60 строк (KeepSourceNumbering=false, firstLevelNumFmt) |
| `pkg/docx/document.go` | 1, 2 | +40 строк (snapshot, прокидывание Format/Options) |
| `pkg/docx/oxml/numbering_custom.go` | 7 | +20 строк (AllAbstractNums, AbstractNumIdOf) |

### Новые файлы (2)

| Файл | Step | Описание |
|---|---|---|
| `pkg/docx/replacecontent_bench_test.go` | 8 | Benchmarks |
| `pkg/docx/replacecontent_fuzz_test.go` | 8 | Fuzz tests |

### Расширяемые тестовые файлы (1)

| Файл | Step | Объём |
|---|---|---|
| `pkg/docx/replacecontent_test.go` | 8 | +50 новых тестов |

### НЕ изменяемые файлы

Все остальные файлы библиотеки остаются нетронутыми:
- `blkcntnr.go` — без изменений
- `resourceimport_notes.go` — без изменений
- Все файлы в `oxml/` (кроме `numbering_custom.go`), `opc/`, `parts/`, `image/`, `enum/` — без изменений
- Все остальные test-файлы — без изменений

---

<a id="критерии-приёмки"></a>
## Критерии приёмки

| # | Критерий | Как проверить |
|---|---|---|
| 1 | `ContentData{Source: doc}` (zero value) = текущее поведение | Все 25 существующих тестов зелёные |
| 2 | UseDestinationStyles: existing style kept, missing copied | 2 новых теста |
| 3 | KeepSourceFormatting: conflict → direct attrs + Normal | Тест: paragraph имеет pPr children из source стиля |
| 4 | KeepSourceFormatting: rStyle на runs тоже expanded | Тест: run имеет rPr children из source character стиля |
| 5 | KeepSourceFormatting + ForceCopyStyles: стиль с `_0` | Тест: target styles содержит `Heading1_0` |
| 6 | KeepDifferentStyles: identical → target, different → expand | 2 новых теста |
| 7 | Deep generic part import: sub-rels imported recursively | Round-trip тест с mock chart structure (3 levels) |
| 8 | Deep import: binary sub-parts copied without rewrite | Тест с non-XML blob |
| 9 | Body snapshot: ошибка → body restored | Тест с forced error |
| 10 | Expand deep merge: `<w:rFonts>` attrs мержатся корректно | Тест: src rFonts ascii + dst rFonts hAnsi = оба attrs |
| 11 | 75+ тестов, все зелёные | `go test ./pkg/docx/...` |
| 12 | Benchmarks: < 1.5x overhead vs baseline | `go test -bench BenchmarkRWC` |
| 13 | Fuzz: 0 паник за 1 минуту | `go test -fuzz FuzzRWC -fuzztime 1m` |

---

## Порядок реализации

```
Step 1 (body snapshot)     ← 1 файл, ~40 строк, независим
    ↓
Step 2 (ContentData API)   ← 3 файла, ~30 строк, foundation
    ↓
Step 3 (3 ImportFormatMode in mergeOneStyle) ← 1 файл, ~100 строк
    ↓
Step 4 (expandDirectFormatting + rStyle)  ← 2 файла, ~150 строк, зависит от Step 3
    ↓
Step 5 (ForceCopyStyles semiHidden) ← уже в Step 3, +5 строк
    ↓
Step 6 (deep generic parts) ← 1 файл, ~120 строк, независим от Steps 3-5
    ↓
Step 7 (KeepSourceNumbering) ← 2 файла, ~80 строк, зависит от Step 2
    ↓
Step 8 (тесты)              ← 3 файла, ~900 строк, последний
```

**Итого новый код:** ~1400 строк (из них ~900 — тесты).
**Существующий код modified:** ~60 строк изменений (интеграционные точки).

---

## Changelog v3 → v4 (исправленные проблемы)

| # | Проблема | Исправление |
|---|---|---|
| 1 | `stylesContentEqual` — `Element.WriteTo` не существует в etree | Используем `etree.NewDocument()` + `SetRoot()` + `Document.WriteTo()` |
| 2 | Deep import не рекурсивен — sub-sub-relationships не обрабатывались | Выделен `importPartDeep` с настоящей рекурсией (depth+1) |
| 3 | `expandDirectFormatting` не обрабатывал rStyle на runs | Добавлены `expandRunStylesInParagraph` и `expandRunStyle` |
| 4 | `filterAttrs` не была определена | Заменена на inline slice filter в `stripNonFormattingChildren` |
| 5 | `abstractNumListStyle` не была определена | Заменена на `firstLevelNumFmt` с полной реализацией |
| 6 | `AllAbstractNums()` / `AllNums()` не существовали в oxml | Добавлен `AllAbstractNums()` + `AbstractNumIdOf()` в `numbering_custom.go`; `NumList()` уже существует |
| 7 | `mergeProperties` — shallow merge на уровне элементов | Заменена на `mergePropertiesDeep` + `mergeAttrs` с attribute-level granularity |
| 8 | `restoreBody` — не верифицирован API CT_Document | Верифицирован: `CT_Document` embeds `oxml.Element` → `RawElement()` возвращает `*etree.Element` |
| 9 | Файловая карта не включала `oxml/numbering_custom.go` | Добавлен в список изменяемых файлов |
| 10 | Указано "20+ тестов" — на самом деле 25 | Исправлено на 25 |
| 11 | `KeepDifferentStyles` не обрабатывал `ForceCopyStyles` | Добавлена ветка `ForceCopyStyles` в case `KeepDifferentStyles` |
