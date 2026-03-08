# ReplaceWithContent — Enterprise-Level Plan

**Статус:** Проектирование (v2 — исправлен после ревью)
**Цель:** Привести `ReplaceWithContent` к уровню Aspose.Words
**Текущая версия метода:** Функциональна, покрывает ~80% сценариев
**Целевая версия:** Enterprise-grade, senior-level reliability

---

## 0. Аудит текущего состояния

### Что уже реализовано (и реализовано хорошо)

| Возможность | Файл | Оценка |
|---|---|---|
| Импорт стилей с транзитивным замыканием (basedOn/next/link BFS) | `resourceimport_styles.go` | 9/10 |
| UseDestinationStyles стратегия | `resourceimport_styles.go:169-206` | 9/10 |
| Материализация неявных стилей при несовпадении default paragraph style | `resourceimport_styles.go:352-408` | 10/10 |
| Импорт нумерации с пересчётом ID и регенерацией nsid | `resourceimport_num.go` | 9/10 |
| Импорт сносок/концевых сносок с полным pipeline | `resourceimport_notes.go` | 9/10 |
| Дедупликация изображений SHA-256 | `contentdata.go:162-192` | 9/10 |
| Импорт внешних гиперссылок | `contentdata.go:150-154` | 9/10 |
| Санитизация аннотационных маркеров (17 типов) | `contentdata.go:359-441` | 9/10 |
| Удаление paragraph-level sectPr | `contentdata.go:422-425` | 9/10 |
| Перенумерация Drawing ID (docPr/cNvPr) | `contentdata.go:459-473` | 9/10 |
| Обработка в headers/footers с дедупликацией StoryPart | `document.go:606-626` | 9/10 |
| Обработка в comments | `document.go:628-633` | 8/10 |
| Рекурсия в ячейки таблиц | `blkcntnr.go:252-264` | 9/10 |
| Разбиение параграфов на фрагменты (SplitParagraphAtTags) | `oxml/replacetable.go` | 9/10 |
| Cross-run tag detection | `oxml/replacetext.go` | 9/10 |
| Source document не модифицируется (deep copy через etree.Copy) | `contentdata.go:287-289` | 10/10 |
| Idempotent import phases (flags) | `resourceimport.go:34,39,50,51` | 9/10 |
| elementBuilder pattern — fresh Copy() на каждый placeholder | `blkcntnr.go:375-388` | 10/10 |
| 20+ integration tests | `replacecontent_test.go` | 8/10 |

### Gap Analysis vs Aspose.Words

| # | Gap | Критичность | Aspose эквивалент |
|---|---|---|---|
| G1 | **ImportFormatMode** — только UseDestinationStyles, нет KeepSourceFormatting и KeepDifferentStyles | HIGH | `ImportFormatMode` enum + `ImportFormatOptions` |
| G2 | **Глубокий импорт generic parts** — chart/VML sub-relationships не копируются | HIGH | `NodeImporter` deep copy |
| G3 | **Защита от порчи при ошибке** — при ошибке документ частично модифицирован | HIGH | `Document.Clone()` перед операцией |
| G4 | **ImportFormatOptions** — нет SmartStyleBehavior, ForceCopyStyles, KeepSourceNumbering, MergePastedLists | MEDIUM | `ImportFormatOptions` class |
| G5 | **OLE Objects** — embedded objects копируются как blob без обработки preview image | MEDIUM | OLE Shape import |
| G6 | **Structured Document Tags** — SDT elements проходят без обработки | LOW | SDT node import (as-is) |
| G7 | **Field codes** — поля проходят без обработки | LOW | Field preservation (as-is) |

**Важные выводы из ревью Aspose.Words:**

1. **Aspose НЕ делает rollback** — рекомендованный подход: `Document.Clone()` до операции
2. **Aspose НЕ мержит font tables** — шрифты остаются как есть, resolution при рендере
3. **Aspose НЕ имеет Document.Validate()** — нет post-import валидации
4. **Aspose НЕ ребиндит SDT** — структура сохраняется, data bindings могут ломаться
5. **KeepSourceFormatting по умолчанию** — раскрывает стиль в direct attributes + Normal, НЕ переименовывает
6. **ForceCopyStyles** — отдельная опция, только она создаёт стили с суффиксами `_0`, `_1`
7. **Field codes** — сохраняются as-is, Word пересчитывает при открытии
8. **Theme** — document-level Theme НЕ копируется/мержится; один theme на документ

---

## 1. Фаза 1 — Защита от порчи документа при ошибке

**Приоритет:** CRITICAL
**Aspose-подход:** `Document.Clone()` перед операцией

### 1.1 Подход: Clone перед мутацией

Aspose.Words не реализует сложный snapshot/rollback. Вместо этого рекомендует:
```
backup = document.Clone()
try:
    document.insert(...)
catch:
    document = backup
```

Для go-docx аналогичный подход — **pre-operation safety copy**:

```go
// SafeReplaceWithContent performs ReplaceWithContent with rollback guarantee.
// On error, the document is guaranteed to remain in its pre-call state.
func (d *Document) SafeReplaceWithContent(old string, cd ContentData) (int, error) {
    // TODO: future implementation — save body XML + rels state, restore on error
    // For now, delegate to ReplaceWithContent as-is.
    return d.ReplaceWithContent(old, cd)
}
```

**Реализация — lazy snapshot через etree.Copy:**

Snapshot НЕ через сериализацию в []byte (дорого). Используем то, что уже работает в проекте — `etree.Element.Copy()`:

```
Файл: pkg/docx/document.go (расширение ReplaceWithContent)
```

```go
type bodySnapshot struct {
    bodyEl    *etree.Element         // deep copy of <w:body>
    storyRels map[*parts.StoryPart][]opc.Relationship // cloned rels
}

func (d *Document) snapshotBody() (*bodySnapshot, error) {
    b, err := d.getBody()
    if err != nil {
        return nil, err
    }
    return &bodySnapshot{
        bodyEl: b.Element().Copy(),
    }, nil
}

func (d *Document) restoreBody(snap *bodySnapshot) {
    body, _ := d.getBody()
    parent := body.Element().Parent()
    parent.RemoveChild(body.Element())
    parent.AddChild(snap.bodyEl)
    d.body = nil // invalidate cache
}
```

**Scope:** Только body и rels. Headers/footers/comments — каждый свой snapshot перед обработкой.

**Важно:** `OpcPackage` не имеет `RemovePart()`. Если при ошибке были добавлены image/generic parts — они остаются в пакете как orphan parts. Это **допустимо**: orphan parts не влияют на корректность документа, Word их игнорирует, и при re-save они могут быть очищены. Aspose.Words ведёт себя аналогично.

### 1.2 Альтернатива для критических сценариев

Для случаев, когда даже orphan parts недопустимы — save/reopen pattern:

```go
// Пользовательский код:
buf, _ := doc.SaveToBytes()     // snapshot
count, err := doc.ReplaceWithContent(tag, cd)
if err != nil {
    doc, _ = OpenBytes(buf)     // полный rollback
}
```

Документируем этот паттерн в godoc метода.

---

## 2. Фаза 2 — ImportFormatMode (3 стратегии слияния стилей)

**Приоритет:** HIGH
**Цель:** Полный паритет с `Aspose.Words.ImportFormatMode`

### 2.1 Enum и Options

```
Файл: pkg/docx/contentdata.go (расширение)
```

```go
type ImportFormatMode int

const (
    // UseDestinationStyles — текущее поведение. Если стиль с тем же ID
    // существует в target, используется target-определение.
    // Отсутствующие стили копируются из source.
    UseDestinationStyles ImportFormatMode = iota

    // KeepSourceFormatting — форматирование source сохраняется.
    // При конфликте ID: formatting source стиля раскрывается в
    // direct attributes на элементах, ссылка меняется на Normal.
    // Это поведение Aspose.Words по умолчанию для этого режима.
    KeepSourceFormatting

    // KeepDifferentStyles — гибрид. Если стили с одинаковым ID имеют
    // ОДИНАКОВОЕ форматирование → используется target (как UseDestination).
    // Если РАЗНОЕ → formatting раскрывается в direct attributes (как KeepSource).
    KeepDifferentStyles
)
```

### 2.2 ImportFormatOptions

```go
type ImportFormatOptions struct {
    // ForceCopyStyles — при KeepSourceFormatting конфликтующие стили
    // копируются с суффиксом _0, _1 вместо раскрытия в direct attributes.
    // Аналог Aspose.Words ImportFormatOptions.ForceCopyStyles.
    // Default: false.
    ForceCopyStyles bool

    // KeepSourceNumbering — при конфликте numId source нумерация
    // сохраняется как отдельный список (новый abstractNum).
    // false (default) = нумерация вливается в существующий список target.
    // Аналог Aspose.Words ImportFormatOptions.KeepSourceNumbering.
    // Default: false (текущее поведение проекта = true по факту,
    // т.к. всегда создаётся новый abstractNum).
    KeepSourceNumbering bool

    // MergePastedLists — вставленные списки объединяются с
    // окружающими списками того же типа.
    // Аналог Aspose.Words ImportFormatOptions.MergePastedLists.
    // Default: false.
    MergePastedLists bool
}
```

### 2.3 Расширение ContentData

```go
type ContentData struct {
    Source  *Document
    Format ImportFormatMode    // default: UseDestinationStyles (zero value)
    Options ImportFormatOptions // default: all false (zero value)
}
```

**Обратная совместимость:** Zero value `ContentData{Source: doc}` = точно текущее поведение.

### 2.4 Реализация KeepSourceFormatting (без ForceCopyStyles)

**Aspose.Words behavior:** При конфликте стилей formatting source стиля раскрывается
в direct paragraph/run properties, а ссылка на стиль меняется на "Normal" (или target default).

```
Файл: pkg/docx/resourceimport_styles.go (расширение mergeOneStyle)
```

```go
func (ri *ResourceImporter) mergeOneStyle(srcStyle *oxml.CT_Style) error {
    id := srcStyle.StyleId()
    // ...existing checks...

    tgtStyles, _ := ri.targetStyles()

    switch ri.importFormatMode {
    case UseDestinationStyles:
        // Текущее поведение — без изменений
        if tgtStyles.GetByID(id) != nil {
            ri.styleMap[id] = id
            return nil
        }
        // Copy missing style...

    case KeepSourceFormatting:
        if tgtStyles.GetByID(id) == nil {
            // Стиль отсутствует в target — копируем из source
            ri.copyStyleToTarget(srcStyle)
            ri.styleMap[id] = id
            return nil
        }
        if ri.opts.ForceCopyStyles {
            // Конфликт + ForceCopyStyles → копируем с суффиксом
            newId := ri.uniqueStyleId(id)
            ri.copyStyleToTargetAs(srcStyle, newId)
            ri.styleMap[id] = newId
            return nil
        }
        // Конфликт + default → expand to direct attributes
        // Записываем маркер для expandDirectFormatting
        ri.expandStyles[id] = srcStyle
        ri.styleMap[id] = ri.targetDefaultParagraphStyleId()
        return nil

    case KeepDifferentStyles:
        existing := tgtStyles.GetByID(id)
        if existing == nil {
            ri.copyStyleToTarget(srcStyle)
            ri.styleMap[id] = id
            return nil
        }
        if stylesContentEqual(srcStyle, existing) {
            // Одинаковые → используем target
            ri.styleMap[id] = id
            return nil
        }
        // Разные → expand to direct attributes (как KeepSourceFormatting)
        ri.expandStyles[id] = srcStyle
        ri.styleMap[id] = ri.targetDefaultParagraphStyleId()
        return nil
    }
}
```

### 2.5 Expand to Direct Attributes

Ключевая логика KeepSourceFormatting — раскрытие стиля в direct formatting:

```go
// expandDirectFormatting applies source style properties directly
// to elements that referenced this style. Called after remapAll
// changes style references to Normal/default.
//
// For each paragraph referencing the conflicting style:
//   1. Resolve full formatting chain (style + basedOn + basedOn...)
//   2. Apply resolved pPr/rPr properties directly to the paragraph
//   3. Style reference already changed to Normal by remapAll
//
// This mirrors Aspose.Words KeepSourceFormatting behavior:
// "the source style formatting is expanded into direct Node
// attributes and the style is changed to Normal."
func (ri *ResourceImporter) expandDirectFormatting(elements []*etree.Element) {
    if len(ri.expandStyles) == 0 {
        return
    }
    for _, root := range elements {
        // Walk paragraphs referencing expanded styles
        // Merge resolved pPr/rPr into element's existing properties
    }
}
```

**Pipeline расширяется:**
```
existing:  sanitize → materialize → remapAll → import rIds
new:       sanitize → materialize → remapAll → expandDirectFormatting → import rIds
```

### 2.6 uniqueStyleId (для ForceCopyStyles)

```go
func (ri *ResourceImporter) uniqueStyleId(base string) string {
    tgtStyles, _ := ri.targetStyles()
    for i := 0; ; i++ {
        candidate := fmt.Sprintf("%s_%d", base, i)
        if tgtStyles.GetByID(candidate) == nil {
            // Также проверяем, что мы сами не создали такой ID
            if _, used := ri.styleMap[candidate]; !used {
                return candidate
            }
        }
    }
}
```

### 2.7 Сравнение стилей (KeepDifferentStyles)

```go
// stylesContentEqual compares two styles by their formatting content,
// ignoring w:name (which may differ for functionally identical styles).
func stylesContentEqual(a, b *oxml.CT_Style) bool {
    ac := a.RawElement().Copy()
    bc := b.RawElement().Copy()
    // Remove w:name — may differ for same formatting
    removeChildByTag(ac, "w", "name")
    removeChildByTag(bc, "w", "name")
    // Remove w:rsid* — revision identifiers, not formatting
    removeRsidAttrs(ac)
    removeRsidAttrs(bc)
    // Canonical comparison
    var bufA, bufB bytes.Buffer
    ac.WriteTo(&bufA, &etree.WriteSettings{CanonicalText: true})
    bc.WriteTo(&bufB, &etree.WriteSettings{CanonicalText: true})
    return bytes.Equal(bufA.Bytes(), bufB.Bytes())
}
```

---

## 3. Фаза 3 — Глубокий импорт generic parts (Charts, VML, Diagrams)

**Приоритет:** HIGH
**Текущая проблема:** `contentdata.go:236-238`:
> "Sub-relationships of the source part are NOT imported (shallow copy)"

Вставленные charts, SmartArt, VML drawings теряют свои данные (embedded Excel, color/style XML).

### 3.1 Рекурсивный Part Import

```
Файл: pkg/docx/contentdata.go (замена importGenericPart)
```

**Текущий shallow pipeline:**
```
srcPart.Blob() → newPart → AddPart → Add relationship
// Sub-relationships потеряны!
```

**Enterprise deep pipeline:**
```
1. srcPart.Blob() → newPart → AddPart → Add relationship
2. Если srcPart имеет Rels():
   2a. Для каждого sub-relationship:
       - External → GetOrAddExtRel on newPart.Rels()
       - Image → GetOrAddImagePart (SHA-256 dedup) → relate to newPart
       - Internal → RECURSE (deep import sub-part)
   2b. Переписать rIds в XML blob newPart
   2c. Сохранить переписанный blob
```

```go
func importGenericPartDeep(
    srcRel *opc.Relationship,
    targetPart *parts.StoryPart,
    targetPkg *parts.WmlPackage,
    importedParts map[opc.PackURI]opc.Part,
    depth int,
) (string, error) {
    if depth > maxPartImportDepth {
        return "", fmt.Errorf("docx: generic part import depth exceeds %d (circular reference?)", maxPartImportDepth)
    }

    srcPart := srcRel.TargetPart
    srcPN := srcPart.PartName()

    // Dedup check
    if existing, ok := importedParts[srcPN]; ok {
        rel := targetPart.Rels().GetOrAdd(srcRel.RelType, existing)
        return rel.RID, nil
    }

    // Copy blob
    blob, err := srcPart.Blob()
    if err != nil {
        return "", fmt.Errorf("reading part blob: %w", err)
    }

    template := partNameTemplate(srcPN)
    newPN := targetPkg.OpcPackage.NextPartname(template)
    newPart := opc.NewBasePart(newPN, srcPart.ContentType(), blob, targetPkg.OpcPackage)
    newPart.SetRels(opc.NewRelationships(newPN.BaseURI()))
    targetPkg.OpcPackage.AddPart(newPart)
    importedParts[srcPN] = newPart

    // NEW: import sub-relationships
    srcSubRels := srcPart.Rels()
    if srcSubRels != nil && len(srcSubRels.All()) > 0 {
        subRidMap := make(map[string]string)

        for _, subRel := range srcSubRels.All() {
            if subRel.IsExternal {
                newRId := newPart.Rels().GetOrAddExtRel(subRel.RelType, subRel.TargetRef)
                subRidMap[subRel.RID] = newRId
                continue
            }
            if subRel.TargetPart == nil {
                continue
            }
            if subRel.RelType == opc.RTImage {
                // Image sub-rel — use existing dedup
                newRId, err := importImageRel(subRel, newPart, targetPkg)
                if err != nil {
                    return "", err
                }
                subRidMap[subRel.RID] = newRId
                continue
            }
            // Recursive generic sub-part
            // Wrap newPart as a temporary story-like target for Rels()
            newSubRId, err := importGenericSubPart(subRel, newPart, targetPkg, importedParts, depth+1)
            if err != nil {
                return "", err
            }
            subRidMap[subRel.RID] = newSubRId
        }

        // Rewrite rIds in the blob if it's XML
        if isXmlContentType(newPart.ContentType()) && len(subRidMap) > 0 {
            rewritten, err := rewriteRIdsInXmlBlob(blob, subRidMap)
            if err == nil {
                newPart.SetBlob(rewritten)
            }
            // If rewrite fails (non-XML blob), keep original — graceful degradation
        }
    }

    // Create relationship from caller
    targetRef := newPN.RelativeRef(targetPart.Rels().BaseURI())
    rel := targetPart.Rels().Add(srcRel.RelType, targetRef, newPart, false)
    return rel.RID, nil
}

const maxPartImportDepth = 10
```

### 3.2 Необходимые новые методы в OPC layer

```go
// В opc/part.go — BasePart:
func (p *BasePart) SetBlob(data []byte)   // уже может быть, проверить
func (p *BasePart) Rels() *Relationships  // нужен getter для sub-rels

// В opc/helpers.go:
func isXmlContentType(ct string) bool {
    return strings.HasSuffix(ct, "+xml") || strings.HasSuffix(ct, "/xml")
}
```

### 3.3 XML Blob rId Rewriting

```go
// rewriteRIdsInXmlBlob parses an XML blob, replaces relationship
// reference attributes (r:id, r:embed, r:link) according to ridMap,
// and re-serializes.
func rewriteRIdsInXmlBlob(blob []byte, ridMap map[string]string) ([]byte, error) {
    doc := etree.NewDocument()
    if err := doc.ReadFromBytes(blob); err != nil {
        return nil, err
    }
    remapRIds(doc.Root().ChildElements(), ridMap)
    // Also remap on root element itself
    for i, attr := range doc.Root().Attr {
        if isRelAttr(attr) {
            if newVal, ok := ridMap[attr.Value]; ok {
                doc.Root().Attr[i].Value = newVal
            }
        }
    }
    return doc.WriteToBytes()
}
```

### 3.4 Chart Import — типичная структура

```
/word/charts/chart1.xml          ← main chart XML
  └→ /word/charts/style1.xml     ← chart style (sub-relationship)
  └→ /word/charts/colors1.xml    ← chart colors (sub-relationship)
  └→ /word/embeddings/wb1.xlsx   ← data source (sub-relationship, binary)
```

Recursive import обработает все 4 файла автоматически. Binary sub-parts (xlsx) копируются как blob без rId rewrite (не XML content type → graceful skip).

### 3.5 Cycle Detection

Dedup map `importedParts` уже предотвращает дублирование. Для protection от circular references (part A → part B → part A) используется depth limit. Дополнительный visiting set не нужен, т.к. dedup map обеспечивает идемпотентность.

---

## 4. Фаза 4 — ImportFormatOptions (расширенные опции)

**Приоритет:** MEDIUM
**Цель:** Паритет с `Aspose.Words.ImportFormatOptions`

### 4.1 KeepSourceNumbering

**Текущее поведение проекта** (`resourceimport_num.go`): всегда создаёт новый abstractNum с новым nsid — фактически это `KeepSourceNumbering = true`.

**Для паритета с Aspose нужен режим `KeepSourceNumbering = false`:**
- При конфликте numId → source нумерация вливается в существующий target список
- Source list items продолжают target нумерацию (e.g., target: 1-5, source: 6-10)

```go
func (ri *ResourceImporter) importNumbering() error {
    // ...existing...
    if !ri.opts.KeepSourceNumbering {
        // Try to find matching abstractNum in target by list style
        // If found, reuse target numId
        tgtNumId := ri.findMatchingNumId(srcAbsNum)
        if tgtNumId > 0 {
            ri.numIdMap[srcNumId] = tgtNumId
            continue
        }
    }
    // ...existing: create new abstractNum...
}
```

### 4.2 MergePastedLists

При `MergePastedLists = true`, вставленные параграфы с нумерацией объединяются с предшествующим/последующим списком того же стиля:

```go
// Вызывается после замены, обходит вставленные параграфы
// и при совпадении стиля списка с соседями — объединяет numId.
func (ri *ResourceImporter) mergePastedLists(elements []*etree.Element, container *etree.Element) {
    // Find preceding/following paragraphs in container
    // If same list style → change numId to match
}
```

### 4.3 SmartStyleBehavior

При `KeepSourceFormatting` + `SmartStyleBehavior = true`:
- Source стили с тем же именем что и в target → конвертируются в direct paragraph attributes
- Только для **paragraph styles** (не character, не table)

Реализация: подмножество expand-to-direct-attributes, ограниченное paragraph-level.

---

## 5. Фаза 5 — OLE Objects

**Приоритет:** MEDIUM
**Aspose-поведение:** OLE objects сохраняются as-is при импорте

### 5.1 OLE Object Structure in OOXML

```xml
<w:object>
  <v:shape style="width:...;height:...">
    <v:imagedata r:id="rId5"/>    <!-- preview image -->
  </v:shape>
  <o:OLEObject r:id="rId6" .../>  <!-- embedded binary -->
</w:object>
```

**Два relationship:**
1. `r:id` на `v:imagedata` → preview image (WMF/EMF/PNG)
2. `r:id` на `o:OLEObject` → embedded binary (.bin / .xlsx / .docx)

### 5.2 Текущее покрытие

`collectReferencedRIds` уже сканирует **все** элементы depth-first и собирает `r:id`, `r:embed`, `r:link`. Оба OLE relationship уже подхватываются.

`importRelationship` обработает:
- Preview image (Case 2: RTImage) → SHA-256 dedup, correct
- Embedded binary — Case 3 (generic part) → **shallow copy only**

### 5.3 Требуемое изменение

С реализацией Фазы 3 (deep generic part import) OLE embedded binary автоматически получит deep import. Preview image уже работает.

**Дополнительно:** Embedded OLE objects могут иметь свои sub-relationships (e.g., embedded Excel → chart в Excel). Но Aspose.Words тоже не рекурсирует в embedded OLE — сохраняет blob as-is. Наш deep import на уровне generic part (blob copy + sub-rels) уже достаточен.

**Вывод:** OLE покрывается Фазой 3 автоматически. Отдельная фаза не нужна.

---

## 6. Фаза 6 — SDT и Field Codes (сохранение as-is)

**Приоритет:** LOW
**Aspose-поведение:** Оба сохраняются as-is при импорте

### 6.1 SDT — текущее состояние

SDT elements (`<w:sdt>`) уже корректно обрабатываются текущим кодом:
- Deep copy через `etree.Copy()` — структура сохраняется
- rId внутри SDT подхватываются `collectReferencedRIds` (DFS)
- Style references внутри SDT подхватываются `collectStyleIdsFromElements`
- Data bindings (`<w:dataBinding>`) — сохраняются as-is (как в Aspose)

**Единственное требование:** если SDT содержит dataBinding на CustomXmlPart — привязка может не работать в target. Это **ожидаемое поведение** (Aspose ведёт себя так же — пользователь должен сам клонировать CustomXmlParts).

**Вывод:** Дополнительная обработка SDT НЕ нужна. Текущее поведение уже корректно.

### 6.2 Field Codes — текущее состояние

Field codes бывают двух видов:
1. `<w:fldSimple>` — inline field
2. `<w:fldChar>` + `<w:instrText>` + `<w:fldChar>` — complex field

Оба уже корректно копируются через `etree.Copy()`. Field results (cached display text) сохраняются. Word пересчитает при открытии.

**Единственный edge case:** `INCLUDEPICTURE` field может содержать путь к локальному файлу. Но Aspose.Words тоже не обрабатывает это специально.

**Вывод:** Дополнительная обработка field codes НЕ нужна. Текущее поведение уже корректно.

---

## 7. Фаза 7 — Расширенное тестирование

**Приоритет:** CRITICAL
**Текущие тесты:** 20 интеграционных в `replacecontent_test.go`
**Цель:** 60+ тестов, покрывающих все сценарии включая новые фичи

### 7.1 Тесты для ImportFormatMode

| Тест | Описание |
|---|---|
| `TestRWC_UseDestination_ExistingStyleKept` | Стиль в target не перезаписывается (текущее поведение) |
| `TestRWC_UseDestination_MissingStyleCopied` | Отсутствующий стиль копируется из source |
| `TestRWC_KeepSource_ConflictExpanded` | Конфликтующий стиль → direct attributes + Normal |
| `TestRWC_KeepSource_NoConflictCopied` | Неконфликтующий стиль копируется как есть |
| `TestRWC_KeepSource_ForceRename` | ForceCopyStyles=true → стиль с суффиксом `_0` |
| `TestRWC_KeepSource_ForceRenameChain` | basedOn цепочка конфликтует → все переименованы |
| `TestRWC_KeepDifferent_IdenticalUseTarget` | Одинаковые стили → target используется |
| `TestRWC_KeepDifferent_DifferentExpanded` | Разные стили → expanded to direct |
| `TestRWC_BackwardCompat_ZeroValue` | `ContentData{Source: doc}` = UseDestinationStyles |

### 7.2 Тесты для Deep Generic Part Import

| Тест | Описание |
|---|---|
| `TestRWC_GenericPart_SubRelsImported` | Sub-relationships копируются |
| `TestRWC_GenericPart_SubRidsRemapped` | rIds в XML blob перемаплены |
| `TestRWC_GenericPart_BinarySubPart` | Non-XML sub-part (xlsx) → blob copy, no rId rewrite |
| `TestRWC_GenericPart_DepthLimit` | Глубина > 10 → ошибка |
| `TestRWC_GenericPart_Dedup` | Один source part → один target part при множественных ref |
| `TestRWC_GenericPart_RoundTrip` | Save → reopen → sub-parts intact |

### 7.3 Тесты для ошибочных сценариев

| Тест | Описание |
|---|---|
| `TestRWC_ErrorRecovery_BodyIntact` | Ошибка в header → body не испорчен |
| `TestRWC_NilSource_Error` | Source=nil → ошибка (существующий тест) |
| `TestRWC_EmptyOld_NoOp` | old="" → 0 замен (существующий тест) |

### 7.4 Тесты для ImportFormatOptions

| Тест | Описание |
|---|---|
| `TestRWC_KeepSourceNumbering_SeparateList` | KeepSourceNumbering=true → отдельный список |
| `TestRWC_KeepSourceNumbering_MergedList` | KeepSourceNumbering=false → нумерация продолжается |
| `TestRWC_MergePastedLists` | Вставленный список объединяется с окружающим |

### 7.5 Stress / Edge Cases

| Тест | Описание |
|---|---|
| `TestRWC_LargeDocument_1000Paragraphs` | Производительность |
| `TestRWC_DeeplyNestedTables_5Levels` | 5 уровней вложенности |
| `TestRWC_100Placeholders_SameTag` | 100 вхождений одного тега |
| `TestRWC_UnicodeTag` | Кириллица, CJK в теге |
| `TestRWC_TagInHyperlinkText` | Тег внутри гиперссылки |
| `TestRWC_MultipleSourcesSameImage` | Два source с одним изображением → dedup |
| `TestRWC_SourceWithFootnoteContainingImage` | Сноска с изображением |
| `TestRWC_SDT_Preserved` | SDT elements проходят через Copy корректно |
| `TestRWC_FieldCode_Preserved` | Field codes (fldSimple, fldChar) сохраняются |

### 7.6 Benchmarks

```go
func BenchmarkReplaceWithContent_Simple(b *testing.B)
func BenchmarkReplaceWithContent_WithImages(b *testing.B)
func BenchmarkReplaceWithContent_100Replacements(b *testing.B)
func BenchmarkReplaceWithContent_KeepSourceFormatting(b *testing.B)
func BenchmarkReplaceWithContent_DeepGenericParts(b *testing.B)
```

### 7.7 Fuzz Testing

```go
func FuzzReplaceWithContent_TagDiscovery(f *testing.F) {
    f.Add("[[test]]")
    f.Add("[<CONTENT>]")
    f.Add("{{tag}}")
    f.Fuzz(func(t *testing.T, tag string) {
        if len(tag) == 0 || len(tag) > 100 {
            return
        }
        target := mustNewDoc(t)
        target.AddParagraph("prefix " + tag + " suffix")
        source := mustNewDoc(t)
        source.AddParagraph("inserted")
        // Must not panic
        target.ReplaceWithContent(tag, ContentData{Source: source})
    })
}
```

---

## 8. Порядок реализации и зависимости

```
Phase 1 (Error protection)          ← фундамент безопасности
    ↓
Phase 2 (ImportFormatMode)          ← основная бизнес-фича
    ↓
Phase 3 (Deep generic parts)        ← независима от Phase 2
    ↓
Phase 4 (ImportFormatOptions)       ← зависит от Phase 2
    ↓
Phase 7 (Tests for everything)      ← последняя, покрывает всё
```

**Фазы 5, 6 (SDT, Fields) — НЕ НУЖНЫ:** текущее поведение уже корректно (as-is, как в Aspose).

### Рекомендуемый порядок спринтов

| Спринт | Фаза | Результат | Файлы |
|---|---|---|---|
| S1 | Phase 1 | Error protection + body snapshot | `document.go` |
| S2 | Phase 2 | 3 ImportFormatMode + ForceCopyStyles | `contentdata.go`, `resourceimport_styles.go` |
| S3 | Phase 3 | Deep generic part import | `contentdata.go` |
| S4 | Phase 4 | KeepSourceNumbering, MergePastedLists | `resourceimport_num.go` |
| S5 | Phase 7 | 60+ тестов, benchmarks, fuzz | `replacecontent_test.go`, новые test files |

---

## 9. Файловая структура после реализации

```
pkg/docx/
├── contentdata.go                  ← расширен (ImportFormatMode, ImportFormatOptions, deep import)
├── resourceimport.go               ← расширен (expandStyles map, opts field)
├── resourceimport_styles.go        ← расширен (3 ImportFormatMode, expand-to-direct, stylesContentEqual)
├── resourceimport_num.go           ← расширен (KeepSourceNumbering, MergePastedLists)
├── resourceimport_notes.go         ← без изменений
├── document.go                     ← расширен (snapshotBody, restoreBody)
├── blkcntnr.go                     ← без изменений
├── replacecontent_test.go          ← расширен (60+ тестов)
├── replacecontent_bench_test.go    ← НОВЫЙ (benchmarks)
├── replacecontent_fuzz_test.go     ← НОВЫЙ (fuzz tests)
```

**Новых файлов:** 2 (только тесты)
**Изменённых файлов:** 5
**Без изменений:** Все остальные файлы библиотеки

---

## 10. Что НЕ нужно делать (anti-requirements)

Это важная секция. Следующие фичи были в первоначальном плане, но **удалены после ревью** как не соответствующие реальному поведению Aspose.Words или избыточные:

| Убранная фича | Причина |
|---|---|
| ~~Theme import / Theme color resolution~~ | Aspose НЕ мержит theme между документами. Один theme на документ. При KeepSourceFormatting formatting раскрывается в direct attributes — этого достаточно |
| ~~Font table merge~~ | Aspose НЕ мержит font tables при импорте. Шрифты resolution — задача рендера, не import. Имена шрифтов сохраняются as-is |
| ~~Post-import Document.Validate()~~ | У Aspose НЕТ такого API. Практический подход — save/reopen для проверки round-trip. Валидация — отдельная фича, не часть ReplaceWithContent |
| ~~SDT rebinding~~ | Aspose сохраняет SDT as-is, bindings могут ломаться. Наш Copy() делает то же самое — уже корректно |
| ~~Field code remap/dirty~~ | Aspose сохраняет fields as-is. Word пересчитывает. Наш Copy() уже корректен |
| ~~CustomXml parts import~~ | Aspose требует ручного клонирования CustomXmlParts. Не задача import engine |
| ~~Document.Clone() для rollback~~ | Слишком дорого для каждого вызова. snapshotBody через Copy() — достаточно для body. Save/reopen паттерн документируем для полного rollback |
| ~~Snapshot для OpcPackage (RemovePart)~~ | RemovePart() не существует. Orphan parts безвредны и игнорируются Word |
| ~~Progress callbacks~~ | У Aspose нет progress callback для import. Over-engineering |
| ~~Glossary document merge~~ | Крайне редкий use case, Aspose поддерживает но не документирует |

---

## 11. Критерии приёмки (Enterprise Definition of Done)

| # | Критерий | Метрика |
|---|---|---|
| DoD-1 | Все 3 ImportFormatMode работают | 9+ тестов на import format |
| DoD-2 | ForceCopyStyles создаёт стили с суффиксами | Тест с проверкой ID |
| DoD-3 | KeepSourceFormatting раскрывает в direct attrs | Тест: проверка pPr/rPr attributes |
| DoD-4 | Deep generic parts: sub-rels imported | Round-trip тест с mock chart |
| DoD-5 | При ошибке body не испорчен | Тест с forced error |
| DoD-6 | 60+ тестов, все зелёные | `go test ./...` |
| DoD-7 | Benchmarks: < 1.5x overhead на UseDestination | `go test -bench .` |
| DoD-8 | Fuzz: 0 паник за 1 минуту | `go test -fuzz . -fuzztime 1m` |
| DoD-9 | `ContentData{Source: doc}` = текущее поведение | Существующие 20 тестов зелёные |
| DoD-10 | Visual regression: SSIM > 0.98 | visual-regtest pipeline |

---

## 12. Сравнительная таблица: Before / After / Aspose

| Возможность | До (текущее) | После (plan v2) | Aspose.Words |
|---|---|---|---|
| Import format modes | 1 (UseDestination) | 3 + ForceCopyStyles | 3 + ForceCopyStyles |
| Style conflict: default | Skip (use target) | Expand to direct attrs | Expand to direct attrs |
| Style conflict: force | N/A | Rename with `_0` suffix | Rename with `_0` suffix |
| KeepDifferentStyles | N/A | Compare → expand if diff | Compare → expand if diff |
| Deep part import | Shallow (blob only) | Recursive + rId remap | Full |
| KeepSourceNumbering | Always keep (hardcoded) | Configurable | Configurable |
| MergePastedLists | N/A | Configurable | Configurable |
| Error recovery | Partial state | Body snapshot + restore | Document.Clone() |
| SDT handling | Pass-through (Copy) | Pass-through (Copy) | Pass-through |
| Field codes | Pass-through (Copy) | Pass-through (Copy) | Pass-through |
| Theme merge | N/A | N/A | N/A |
| Font table merge | N/A | N/A | N/A |
| Validation | N/A | N/A | N/A |
| Test coverage | 20 tests | 60+ tests | Internal |
| Benchmarks | None | 5 scenarios | Internal |
| Fuzz testing | None | Tag discovery fuzz | N/A |
