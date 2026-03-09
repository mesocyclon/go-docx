# Aspose.Words Compliance Plan — ReplaceWithContent

Приведение `ReplaceWithContent` к полному соответствию поведению Aspose.Words.
Только **улучшения** — никаких деградаций существующей функциональности.

Каждая задача самодостаточна, может быть реализована и протестирована независимо.

---

## Задача 1: KeepDifferentStyles — copy вместо expand

**Приоритет:** Критический
**Файлы:** `resourceimport_styles.go`, `contentdata.go`
**Тесты:** `replacecontent_test.go`, `resourceimport_styles_test.go`

### Проблема

При `KeepDifferentStyles` и **разном** форматировании go-docx ведёт себя как
`KeepSourceFormatting` — разворачивает стиль в прямые атрибуты. Aspose.Words
в этом случае **всегда копирует** стиль с суффиксом `_N`.

```
Aspose:   different → copy with suffix (always)
go-docx:  different → expand to direct attrs (default) / copy with suffix (ForceCopyStyles)
```

`ForceCopyStyles` в Aspose документирован **только для KeepSourceFormatting**.
Для `KeepDifferentStyles` он не нужен, потому что копирование происходит всегда.

### Что делать

В `mergeOneStyle` (resourceimport_styles.go:261–272), ветка
`KeepDifferentStyles`, блок `else` (стили различаются).

**Текущий код:**

```go
} else {
    // Different formatting → behave like KeepSourceFormatting.
    if ri.opts.ForceCopyStyles {
        return ri.copyStyleToTarget(srcStyle, ri.uniqueStyleId(id))
    }
    ri.expandStyles[id] = srcStyle
    ri.styleMap[id] = ri.targetDefaultParaStyleId()
}
```

**Целевой код:**

```go
} else {
    // Aspose.Words KeepDifferentStyles: different formatting →
    // always copy with unique suffix. ForceCopyStyles is irrelevant
    // for this mode (it's a KeepSourceFormatting-only option).
    return ri.copyStyleToTarget(srcStyle, ri.uniqueStyleId(id))
}
```

Инфраструктура уже готова:
- `copyStyleToTarget` (строка 285) — deep-copy + strip default + rename
- `uniqueStyleId` (строка 362) — генерация суффикса `_0`, `_1`, ...
- Pass 2 `fixupCopiedStyles` (строка 1203) — автоматически обработает
  новые клоны: remap basedOn/next/link + compensateDocDefaults
- `copiedClones` / `copiedStyleIds` — автоматически заполняются
  в `copyStyleToTarget` (строки 328–335)

### Обновление docstring KeepDifferentStyles

В `contentdata.go:54–59` комментарий станет неверным после изменения.

**Было:**
```go
//   - If formatting differs → behaves like KeepSourceFormatting (expands
//     to direct attributes, or copies with suffix if ForceCopyStyles).
```

**Стало:**
```go
//   - If formatting differs → always copies the source style with a
//     unique suffix (_0, _1, ...). ForceCopyStyles is irrelevant for
//     this mode (it's a KeepSourceFormatting-only option).
```

### Обновление существующего теста

Тест `TestRWC_KeepDifferent_DifferentExpanded` (replacecontent_test.go:2190)
проверяет **старое** expand-поведение: ожидает `jc=center` в прямых
атрибутах параграфа. После изменения этот тест **сломается**.

**Действие:** переписать `TestRWC_KeepDifferent_DifferentExpanded`:
- Переименовать в `TestRWC_KeepDifferent_DifferentAlwaysCopied`
- Вместо проверки `jc=center` в direct attrs проверить:
  - `DiffStyle_0` создан в target styles
  - `semiHidden` на скопированном стиле
  - Параграф ссылается на `DiffStyle_0` (а не на default)
  - Прямые атрибуты `jc` **не** инжектированы

### Граничные случаи

1. **Стиль отсутствует в target** → без изменений (copy as-is, все 3 режима)
2. **Стиль одинаковый** → без изменений (reuse target)
3. **Стиль разный, basedOn родитель тоже разный** → оба копируются,
   Pass 2 remaps basedOn в скопированный родитель
4. **Стиль разный, basedOn на стиль целевого документа** →
   копия получает compensateUncovered дельту docDefaults
5. **Стиль разный, корневой (без basedOn)** → копия получает compensateAll
6. **Стиль разный, character type** → копируется, docDefaults rPr компенсация
7. **Несколько стилей с разным форматированием** → каждый получает
   уникальный суффикс (_0, _1, _2)
8. **ForceCopyStyles=true + KeepDifferentStyles** → поведение идентично
   ForceCopyStyles=false (всегда copy), опция игнорируется

### Тесты

Расположение: `replacecontent_test.go`

```
TestRWC_KeepDifferent_DifferentAlwaysCopied
    (переписанный TestRWC_KeepDifferent_DifferentExpanded)
    Стиль DiffStyle с разным форматированием, без ForceCopyStyles →
    стиль скопирован как DiffStyle_0, semiHidden, параграф ссылается
    на DiffStyle_0. Прямые атрибуты jc НЕ инжектированы.

TestRWC_KeepDifferent_SameReusesTarget
    Стиль с одинаковым форматированием в src и tgt →
    стиль НЕ копируется, параграф ссылается на оригинальный target styleId.

TestRWC_KeepDifferent_MissingCopiedAsIs
    Стиль отсутствует в target →
    стиль копируется под исходным ID (без суффикса).

TestRWC_KeepDifferent_ForceCopyIgnored
    ForceCopyStyles=true + разные стили →
    результат идентичен ForceCopyStyles=false (оба — copy with suffix).
    Сравнить что DiffStyle_0 создан в обоих случаях.

TestRWC_KeepDifferent_ChainBothDifferent
    Два стиля: Child basedOn Parent, оба есть в target с разным
    форматированием → оба скопированы с суффиксами,
    basedOn в Child_0 ремаплен на Parent_0.

TestRWC_KeepDifferent_ChainParentInTarget
    Стиль разный, basedOn указывает на стиль target (одинаковый) →
    стиль скопирован с суффиксом, basedOn указывает на оригинальный
    target style. docDefaults compensateUncovered применён.

TestRWC_KeepDifferent_RootStyleCompensateAll
    Корневой стиль (без basedOn), разный, src/tgt docDefaults различаются →
    копия с суффиксом + compensateAll дельта инжектирована в pPr/rPr.

TestRWC_KeepDifferent_CharacterStyle
    Разный character style (w:type="character") →
    копируется с суффиксом, rPr сохранён.

TestRWC_KeepDifferent_MultipleDifferent
    Три стиля с разным форматированием →
    суффиксы _0, _1, _2, без коллизий.

TestRWC_KeepDifferent_DocDefaultsCompensation
    Разный стиль, src/tgt docDefaults различаются (src: sz=24, tgt: sz=20),
    корневой стиль → копия с суффиксом + compensateAll дельта корректно
    инжектирована в клон. Верифицировать что rPr содержит sz delta.
    (Гарантирует что Pass 2 fixupCopiedStyles обрабатывает клоны,
    созданные через KeepDifferentStyles always-copy path.)

TestRWC_KeepDifferent_RoundTrip
    Полный round-trip: insert + Save + reopen + verify:
    скопированные стили сохранились, pStyle ссылки корректны.
```

---

## Задача 2: stylesContentEqual — рекурсивное сравнение цепочки

**Приоритет:** Высокий
**Файлы:** `resourceimport_styles.go`
**Тесты:** `resourceimport_styles_test.go`

### Проблема

Aspose `Style.Equals()` рекурсивно сравнивает `basedOn`, `next`, `link`
цепочки. go-docx `stylesContentEqual` сравнивает только XML самого элемента.

Два стиля с одинаковым собственным XML, но разными `basedOn` родителями,
будут считаться identical в go-docx, но different в Aspose — их **эффективное
форматирование** отличается из-за разных inherited properties.

### Что делать

#### Шаг 1: Извлечь walkStyleChain

Из `resolveStyleChainOpt` (resourceimport_styles.go:817–886) извлечь
chain walk + merge в standalone package-level функцию:

```go
// walkStyleChain walks the basedOn chain starting from style, resolving
// properties through the given styles part. Returns merged pPr and rPr
// (derived overrides base). Does NOT include docDefaults — that's the
// caller's responsibility. Does NOT strip pStyle/rStyle — caller decides.
//
// Cycle-safe via visited set.
func walkStyleChain(
    style *oxml.CT_Style,
    styles *oxml.CT_Styles,
) (pPr, rPr *etree.Element)
```

Извлекаемый код: строки 823–873 (build chain + merge loop).
**Не** включать: docDefaults prepend (строки 848–855) и strip
pStyle/rStyle (строки 878–883) — они остаются в `resolveStyleChainOpt`.

Обновить `resolveStyleChainOpt`:

```go
func (ri *ResourceImporter) resolveStyleChainOpt(
    style *oxml.CT_Style, includeDocDefaults bool,
) (pPr, rPr *etree.Element) {
    srcStyles, err := ri.sourceStyles()
    if err != nil {
        return nil, nil
    }

    pPr, rPr = walkStyleChain(style, srcStyles)

    // Optionally prepend source docDefaults as base (Fix A).
    if includeDocDefaults {
        if ri.srcDocDefaultsPPr != nil {
            base := ri.srcDocDefaultsPPr.Copy()
            if pPr != nil {
                overridePropertiesDeep(base, pPr)
            }
            pPr = base
        }
        if ri.srcDocDefaultsRPr != nil {
            base := ri.srcDocDefaultsRPr.Copy()
            if rPr != nil {
                overridePropertiesDeep(base, rPr)
            }
            rPr = base
        }
    }

    // Strip style-internal references (not meaningful as direct attrs).
    if pPr != nil {
        removeChild(pPr, "w", "pStyle")
    }
    if rPr != nil {
        removeChild(rPr, "w", "rStyle")
    }
    return
}
```

#### Шаг 2: stylesEffectiveEqual

```go
// stylesEffectiveEqual compares two styles by their resolved effective
// formatting, walking the full basedOn chain in each document's styles.
//
// Mirrors Aspose.Words Style.Equals():
//   - Recursively compares basedOn/linked chains
//   - Excludes docDefaults (per Aspose: "Styles defaults are not
//     included in comparison")
//   - Ignores w:name and rsid* attributes
//
// Method on ResourceImporter because it needs access to both
// sourceStyles() and targetStyles() for chain resolution.
func (ri *ResourceImporter) stylesEffectiveEqual(
    srcStyle *oxml.CT_Style,
    tgtStyle *oxml.CT_Style,
) bool
```

Алгоритм:
1. `srcPPr, srcRPr := walkStyleChain(srcStyle, srcStyles)`
2. `tgtPPr, tgtRPr := walkStyleChain(tgtStyle, tgtStyles)`
3. Strip pStyle/rStyle из обоих результатов
4. Strip rsid* из обоих результатов
5. Сериализовать в canonical XML, сравнить побайтово
6. Оба nil → equal; один nil другой нет → not equal

#### Шаг 3: Обновить вызов в mergeOneStyle

Заменить строку 262:
```go
// Было:
if stylesContentEqual(srcStyle, existing) {
// Стало:
if ri.stylesEffectiveEqual(srcStyle, existing) {
```

`stylesContentEqual` **оставить** как package-internal utility —
она используется в существующих тестах, удаление создаст
ненужный churn. Пометить комментарием `// Deprecated: use
stylesEffectiveEqual for import decisions.`

### Граничные случаи

1. **Оба стиля корневые, одинаковый pPr/rPr** → equal
2. **Оба корневые, разный rPr** → not equal
3. **Одинаковый свой XML, разные basedOn → разный resolved** → not equal
4. **Разный свой XML, одинаковый resolved** → equal
   (child override + разные parents = одинаковый effective)
5. **Цикл в basedOn (A→B→A)** → cycle protection в walkStyleChain
6. **basedOn ссылается на несуществующий стиль** → chain обрывается,
   сравнение по доступной части
7. **Один корневой, другой с basedOn, одинаковый resolved** → equal
8. **Table style** → pPr/rPr resolved одинаково
9. **Оба nil pPr, оба nil rPr** → equal (пустые стили)

### Тесты

Расположение: `resourceimport_styles_test.go`

Каждый тест создаёт source и target `CT_Styles` через `buildStylesXml`,
конструирует `ResourceImporter` с ними, вызывает `stylesEffectiveEqual`.

```
TestStylesEffectiveEqual_IdenticalRoot
    Два корневых стиля с одинаковым pPr(jc=center) + rPr(b) → true.

TestStylesEffectiveEqual_DifferentRoot
    Два корневых стиля: src rPr(b), tgt rPr(i) → false.

TestStylesEffectiveEqual_SameXmlDifferentParent
    Стиль X: pPr(jc=center), basedOn=P.
    Source P: rPr(sz=24). Target P: rPr(sz=20).
    Resolved X отличается по sz → false.

TestStylesEffectiveEqual_DifferentXmlSameResolved
    Source X: basedOn=P1(jc=left), own pPr(jc=center) → effective jc=center.
    Target X: basedOn=P2(jc=center), no own pPr → effective jc=center.
    Одинаковый effective → true.

TestStylesEffectiveEqual_CircularProtection
    A basedOn B, B basedOn A → не зависает, возвращает результат.

TestStylesEffectiveEqual_MissingBasedOn
    basedOn ссылается на несуществующий стиль → chain обрывается,
    сравнение по доступной части цепочки.

TestStylesEffectiveEqual_IgnoresName
    Разные w:name ("Heading 1" vs "Заголовок 1"), одинаковый
    formatting → true.

TestStylesEffectiveEqual_IgnoresRsid
    Разные rsid* атрибуты, одинаковое форматирование → true.

TestStylesEffectiveEqual_ExcludesDocDefaults
    Source docDefaults: rPr(sz=24). Target docDefaults: rPr(sz=20).
    Стили одинаковы по resolved chain (без docDefaults) → true.

TestWalkStyleChain_Basic
    Стиль с basedOn → returns merged pPr/rPr.

TestWalkStyleChain_NoBasedOn
    Корневой стиль → returns own pPr/rPr.

TestWalkStyleChain_CycleProtection
    A→B→A → не зависает.
```

Integration с Задачей 1 (replacecontent_test.go):

```
TestRWC_KeepDifferent_EffectiveEqual_SameResolvedDiffXml
    Source и target стили с разным собственным XML, но одинаковым
    effective formatting (source: basedOn=P1(jc=left), own jc=center;
    target: basedOn=P2(jc=center), no own jc) → effective jc=center
    в обоих → стиль НЕ копируется, используется target definition.
    Верифицирует combined behavior Задач 1+2.
```

---

## Задача 3: KeepSourceNumbering=false — multi-level comparison

**Приоритет:** Средний
**Файлы:** `resourceimport_num.go`
**Тесты:** `resourceimport_num_test.go`, `replacecontent_test.go`

### Проблема

go-docx матчит списки по `numFmt` **только первого уровня**. Aspose
сравнивает **list definition identifiers** — более широкий набор
характеристик. Текущая эвристика может ложно объединить разные списки
(одинаковый numFmt уровня 0, но разная lvlText или разные sub-levels).

### Что делать

#### Шаг 1: abstractNumsCompatible

Добавить функцию (resourceimport_num.go):

```go
// abstractNumsCompatible reports whether two abstractNum definitions
// are semantically compatible for list merging. Two definitions are
// compatible when every level present in BOTH has identical numFmt
// and lvlText values.
//
// Levels present in only one definition are ignored — Word handles
// missing levels gracefully by falling back to the definition that
// has them.
//
// Returns false if either has no levels or if no levels overlap.
//
// Note: w:start/@w:val (start number) is intentionally excluded from
// comparison. Aspose considers lists with same numFmt/lvlText but
// different start values compatible for merging — Word adjusts
// continuation numbering automatically.
func abstractNumsCompatible(src, tgt *etree.Element) bool
```

Алгоритм:
1. Извлечь `map[ilvl]{numFmt, lvlText}` из обоих
2. Найти **пересечение** ilvl ключей
3. Если пересечение пустое → false
4. Для каждого ilvl в пересечении: если numFmt или lvlText разные → false
5. Все совпали → true

Вспомогательная функция:

```go
// levelSignature holds the identity-defining properties of a single
// numbering level for merge compatibility checks.
type levelSignature struct {
    numFmt  string // w:numFmt/@w:val (e.g. "decimal", "bullet")
    lvlText string // w:lvlText/@w:val (e.g. "%1.", "•")
}

// extractLevelSignatures returns a map from ilvl string ("0", "1", ...)
// to levelSignature for each <w:lvl> in an abstractNum element.
func extractLevelSignatures(absNum *etree.Element) map[string]levelSignature
```

#### Шаг 2: Обновить findMatchingTargetNum

В `findMatchingTargetNum` (resourceimport_num.go:287–322) заменить:

```go
// Было (строка 302–309):
srcFmt := firstLevelNumFmt(srcAbsNum)
if srcFmt == "" {
    return 0
}
for _, tgtAbsNum := range tgtNumbering.AllAbstractNums() {
    if firstLevelNumFmt(tgtAbsNum) != srcFmt {
        continue
    }

// Стало:
if !hasLevels(srcAbsNum) {
    return 0
}
for _, tgtAbsNum := range tgtNumbering.AllAbstractNums() {
    if !abstractNumsCompatible(srcAbsNum, tgtAbsNum) {
        continue
    }
```

`hasLevels` — простая проверка наличия хотя бы одного `<w:lvl>`.

#### Шаг 3: Удалить firstLevelNumFmt

После замены `firstLevelNumFmt` больше нигде не используется.
Удалить функцию (строки 324–344) и её тесты
(`TestFirstLevelNumFmt_*` в resourceimport_num_test.go:257–318).

### Граничные случаи

1. **Оба: 1 уровень, decimal, "%1."** → compatible (merge)
2. **Оба: 1 уровень, bullet, "•"** → compatible
3. **Оба: decimal, но разный lvlText ("%1." vs "%1)")** → not compatible
4. **Оба: 3 уровня, все совпадают** → compatible
5. **Source: 3 уровня, target: 1 уровень, level 0 совпадает** →
   compatible (пересечение = {0}, совпадает)
6. **Разный numFmt на уровне 2, совпадает на уровне 0** → not compatible
7. **Source: bullet, target: decimal** → not compatible
8. **Оба пустые (нет lvl)** → not compatible
9. **Один пустой, другой с уровнями** → not compatible (пересечение пустое)
10. **Level 0 missing в source, level 1 совпадает** →
    compatible (пересечение = {1}, совпадает)

### Тесты

Расположение: `resourceimport_num_test.go`

```
TestAbstractNumsCompatible_SingleLevel_Same
    Оба: lvl 0, decimal, "%1." → true.

TestAbstractNumsCompatible_SingleLevel_DifferentFmt
    lvl 0: decimal vs bullet → false.

TestAbstractNumsCompatible_SingleLevel_DifferentText
    lvl 0: decimal "%1." vs decimal "%1)" → false.

TestAbstractNumsCompatible_MultiLevel_AllMatch
    3 уровня, все numFmt + lvlText совпадают → true.

TestAbstractNumsCompatible_MultiLevel_Level2Differs
    lvl 0,1 совпадают, lvl 2 различается → false.

TestAbstractNumsCompatible_DifferentLevelCount_OverlapOk
    Source: lvl 0,1,2. Target: lvl 0. Level 0 совпадает → true.

TestAbstractNumsCompatible_NoLevels
    Оба без <w:lvl> → false.

TestAbstractNumsCompatible_OneEmpty
    Source без lvl, target с lvl → false (пересечение пустое).

TestAbstractNumsCompatible_NoOverlap
    Source: только lvl 0. Target: только lvl 1 → false.

TestExtractLevelSignatures_Basic
    abstractNum с 3 уровнями → корректная map.

TestExtractLevelSignatures_Empty
    abstractNum без lvl → пустая map.

TestExtractLevelSignatures_MissingNumFmt
    lvl без numFmt → numFmt="" в signature.
```

Integration (replacecontent_test.go):

```
TestRWC_NumMerge_MultiLevel_Compatible
    Source и target с 3-уровневым decimal списком, все совпадают →
    absNums не увеличились (merged).

TestRWC_NumMerge_MultiLevel_Incompatible
    Source: 3 уровня (decimal, lowerLetter, lowerRoman).
    Target: 3 уровня (decimal, upperLetter, upperRoman).
    Level 0 совпадает, levels 1,2 различаются →
    absNums увеличились (separate).

TestRWC_NumMerge_DifferentLvlText_Separate
    Оба decimal + lvl 0, но "%1." vs "%1)" →
    absNums увеличились (not merged despite same numFmt).
```

---

## Задача 4: ImportFormatOptions.IgnoreHeaderFooter

**Приоритет:** Средний
**Файлы:** `contentdata.go`, `document.go`, `resourceimport_styles.go`
**Тесты:** `replacecontent_test.go`

### Проблема

В Aspose `IgnoreHeaderFooter` (default: `true`) при `KeepSourceFormatting`
не применяет source formatting к headers/footers — они используют
стили целевого документа. go-docx обрабатывает headers/footers
идентично body.

### Архитектурное решение

**Ключевое ограничение:** нельзя создавать второй `ResourceImporter` —
это приведёт к дублированию стилей и numbering в target (оба RI
запишут копии одних стилей, каждый со своим idempotent state).

**Правильный подход:** один RI, но `prepareContentElements` для hdr/ftr
пропускает `expandDirectFormatting` и использует UseDestinationStyles
маппинг при remap. Это достигается флагом в `prepareContentElements`.

### Что делать

#### Шаг 1: Добавить поле в ImportFormatOptions

В `contentdata.go`:

```go
type ImportFormatOptions struct {
    ForceCopyStyles     bool
    KeepSourceNumbering bool

    // IgnoreHeaderFooter controls whether source formatting is applied
    // to header/footer content during import.
    //
    // When true, headers and footers always use destination styles
    // regardless of ImportFormatMode. Conflicting styles are NOT
    // expanded to direct attributes and NOT copied with suffix —
    // the target style definition is used as-is.
    //
    // When false (default), headers/footers are processed identically
    // to the document body.
    //
    // Aspose.Words default: true. Go zero-value: false (backward
    // compatible — existing behavior unchanged).
    //
    // Mirrors Aspose.Words ImportFormatOptions.IgnoreHeaderFooter.
    IgnoreHeaderFooter bool
}
```

#### Шаг 2: prepareContentElements — добавить параметр useDestStyles

```go
func prepareContentElements(
    sourceDoc *Document,
    targetPart *parts.StoryPart,
    ri *ResourceImporter,
    useDestStyles bool, // true → skip expand, remap as UseDestinationStyles
) (*preparedContent, error)
```

Внутри, при `useDestStyles=true`:
- **Пропустить** `ri.expandDirectFormatting(elements)` (Step 2d)
- В `ri.remapAll(elements)` (Step 2e) стили, помеченные в
  `ri.expandStyles`, будут ремаплены на `targetDefaultParaStyleId()`
  — это уже правильное поведение, потому что при UseDestinationStyles
  конфликтные стили маппятся на target ID, а `expandStyles` пуст.

Но есть нюанс: `ri.styleMap` уже заполнен основным mode. Для
UseDestinationStyles конфликтные стили должны маппиться на target
styleId. Значит нужно **отдельный styleMap для hdr/ftr**.

**Упрощение:** вместо двух styleMap можно на этапе remap для hdr/ftr
подменять mapping: если стиль в `expandStyles` (помечен для expand),
remap его на target styleId вместо default. Для этого:

```go
func prepareContentElements(
    sourceDoc *Document,
    targetPart *parts.StoryPart,
    ri *ResourceImporter,
    ignoreSourceFormatting bool,
) (*preparedContent, error) {
    // ... steps 1-2b-2c unchanged ...

    if !ignoreSourceFormatting {
        // Normal path: expand direct formatting for KeepSourceFormatting.
        ri.expandDirectFormatting(elements)
    }

    // Remap references.
    if ignoreSourceFormatting {
        ri.remapAllUseDestStyles(elements)
    } else {
        ri.remapAll(elements)
    }

    // ... steps 3-5 unchanged ...
}
```

#### Шаг 3: remapAllUseDestStyles

Новый метод на `ResourceImporter` (resourceimport.go):

```go
// remapAllUseDestStyles remaps resource references using
// UseDestinationStyles strategy, regardless of the RI's actual mode.
//
// For conflicting styles (both expanded and copied), leaves the
// original source styleId unchanged — the target document already
// has a style definition under that same ID, which is exactly
// the "use destination styles" semantics.
//
// For non-conflicting styles (exist only in source, or built-in name
// fallback where srcId ≠ targetId), uses the standard styleMap mapping.
//
// Used when IgnoreHeaderFooter=true for header/footer content.
func (ri *ResourceImporter) remapAllUseDestStyles(elements []*etree.Element)
```

Идентичен `remapAll`, но для pStyle/rStyle/tblStyle:
1. Если стиль в `ri.expandStyles` → **не ремапить** (оставить как есть,
   target имеет стиль под тем же ID — UseDestinationStyles)
2. Если стиль в `ri.copiedStyleIds` → **не ремапить** (оставить как есть,
   target имеет оригинальный стиль под тем же ID; копию с суффиксом
   игнорируем для hdr/ftr)
3. Если стиль в `ri.styleMap` И НЕ в expandStyles/copiedStyleIds →
   использовать mapping (для стилей только из source, или built-in
   name fallback: RU `"a"` → EN `"Normal"`)
4. Иначе → оставить как есть

**Обоснование условий 1–2:** конфликт стилей означает, что source
styleId == target styleId (именно совпадение ID создаёт конфликт).
Оставляя ID без ремапа, параграф ссылается на target definition.

**Исключение для built-in name fallback:** при cross-locale match
(например, source `"a"` → target `"Normal"` для одного built-in стиля)
styleMap содержит `"a" → "Normal"`, но `"a"` НЕ в copiedStyleIds
(стиль не копировался, а маппился на target). Условие 3 корректно
использует mapping.

#### Шаг 4: Обновить вызовы prepareContentElements

В `document.go`:
```go
// Body — всегда полный pipeline.
bodyPrep, err := prepareContentElements(cd.Source, &d.part.StoryPart, ri, false)
```

В `section.go`, `replaceWithContentDedup`:
```go
// Header/footer — conditional.
ignoreHF := ri.opts.IgnoreHeaderFooter &&
    ri.importFormatMode != UseDestinationStyles
prep, err := prepareContentElements(sourceDoc, bic.part, ri, ignoreHF)
```

Это требует, чтобы `replaceWithContentDedup` знал об `opts` — но
он уже получает `ri`, у которого есть `ri.opts` и `ri.importFormatMode`.

В `document.go`, comments — всегда полный pipeline:
```go
// Comments — NOT affected by IgnoreHeaderFooter (matches Aspose behavior).
prep, err := prepareContentElements(sourceDoc, c.BlockItemContainer.part, ri, false)
```

#### Существующие call sites, которые сломаются

Добавление 4-го параметра `ignoreSourceFormatting bool` в
`prepareContentElements` требует обновления всех вызовов:

| Файл | Строка | Контекст | Значение |
|------|--------|----------|----------|
| `document.go` | 657 | body | `false` |
| `document.go` | 726 | comments | `false` |
| `section.go` | 496 | hdr/ftr | conditional |
| `contentdata_test.go` | 810 | test | `false` |
| `contentdata_test.go` | 988 | test | `false` |
| `contentdata_test.go` | 1114 | test | `false` |
| `contentdata_test.go` | 1171 | test | `false` |
| `contentdata_test.go` | 1196 | test | `false` |
| `contentdata_test.go` | 1276 | test | `false` |
| `contentdata_test.go` | 1356 | test | `false` |

### Граничные случаи

1. **IgnoreHeaderFooter=false (default)** → без изменений
2. **IgnoreHeaderFooter=true + UseDestinationStyles** → no-op (уже dest)
3. **IgnoreHeaderFooter=true + KeepSourceFormatting** → hdr/ftr: no expand,
   конфликтные стили остаются на target styleId
4. **IgnoreHeaderFooter=true + KeepDifferentStyles** → hdr/ftr: no copy
   with suffix, конфликтные стили на target styleId
5. **Body не затронут** — body всегда через полный pipeline
6. **Нет headers/footers** → no-op

### Тесты

Расположение: `replacecontent_test.go`

```
TestRWC_IgnoreHeaderFooter_Default_ProcessesSameAsBody
    IgnoreHeaderFooter=false + KeepSourceFormatting → header content
    получает expanded formatting (jc/bold в direct attrs).

TestRWC_IgnoreHeaderFooter_True_UsesDestStyles
    IgnoreHeaderFooter=true + KeepSourceFormatting → header content
    НЕ получает expanded formatting, стиль не переименован,
    pStyle ссылается на target styleId.

TestRWC_IgnoreHeaderFooter_True_UseDestNoOp
    IgnoreHeaderFooter=true + UseDestinationStyles →
    поведение идентично IgnoreHeaderFooter=false.

TestRWC_IgnoreHeaderFooter_BodyUnaffected
    IgnoreHeaderFooter=true + KeepSourceFormatting →
    body получает expand (full pipeline),
    header НЕ получает expand (dest styles).
    Одинаковый контент, разное поведение.

TestRWC_IgnoreHeaderFooter_True_KeepDifferent
    IgnoreHeaderFooter=true + KeepDifferentStyles →
    header: разный стиль НЕ копируется с суффиксом,
    используется target definition.

TestRWC_IgnoreHeaderFooter_True_ForceCopy
    IgnoreHeaderFooter=true + KeepSourceFormatting + ForceCopyStyles →
    body: стиль скопирован как Style_0.
    header: стиль НЕ ремаплен на Style_0, ссылается на оригинальный
    target styleId (copiedStyleIds check).

TestRWC_IgnoreHeaderFooter_CommentsUnaffected
    IgnoreHeaderFooter=true + KeepSourceFormatting →
    comments получают expanded formatting (полный pipeline),
    аналогично body.
```

---

## Задача 5: ImportFormatOptions.MergePastedLists

**Приоритет:** Низкий
**Файлы:** `contentdata.go`, `document.go`, `blkcntnr.go`
**Тесты:** `replacecontent_test.go`

### Проблема

Aspose.Words `MergePastedLists` (default: `false`) позволяет вставленным
спискам сливаться с окружающими списками в целевом документе.
В go-docx этой опции нет.

### Архитектурное решение

**Ключевое ограничение:** `BlockItemContainer.replaceWithContent` не знает
о `ResourceImporter` — получает только `*preparedContent`. Merge списков
после вставки требует доступа к numbering-структурам target документа.

**Правильный подход:** реализовать merge на уровне `Document.ReplaceWithContent`
(document.go), **после** всех replacements. Post-processing обходит body,
находит границы вставки (через marker или по факту — соседние list
paragraphs с разными numId) и перезаписывает numId.

Альтернативный (более чистый) подход: merge выполняется в
`preparedContent` callback — `replaceWithContent` в blkcntnr.go
уже возвращает вставленные элементы через `spliceElement`. Расширить
callback, чтобы после splice проверял соседей.

### Что делать

#### Шаг 1: Добавить поле в ImportFormatOptions

```go
    // MergePastedLists merges inserted list paragraphs with adjacent
    // list paragraphs in the target document when they use compatible
    // numbering (same abstractNum definition or compatible structure).
    //
    // When true, the first/last inserted list paragraph adopts the
    // numId of the adjacent target list paragraph if compatible.
    // When false (default), inserted lists remain independent.
    //
    // Mirrors Aspose.Words ImportFormatOptions.MergePastedLists.
    MergePastedLists bool
```

#### Шаг 2: Post-processing в Document.ReplaceWithContent

**Scope:** body, headers/footers, comments — все контейнеры.
Aspose применяет MergePastedLists ко всем content containers.

После body replacement (document.go, после строки 674):

```go
if cd.Options.MergePastedLists {
    mergeAdjacentLists(b.Element(), ri.targetDoc)
}
```

После hdr/ftr replacement (в цикле по hfs, после `hf.replaceWithContentDedup`):
```go
if cd.Options.MergePastedLists {
    // hf.element is the <w:hdr>/<w:ftr> container.
    mergeAdjacentLists(hf.element(), ri.targetDoc)
}
```

После comments replacement — аналогично per comment body.

Функция `mergeAdjacentLists` (новый файл или в contentdata.go):

```go
// mergeAdjacentLists scans the container for adjacent list paragraphs
// with different numId values and merges them when they reference
// compatible abstractNum definitions. "Compatible" = same abstractNumId
// in target numbering, meaning they were originally the same list type.
//
// This handles the case where inserted content starts/ends with a list
// paragraph adjacent to a target list paragraph — the inserted numId is
// replaced with the target's numId so Word treats them as one list.
func mergeAdjacentLists(container *etree.Element, doc *Document)
```

Алгоритм:
1. Walk children of container, tracking previous list paragraph
2. When current is list paragraph and prev is list paragraph:
   - If different numId but same abstractNumId in target numbering → merge
     (set current's numId to prev's numId)
3. Recurse into tables (cells are containers too)

Проверка «same abstractNumId» гарантирует, что merge происходит
только для списков одного типа (а не произвольных decimal+bullet).

**Важно:** comparison uses **target** numbering part. После
`importNumbering()` все numId в inserted elements уже ремаплены
на target-side значения. Оба соседних параграфа (target + inserted)
ссылаются на abstractNumId из target `numbering.xml`.

#### Шаг 3: Вынести extractNumId helper

```go
// extractNumId returns the w:numId value from a paragraph's pPr/numPr,
// or 0 if not a list paragraph.
func extractNumId(p *etree.Element) int
```

### Граничные случаи

1. **MergePastedLists=false (default)** → no-op
2. **Нет соседних списков** → no-op
3. **Соседний список + совместимый (same abstractNumId)** → merge
4. **Соседний список + несовместимый (different abstractNumId)** → no merge
5. **Вставка между двумя списками** → merge обоих краёв
6. **Вложенная таблица с списком** → recurse handles
7. **Несколько placeholder в одном body** → каждая граница проверяется

### Тесты

Расположение: `replacecontent_test.go`

```
TestRWC_MergePastedLists_Default_Separate
    MergePastedLists=false, вставка списка рядом со списком →
    два отдельных списка (разные numId).

TestRWC_MergePastedLists_True_MergedWithPrev
    MergePastedLists=true, target имеет decimal list, вставка
    начинается с decimal list → inserted numId = prev numId.

TestRWC_MergePastedLists_True_IncompatibleNoMerge
    MergePastedLists=true, target: bullet, source: decimal →
    разные abstractNumId → списки остаются раздельными.

TestRWC_MergePastedLists_True_NoAdjacentList
    MergePastedLists=true, вставка рядом с обычным параграфом →
    no-op.

TestRWC_MergePastedLists_True_BothEdges
    MergePastedLists=true, вставка между двумя совместимыми
    списками → оба края merged.

TestRWC_MergePastedLists_True_InTable
    MergePastedLists=true, placeholder внутри таблицы,
    вставка рядом со списком в ячейке → merge в ячейке.
```

---

## Задача 6: Аннотации — bookmarks import (опционально)

**Приоритет:** Низкий
**Файлы:** `contentdata.go`
**Тесты:** `contentdata_test.go`

### Проблема

Aspose.Words **импортирует** bookmarks. go-docx **удаляет** все
аннотации, включая bookmarks. Bookmarks полезны для cross-references
и named anchors.

### Что делать

#### Шаг 1: Убрать bookmarks из annotationMarkers

В `contentdata.go`, map `annotationMarkers` (строка 615):
удалить `"bookmarkStart"` и `"bookmarkEnd"`.

#### Шаг 2: renumberBookmarkIDs

Добавить в pipeline `prepareContentElements`, после `sanitizeForInsertion`
и перед `materializeImplicitStyles`:

```go
// Step 2b2: Renumber bookmark IDs and deduplicate names.
renumberBookmarks(elements, targetPart)
```

Функция:

```go
// renumberBookmarks rewrites w:id on bookmarkStart/bookmarkEnd pairs
// with fresh values, and appends a numeric suffix to w:name on
// bookmarkStart to avoid name collisions in the target document.
//
// Bookmark w:id must be unique per document (paired start/end share
// the same id). Bookmark w:name must also be unique — Word silently
// discards duplicates.
func renumberBookmarks(elements []*etree.Element, docPart *parts.DocumentPart)
```

Алгоритм:
1. DFS pass 1: collect all bookmarkStart `w:id` values → build
   `oldId → newId` map using `docPart.NextBookmarkID()` counter.
   Simultaneously append suffix to `w:name` (e.g. `_imp1`, `_imp2`)
   to guarantee uniqueness.
2. DFS pass 2: rewrite `w:id` on bookmarkStart and bookmarkEnd using map.

Или в один проход: rewrite immediately, matching start/end by shared
old `w:id`.

#### Шаг 3: NextBookmarkID на DocumentPart (не StoryPart!)

Bookmark w:id уникален **per document** (не per StoryPart).
Счётчик должен быть на `DocumentPart`, не на `StoryPart` —
иначе body и headers могут сгенерировать одинаковые id.

Реализация: scan ALL story parts (body + headers + footers +
comments) для max existing bookmark id on first call, cache
on `DocumentPart`, increment atomically.

```go
// NextBookmarkID returns a fresh document-unique bookmark ID.
// Thread-safe within a single Document (matches Document concurrency model).
func (dp *DocumentPart) NextBookmarkID() int
```

### Замечание

Эта задача опциональна. Удаление аннотаций — валидная стратегия
для content insertion. Реализовывать только при явном запросе.

### Граничные случаи

1. **Source без bookmarks** → no-op
2. **Один bookmark (start + end)** → id перенумерован, name уникален
3. **Несколько bookmarks** → все id уникальны, все name уникальны
4. **Bookmark name коллизия с target** → suffix предотвращает
5. **Вложенные bookmarks** (start1, start2, end2, end1) → обе пары
   корректно перенумерованы
6. **Bookmark spanning multiple paragraphs** → start и end получают
   одинаковый новый id
7. **Множественные вставки одного source** → каждая вставка получает
   свои уникальные id и name

### Тесты

Расположение: `contentdata_test.go`

```
TestSanitize_BookmarksPreserved
    После sanitizeForInsertion bookmarkStart/End не удаляются
    (в отличие от commentRangeStart, которые удаляются).

TestRenumberBookmarks_SinglePair
    Один bookmark: start(id=5, name="bm1") + end(id=5) →
    перенумерованы в одинаковый новый id, name получает суффикс.

TestRenumberBookmarks_MultiplePairs
    Три bookmark пары → все id уникальны.

TestRenumberBookmarks_NestedPairs
    Вложенные bookmarks → корректная перенумерация обеих пар.

TestRenumberBookmarks_NameDedup
    Два вызова renumberBookmarks → name суффиксы не коллидируют.

TestRWC_BookmarksImported_RoundTrip
    Integration: source с bookmarks → Save + reopen → bookmarks
    присутствуют в target с уникальными id и name.
```

---

## Порядок реализации

```
Задача 1 ──→ Задача 2 (weak dependency: test expectations)
             ↓
             (параллельно) → Задача 3
                           → Задача 4
                           → Задача 5
                           → Задача 6
```

**Задача 1** → **Задача 2**: слабая зависимость (weak dependency).
Задача 1 исправляет основное поведение (copy вместо expand).
Задача 2 улучшает точность сравнения (effective vs raw XML).
Извлечение `walkStyleChain` из `resolveStyleChainOpt` — чистый
рефакторинг, технически может выполняться параллельно с Задачей 1.
Зависимость только на уровне тестов: после Задачи 1 тест
`TestRWC_KeepDifferent_DifferentExpanded` переписан, а Задача 2
меняет сравнение в `mergeOneStyle`. Рекомендация: последовательно
для минимизации конфликтов при merge.

**Задачи 3–6**: полностью независимы друг от друга и от Задач 1–2.
Могут выполняться параллельно.

---

## Обновление visual-regtest

После Задач 1–2:
- `"__keep_diff"` — выходные файлы изменятся (expand → copy with suffix).
  Перегенерировать. Визуальное качество должно быть **идентичным или лучше**
  (стиль-copy точнее передаёт форматирование, чем expand).
- `"__keep_diff_force"` — станет идентичен `"__keep_diff"`. Оставить
  для regression guard.

После Задачи 4:
- Добавить вариант `"__ignore_hf"` с `IgnoreHeaderFooter: true`.

---

## Обновление README_ReplaceWithContent.md

1. **KeepDifferentStyles**: описать always-copy, убрать expand.
2. **ForceCopyStyles**: уточнить «работает только с KeepSourceFormatting.
   Для KeepDifferentStyles не требуется — копирование при расхождении
   выполняется автоматически».
3. **ImportFormatOptions**: добавить IgnoreHeaderFooter, MergePastedLists.
4. **Ограничения**: обновить KeepSourceNumbering заметку.

---

## Вне scope этого плана

Следующие опции Aspose.Words **не реализуются** в рамках данного плана.
Перечислены для полноты и предотвращения scope creep:

- **`SmartStyleBehavior`** — сложная эвристика выбора стилей при
  конфликте. Требует отдельного исследования поведения Aspose.
- **`AdjustSentenceAndWordSpacing`** — микротипографическая подстройка
  интервалов. Низкий приоритет, косметический эффект.
- **Table style conditional formatting** — Aspose имеет расширенную
  обработку `tblStylePr` (first row, last column, banding). Текущая
  реализация копирует table styles as-is, что достаточно для большинства
  сценариев.

Эти опции могут быть добавлены отдельными задачами при явном запросе.

---

## Чек-лист готовности к реализации

Для каждой задачи перед началом работы убедиться:

- [ ] Точные строки кода для изменения верифицированы (файл:строка)
- [ ] Список существующих тестов, которые сломаются, составлен
- [ ] Новые тесты не конфликтуют по имени с существующими
- [ ] Зависимости от других задач отсутствуют или уже выполнены
- [ ] `go test ./pkg/docx/ -count=1` проходит до начала работы
