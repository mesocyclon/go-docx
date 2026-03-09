# Fix Plan: docDefaults Reconciliation & Built-in Style Resolution

**Версия:** 2.0
**Статус:** Ready for implementation
**Область:** `pkg/docx/resourceimport_styles.go`, `pkg/docx/oxml/styles_custom.go`, `pkg/docx/resourceimport.go`
**Связь:** Дополнение к ENTERPRISE_PLAN_FINAL_v2.md

---

## 1. Диагноз: три связанных дефекта

### 1.1 Дефект A — docDefaults не компенсируются при копировании стилей (CRITICAL)

**Суть:** Когда стиль копируется из source в target через `copyStyleToTarget`, он переносится as-is. Но стиль проектировался для source docDefaults. В target другие docDefaults — и стиль, сидя на них, даёт другой визуальный результат.

**Конкретный пример из теста:**

```
Source docDefaults (mark/1_page.docx):
  rPrDefault: rFonts="Times New Roman", НЕТ sz → OOXML implicit 10pt (val="20")
  pPrDefault: ПУСТОЙ → OOXML implicit: after=0, line=240 (single spacing)

Target docDefaults (template.docx):
  rPrDefault: rFonts=theme:minorHAnsi, sz=22 (11pt), szCs=22
  pPrDefault: spacing after=160, line=259 (1.15 spacing)
```

Скопированный стиль `Normal` из source:

```xml
<w:style w:styleId="Normal">
  <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
  <!-- нет sz, нет spacing в pPr → наследует из docDefaults -->
</w:style>
```

В source: Arial 10pt, single spacing.
В target: Arial **11pt**, spacing **after=160 line=259** — форматирование сломано.

**Затронуто:** ВСЕ скопированные стили, у которых хотя бы одно свойство наследуется из docDefaults (а не задано явно). Это практически все реальные документы.

**Два пути воздействия:**
- **Copy path** (стиль копируется в target): стиль наследует чужие docDefaults → нужна инъекция delta в clone.
- **Expand path** (стиль разворачивается в direct attributes): `resolveStyleChain` не включает source docDefaults как базу → resolved properties неполные.

### 1.2 Дефект B — Built-in стили не сопоставляются по w:name (HIGH)

**Суть:** `mergeOneStyle` ищет конфликт через `GetByID(styleId)`. Но built-in стили в локализованных документах имеют нестандартные styleId:

```
Target (RU): styleId="a"     name="Normal"     (default paragraph)
Source (EN): styleId="Normal" name="Normal"     (default paragraph)
```

`GetByID("Normal")` возвращает nil в target → стиль считается "отсутствующим" → копируется целиком. Результат: **два стиля с одним именем** и два `w:default="1"` одного типа.

**Реальный output содержит:**
- `styleId="a"` type=paragraph default=1 name="Normal"
- `styleId="Normal"` type=paragraph default=1 name="Normal"
- `styleId="a1"` type=table default=1 name="Normal Table"
- `styleId="TableNormal"` type=table default=1 name="Normal Table"

### 1.3 Дефект C — Дубликат w:default="1" (MEDIUM)

**Суть:** Следствие дефекта B. Два стиля одного типа с `w:default="1"` — невалидный OOXML. Word выбирает один из них непредсказуемо, что приводит к каскадному искажению всех стилей, наследующих от default.

---

## 2. Архитектурное решение

### 2.1 Принципы

1. **Один docDefaultsDelta на ResourceImporter** — вычисляется один раз при создании, применяется ко всем скопированным стилям. Корректно работает при вставке нескольких source-документов в один target, т.к. каждый source-документ создаёт свой ResourceImporter.

2. **Name-based fallback для built-in стилей** — `mergeOneStyle` сначала ищет по styleId, при промахе ищет по w:name + w:type. Если найден — это тот же built-in стиль, просто с другим id (локализация). Устанавливается маппинг `styleMap[srcId] = tgtId`.

3. **Immutability source** — source документ никогда не модифицируется. docDefaults delta вычисляется readonly, компенсация применяется к clone.

4. **Не трогаем target docDefaults** — они принадлежат target документу. Все компенсации через инъекцию явных свойств в копии стилей.

5. **Двухпроходная обработка стилей** — merge loop (Pass 1) строит полный styleMap и собирает клоны. Fixup loop (Pass 2) выполняет delta-компенсацию и basedOn/next/link ремаппинг когда styleMap уже полный. Это решает как docDefaults-компенсацию, так и BUG 1 (basedOn remapping до готовности styleMap).

6. **Fix A + Fix B** — оба пути (expand и copy) покрываются:
   - **Fix A (expand path):** `resolveStyleChain` начинает с source docDefaults как базы цепочки.
   - **Fix B (copy path):** `compensateDocDefaults` инжектирует delta в скопированные клоны.

### 2.2 Pipeline (до и после)

**Было:**

```
importStyles()
  detectDefaultStyleMismatch()
  collectStyleIdsFromBody()
  collectStyleClosure()
  for style in closure:
    mergeOneStyle(style)         ← GetByID only, копирует+remaps as-is
```

**Станет:**

```
importStyles()
  computeDocDefaultsDelta()      ← [NEW] один раз, readonly
  detectDefaultStyleMismatch()   ← без изменений
  collectStyleIdsFromBody()
  collectStyleClosure()

  // Pass 1: merge — строит styleMap, собирает клоны
  for style in closure:
    mergeOneStyle(style)         ← [CHANGED] +name fallback, clone WITHOUT remap/compensate

  // Pass 2: fixup — когда styleMap полный
  fixupCopiedStyles()            ← [NEW] basedOn/next/link remap + delta compensate
```

---

## 3. Детальный план изменений

### Step 1 — docDefaultsDelta: вычисление и хранение

**Файл:** `pkg/docx/resourceimport_styles.go`

**Новая структура:**

```go
// docDefaultsDelta represents the difference between source and target
// docDefaults that must be injected into copied styles to preserve their
// visual appearance in the target document.
//
// OOXML formatting resolution:
//   effective = docDefaults + style chain + direct formatting
//
// When a style is designed for source docDefaults but placed atop target
// docDefaults, any property it inherits (doesn't explicitly set) will
// change value. The delta captures exactly these "inherited" properties
// so they can be materialized as explicit values in the copied style.
//
// Fields are *etree.Element (w:pPr and w:rPr) containing ONLY the
// properties that differ between source and target docDefaults.
// nil means no delta for that property group (docDefaults are identical
// or both absent).
type docDefaultsDelta struct {
    // rPr properties from source rPrDefault that either don't exist in
    // target rPrDefault or have different attribute values.
    // Example: source has no sz (implicit 10pt), target has sz=22 →
    // delta.rPr contains <w:sz w:val="20"/> (source effective value).
    rPr *etree.Element

    // pPr properties from source pPrDefault that either don't exist in
    // target pPrDefault or have different attribute values.
    // Example: source has no spacing (implicit single), target has
    // spacing after=160 line=259 →
    // delta.pPr contains <w:spacing w:after="0" w:line="240"/>
    pPr *etree.Element
}
```

**Алгоритм `computeDocDefaultsDelta`:**

```
1. Извлечь source w:docDefaults/w:rPrDefault/w:rPr  → srcRPr  (может быть nil)
2. Извлечь target w:docDefaults/w:rPrDefault/w:rPr  → tgtRPr  (может быть nil)
3. Извлечь source w:docDefaults/w:pPrDefault/w:pPr  → srcPPr  (может быть nil)
4. Извлечь target w:docDefaults/w:pPrDefault/w:pPr  → tgtPPr  (может быть nil)

5. delta.rPr = diffProperties(srcRPr, tgtRPr, ooxmlImplicitRPr)
6. delta.pPr = diffProperties(srcPPr, tgtPPr, ooxmlImplicitPPr)

Если оба nil → docDefDelta = nil (fast path, no compensation needed)
```

**Алгоритм `diffProperties(src, tgt, implicitDefaults)`:**

Возвращает `*etree.Element` с delta-свойствами или nil если различий нет.

```
result = new <w:rPr> или <w:pPr> (пустой)

// Случай 1: свойства, которые есть в source но отличаются от target
for each child-element в src (или implicitDefaults если src nil):
    srcChild = элемент из src (или implicit default)
    tgtChild = matching child в tgt (same space:tag)

    if tgtChild == nil:
        // Target не задаёт — source значение сохраняется при любых docDefaults
        // delta не нужна (цель delta — компенсировать ИЗМЕНЕНИЕ, а не отсутствие)
        continue

    if атрибуты srcChild ≠ атрибуты tgtChild:
        // Создать delta-элемент ТОЛЬКО с атрибутами, которые различаются:
        // - атрибуты из src, чьё значение ≠ значению в tgt
        // - атрибуты из src, отсутствующие в tgt
        deltaChild = new element(same space:tag)
        for each attr в srcChild:
            tgtVal = tgtChild.SelectAttrValue(attr.Key)
            if tgtVal == "" || tgtVal != attr.Value:
                deltaChild.CreateAttr(attr.Key, attr.Value)
        if len(deltaChild.Attr) > 0:
            result.AddChild(deltaChild)

// Случай 2: свойства, которые есть ТОЛЬКО в target (source их не задаёт)
// Source наследует OOXML implicit default. Если implicit ≠ target → delta нужна.
for each child-element в tgt:
    if matching child НЕТ в src:
        tag = child.Tag  // e.g. "sz", "spacing"
        implDef = implicitDefaults[tag]
        if implDef != nil:
            // Создать delta-элемент с implicit default значениями,
            // но ТОЛЬКО для атрибутов, которые отличаются от target.
            deltaChild = new element(same space:tag)
            for each attr в implDef:
                tgtVal = tgtChild.SelectAttrValue(attr.Key)
                if tgtVal != attr.Value:
                    deltaChild.CreateAttr(attr.Key, attr.Value)
            if len(deltaChild.Attr) > 0:
                result.AddChild(deltaChild)

if len(result.ChildElements()) == 0:
    return nil
return result
```

**Обработка OOXML implicit defaults:**

Когда source НЕ имеет `<w:sz>` в docDefaults, а target ИМЕЕТ `<w:sz w:val="22">` — в delta нужно записать `<w:sz w:val="20"/>` (OOXML spec default для sz = 10pt = val 20). Таблица implicit defaults:

```go
// ooxmlImplicitRPr contains OOXML-spec default values for run properties
// that are assumed when docDefaults omits them. These values are needed
// when source docDefaults omits a property but target specifies it —
// the delta must contain the spec default so it can be injected.
//
// Reference: ECMA-376 Part 1, §17.3.2 (rPr), §17.3.1 (pPr)
var ooxmlImplicitRPr = map[string]map[string]string{
    "sz":     {"w:val": "20"},   // 10pt — ECMA-376 §17.3.2.38
    "szCs":   {"w:val": "20"},   // 10pt — ECMA-376 §17.3.2.39
    // rFonts: implementation-defined per ECMA-376 §17.3.2.26.
    // De facto standard across all Word versions is Times New Roman.
    // When source omits rFonts and target specifies theme/explicit fonts,
    // we inject the de facto default to prevent silent font substitution.
    "rFonts": {"w:ascii": "Times New Roman", "w:hAnsi": "Times New Roman",
               "w:eastAsia": "Times New Roman", "w:cs": "Times New Roman"},
}

var ooxmlImplicitPPr = map[string]map[string]string{
    "spacing": {"w:after": "0", "w:before": "0",
                "w:line": "240", "w:lineRule": "auto"},
    "jc":      {"w:val": "left"},
    "ind":     {"w:left": "0", "w:right": "0"},
}
```

**Новые поля в ResourceImporter:**

```go
type ResourceImporter struct {
    // ... existing fields ...

    // docDefDelta holds the computed docDefaults difference between
    // source and target. Computed once in importStyles(), applied to
    // copied style clones in fixupCopiedStyles(). nil if docDefaults
    // are identical (no compensation needed).
    docDefDelta *docDefaultsDelta

    // Source docDefaults elements (readonly copies for resolveStyleChain).
    // nil if source has no docDefaults.
    srcDocDefaultsPPr *etree.Element
    srcDocDefaultsRPr *etree.Element

    // copiedClones collects style clones during Pass 1 (merge loop).
    // Pass 2 (fixupCopiedStyles) processes them after styleMap is complete.
    copiedClones []copiedStyleEntry
}

type copiedStyleEntry struct {
    clone    *etree.Element  // deep-copied style element, already added to target
    srcStyle *oxml.CT_Style  // original source style (readonly, for chain resolution)
}
```

### Step 2 — Двухпроходная архитектура (Pass 1 + Pass 2)

**Файл:** `pkg/docx/resourceimport_styles.go`

**Pass 1 — merge loop (строит styleMap, собирает клоны):**

`copyStyleToTarget` изменяется: клонирует, strip default, добавляет в target, записывает в styleMap, **но НЕ вызывает** `remapStyleRefsInElement` и `compensateDocDefaults`. Вместо этого сохраняет clone в `copiedClones`:

```go
func (ri *ResourceImporter) copyStyleToTarget(srcStyle *oxml.CT_Style, targetId string) error {
    tgtStyles, err := ri.targetStyles()
    if err != nil { ... }

    clone := srcStyle.RawElement().Copy()

    // Strip w:default="1" — target already has its own defaults.
    stripDefaultFlag(clone)

    // Rename styleId and display name when copying under a new ID.
    if targetId != srcStyle.StyleId() {
        clone.CreateAttr("w:styleId", targetId)
        if nameEl := findChild(clone, "w", "name"); nameEl != nil { ... }
        // semiHidden + unhideWhenUsed for renamed copies
        ...
    }

    // Remap numId (numbering was already fully imported in Phase 1,
    // numIdMap is complete at this point — safe to remap immediately).
    ri.remapNumIdsInElement(clone)

    // DO NOT remap basedOn/next/link here — styleMap is incomplete.
    // DO NOT compensate docDefaults here — need full styleMap.
    // Both are deferred to Pass 2 (fixupCopiedStyles).

    tgtStyles.RawElement().AddChild(clone)
    ri.styleMap[srcStyle.StyleId()] = targetId

    // Collect for Pass 2.
    ri.copiedClones = append(ri.copiedClones, copiedStyleEntry{
        clone:    clone,
        srcStyle: srcStyle,
    })
    return nil
}
```

**Pass 2 — fixupCopiedStyles (после merge loop):**

```go
// fixupCopiedStyles performs deferred operations on copied style clones
// that require the full styleMap to be available:
//
//   1. Remap basedOn/next/link references through styleMap.
//      (Fixes BUG 1: during BFS merge, children are processed before
//      parents, so parent mappings are not yet in styleMap.)
//
//   2. Compensate docDefaults delta for paragraph/character styles.
//      (Fixes Defect A: need to know which parents are copied vs mapped
//      to determine compensation strategy.)
//
// Called once, after the merge loop completes in importStyles().
func (ri *ResourceImporter) fixupCopiedStyles() {
    for _, entry := range ri.copiedClones {
        // 1. Remap basedOn/next/link — styleMap is now complete.
        ri.remapStyleRefsInElement(entry.clone)

        // 2. Compensate docDefaults delta.
        ri.compensateDocDefaults(entry.clone, entry.srcStyle)
    }
}
```

**Обновлённый importStyles:**

```go
func (ri *ResourceImporter) importStyles() error {
    if ri.styleDone {
        return nil
    }
    ri.styleDone = true

    // Step 0: Compute docDefaults delta (readonly, one-time).
    if err := ri.computeDocDefaultsDelta(); err != nil {
        return fmt.Errorf("docx: computing docDefaults delta: %w", err)
    }

    // Step 1: Detect default paragraph style mismatch.
    if err := ri.detectDefaultStyleMismatch(); err != nil {
        return fmt.Errorf("docx: detecting default style mismatch: %w", err)
    }

    // Step 2: Collect seed styleIds from source body.
    seedIds := collectStyleIdsFromBody(ri.sourceDoc)
    if ri.srcDefaultParaStyleId != "" {
        seedIds = appendUnique(seedIds, ri.srcDefaultParaStyleId)
    }
    if len(seedIds) == 0 {
        return nil
    }

    // Step 3: Compute transitive closure.
    closure := ri.collectStyleClosure(seedIds)

    // Pass 1: Merge each style — builds styleMap, collects clones.
    for _, srcStyle := range closure {
        if err := ri.mergeOneStyle(srcStyle); err != nil {
            return err
        }
    }

    // Pass 2: Fixup clones — remap refs + compensate delta.
    // styleMap is now complete for all styles in the closure.
    ri.fixupCopiedStyles()

    return nil
}
```

### Step 3 — Fix A: resolveStyleChain с source docDefaults как базой

**Файл:** `pkg/docx/resourceimport_styles.go`, метод `resolveStyleChain`

**Проблема:** `resolveStyleChain` начинает с пустого pPr/rPr и накладывает цепочку стилей. Но стили проектировались поверх source docDefaults — свойства, не заданные нигде в цепочке, должны браться из source docDefaults, а не оставаться пустыми.

**Это Fix A из ENTERPRISE_PLAN:** resolve chain начинает с source docDefaults как базы.

**Влияет на:** expand path (KeepSourceFormatting, KeepDifferentStyles) — resolved properties теперь включают source docDefaults для свойств, не заданных в цепочке. При инжекции в direct attributes эти свойства перекроют target docDefaults.

**Изменения:**

```go
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
            break
        }
        visited[id] = true
        chain = append(chain, current)
        basedOn, _ := current.BasedOnVal()
        if basedOn == "" {
            break
        }
        current = srcStyles.GetByID(basedOn)
    }

    // [NEW] Start with source docDefaults as the base of the chain.
    // This ensures properties inherited from source docDefaults are
    // included in the resolved result. Without this, resolved properties
    // are incomplete — any property not explicitly set in the chain
    // would be missing, causing it to inherit from TARGET docDefaults
    // after insertion (wrong visual result).
    if ri.srcDocDefaultsPPr != nil {
        pPr = ri.srcDocDefaultsPPr.Copy()
    }
    if ri.srcDocDefaultsRPr != nil {
        rPr = ri.srcDocDefaultsRPr.Copy()
    }

    // Merge from base to derived (so derived overrides base).
    for i := len(chain) - 1; i >= 0; i-- {
        raw := chain[i].RawElement()
        if p := findChild(raw, "w", "pPr"); p != nil {
            if pPr == nil {
                pPr = p.Copy()
            } else {
                overridePropertiesDeep(pPr, p)
            }
        }
        if r := findChild(raw, "w", "rPr"); r != nil {
            if rPr == nil {
                rPr = r.Copy()
            } else {
                overridePropertiesDeep(rPr, r)
            }
        }
    }

    // Strip pStyle/rStyle from resolved properties.
    if pPr != nil {
        removeChild(pPr, "w", "pStyle")
    }
    if rPr != nil {
        removeChild(rPr, "w", "rStyle")
    }
    return
}
```

**Хранение source docDefaults** — заполняется в `computeDocDefaultsDelta`:

```go
func (ri *ResourceImporter) computeDocDefaultsDelta() error {
    srcStyles, err := ri.sourceStyles()
    if err != nil { return nil } // no styles part — no delta
    tgtStyles, err := ri.targetStyles()
    if err != nil { return err }

    // Store source docDefaults for resolveStyleChain (Fix A).
    ri.srcDocDefaultsPPr = srcStyles.DocDefaultsPPr() // readonly ref, Copy() in resolve
    ri.srcDocDefaultsRPr = srcStyles.DocDefaultsRPr()

    // Compute delta (Fix B).
    srcPPr := srcStyles.DocDefaultsPPr()
    tgtPPr := tgtStyles.DocDefaultsPPr()
    srcRPr := srcStyles.DocDefaultsRPr()
    tgtRPr := tgtStyles.DocDefaultsRPr()

    deltaPPr := diffProperties(srcPPr, tgtPPr, ooxmlImplicitPPr)
    deltaRPr := diffProperties(srcRPr, tgtRPr, ooxmlImplicitRPr)

    if deltaPPr == nil && deltaRPr == nil {
        return nil // identical docDefaults — fast path
    }
    ri.docDefDelta = &docDefaultsDelta{pPr: deltaPPr, rPr: deltaRPr}
    return nil
}
```

### Step 4 — Fix B: compensateDocDefaults для скопированных стилей

**Файл:** `pkg/docx/resourceimport_styles.go`

Вызывается в Pass 2 (`fixupCopiedStyles`), когда styleMap полный.

```go
// compensateDocDefaults injects source docDefaults delta into a copied
// style clone so that properties inherited from source docDefaults are
// preserved in the target document (which has different docDefaults).
//
// Only processes paragraph and character styles. Table, numbering, and
// other style types are skipped — docDefaults pPr/rPr do not cascade
// into tblPr or list formatting per OOXML spec.
//
// Compensation strategy depends on the style's position in the basedOn chain:
//
//   Root style (no basedOn, or basedOn not in source):
//     Inject ALL delta properties not explicitly defined in the clone.
//
//   Style whose basedOn parent is also COPIED (in copiedClones):
//     Inject NOTHING — parent's compensation covers transitively.
//
//   Style whose basedOn parent is MAPPED to target (in styleMap but not copied):
//     Resolve full SOURCE chain → inject delta properties not defined
//     anywhere in the source chain. Properties defined in a mapped parent
//     are NOT compensated (UseDestinationStyles intentionally uses target
//     definition — see §2.3).
//
// Precondition: styleMap is complete (called in Pass 2).
func (ri *ResourceImporter) compensateDocDefaults(clone *etree.Element, srcStyle *oxml.CT_Style) {
    if ri.docDefDelta == nil {
        return
    }

    // Only paragraph and character styles inherit from pPrDefault/rPrDefault.
    styleType := clone.SelectAttrValue("w:type", "")
    if styleType != "paragraph" && styleType != "character" {
        return
    }

    // Determine compensation strategy based on basedOn chain.
    strategy := ri.compensationStrategy(srcStyle)

    switch strategy {
    case compensateAll:
        // Root style: inject all delta not in clone itself.
        ri.injectDelta(clone, nil)
    case compensateNone:
        // Parent also copied: skip (transitive coverage).
        return
    case compensateUncovered:
        // Parent in target: resolve source chain, inject uncovered delta.
        resolvedPPr, resolvedRPr := ri.resolveStyleChain(srcStyle)
        ri.injectDelta(clone, &resolvedChain{pPr: resolvedPPr, rPr: resolvedRPr})
    }
}

type compensationAction int

const (
    compensateAll       compensationAction = iota // root: inject all delta
    compensateNone                                // parent copied: skip
    compensateUncovered                           // parent in target: inject uncovered
)

// compensationStrategy determines how to compensate a copied style.
// Uses styleMap (complete at this point) and copiedClones to classify.
func (ri *ResourceImporter) compensationStrategy(srcStyle *oxml.CT_Style) compensationAction {
    basedOn, _ := srcStyle.BasedOnVal()
    if basedOn == "" {
        return compensateAll // no parent → root
    }

    // Check if parent was copied (exists in copiedClones).
    if ri.isParentCopied(basedOn) {
        return compensateNone // parent copied → transitive
    }

    // Parent is mapped to target or not in closure at all.
    return compensateUncovered
}

// isParentCopied checks if a style with the given source ID was copied
// (as opposed to mapped to an existing target style).
func (ri *ResourceImporter) isParentCopied(srcStyleId string) bool {
    for _, entry := range ri.copiedClones {
        if entry.srcStyle.StyleId() == srcStyleId {
            return true
        }
    }
    return false
}

type resolvedChain struct {
    pPr *etree.Element
    rPr *etree.Element
}

// injectDelta merges delta properties into clone's pPr/rPr.
// If chain is nil (root style): injects all delta not in clone.
// If chain is non-nil: injects only delta properties NOT present in chain.
func (ri *ResourceImporter) injectDelta(clone *etree.Element, chain *resolvedChain) {
    delta := ri.docDefDelta

    if delta.pPr != nil {
        clonePPr := findChild(clone, "w", "pPr")
        if ri.shouldInjectGroup(delta.pPr, clonePPr, chain, "pPr") {
            if clonePPr == nil {
                clonePPr = etree.NewElement("w:pPr")
                // pPr должен быть перед rPr в w:style per schema order
                clone.InsertChildAt(insertIndexForPPr(clone), clonePPr)
            }
            ri.injectGroupFiltered(clonePPr, delta.pPr, chain, "pPr")
        }
    }

    if delta.rPr != nil {
        cloneRPr := findChild(clone, "w", "rPr")
        if ri.shouldInjectGroup(delta.rPr, cloneRPr, chain, "rPr") {
            if cloneRPr == nil {
                cloneRPr = etree.NewElement("w:rPr")
                clone.AddChild(cloneRPr)
            }
            ri.injectGroupFiltered(cloneRPr, delta.rPr, chain, "rPr")
        }
    }
}

// injectGroupFiltered merges delta children into dst, skipping properties
// that are already in dst (explicit > delta) or covered by chain.
func (ri *ResourceImporter) injectGroupFiltered(
    dst *etree.Element,
    deltaSrc *etree.Element,
    chain *resolvedChain,
    group string, // "pPr" or "rPr"
) {
    for _, deltaChild := range deltaSrc.ChildElements() {
        // Skip if clone already defines this property (explicit > delta).
        if findChild(dst, deltaChild.Space, deltaChild.Tag) != nil {
            continue
        }
        // Skip if source chain covers this property (only for compensateUncovered).
        if chain != nil {
            var chainGroup *etree.Element
            if group == "pPr" {
                chainGroup = chain.pPr
            } else {
                chainGroup = chain.rPr
            }
            if chainGroup != nil && findChild(chainGroup, deltaChild.Space, deltaChild.Tag) != nil {
                continue
            }
        }
        // Inject delta property.
        dst.AddChild(deltaChild.Copy())
    }
}
```

### Step 5 — Built-in style name-based resolution

**Файл:** `pkg/docx/resourceimport_styles.go`, метод `mergeOneStyle`

**Изменения в mergeOneStyle:**

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
    if err != nil { ... }

    existing := tgtStyles.GetByID(id)

    // [NEW] Name-based fallback for built-in styles.
    // Localized documents use locale-specific styleId for built-in styles:
    //   RU: "a" for Normal, "a1" for Normal Table
    //   EN: "Normal", "TableNormal"
    // OOXML guarantees w:name uniqueness per type for built-in styles.
    if existing == nil && srcStyle.IsBuiltin() {
        existing = ri.findBuiltinByName(srcStyle, tgtStyles)
    }

    // --- Style NOT in target: always copy (all 3 modes agree) ---
    if existing == nil {
        return ri.copyStyleToTarget(srcStyle, id)
    }

    // --- Style EXISTS in target: behavior depends on mode ---
    // [CHANGED] Use actual target styleId (may differ from source id).
    targetId := existing.StyleId()
    switch ri.importFormatMode {

    case UseDestinationStyles:
        ri.styleMap[id] = targetId

    case KeepSourceFormatting:
        if ri.opts.ForceCopyStyles {
            return ri.copyStyleToTarget(srcStyle, ri.uniqueStyleId(id))
        }
        ri.expandStyles[id] = srcStyle
        ri.styleMap[id] = ri.targetDefaultParaStyleId()

    case KeepDifferentStyles:
        if stylesContentEqual(srcStyle, existing) {
            ri.styleMap[id] = targetId
        } else {
            if ri.opts.ForceCopyStyles {
                return ri.copyStyleToTarget(srcStyle, ri.uniqueStyleId(id))
            }
            ri.expandStyles[id] = srcStyle
            ri.styleMap[id] = ri.targetDefaultParaStyleId()
        }
    }
    return nil
}
```

**Новый метод `findBuiltinByName`:**

```go
// findBuiltinByName resolves a source built-in style to a target built-in
// style using w:name and w:type matching. Returns nil if no match found.
//
// This handles the OOXML localization pattern where built-in styles have
// locale-specific styleId but identical w:name values across locales.
// Example: Normal → "Normal" (EN), "a" (RU), "a" (JP).
func (ri *ResourceImporter) findBuiltinByName(
    srcStyle *oxml.CT_Style,
    tgtStyles *oxml.CT_Styles,
) *oxml.CT_Style {
    srcName, err := srcStyle.NameVal()
    if err != nil || srcName == "" {
        return nil
    }
    srcType := srcStyle.Type()

    candidate := tgtStyles.GetByName(srcName)
    if candidate == nil {
        return nil
    }
    // Must match type (paragraph "Normal" ≠ character "Normal").
    if candidate.Type() != srcType {
        return nil
    }
    // Must be built-in too (avoid matching custom style with same name).
    if candidate.CustomStyle() {
        return nil
    }
    return candidate
}
```

### Step 6 — Устранение дубликата w:default="1"

**Файл:** `pkg/docx/resourceimport_styles.go`, метод `copyStyleToTarget`

Выполняется при всех копированиях — и когда name fallback не нашёл match, и при ForceCopyStyles:

```go
// stripDefaultFlag removes w:default="1" from a cloned style element.
// A copied style must never declare itself as default — the target
// already has its own default for each type. Two defaults of the
// same type produce invalid OOXML (§17.7.4.9).
func stripDefaultFlag(clone *etree.Element) {
    filtered := clone.Attr[:0]
    for _, attr := range clone.Attr {
        if attr.Key == "default" && (attr.Space == "w" ||
            strings.Contains(attr.Space, "wordprocessingml")) {
            continue // strip
        }
        filtered = append(filtered, attr)
    }
    clone.Attr = filtered
}
```

### Step 7 — Доступ к docDefaults через etree FindElement

**Файл:** `pkg/docx/oxml/styles_custom.go`

Используем `FindElement` с XPath — одна строка вместо тройного `findChild`. Не дублируем `findChild` между пакетами:

```go
// DocDefaultsRPr returns the w:rPr element inside w:docDefaults/w:rPrDefault,
// or nil if not present. The returned element belongs to the styles.xml tree —
// do NOT modify it directly. Use Copy() if modification is needed.
//
// docDefaults are preserved through round-trip by etree but not modeled
// in the codegen schema (they are rarely needed outside of import).
func (ss *CT_Styles) DocDefaultsRPr() *etree.Element {
    return ss.RawElement().FindElement("w:docDefaults/w:rPrDefault/w:rPr")
}

// DocDefaultsPPr returns the w:pPr element inside w:docDefaults/w:pPrDefault,
// or nil. Same immutability contract as DocDefaultsRPr.
func (ss *CT_Styles) DocDefaultsPPr() *etree.Element {
    return ss.RawElement().FindElement("w:docDefaults/w:pPrDefault/w:pPr")
}
```

Не нужен `findChildElement` в oxml пакете. Не нужен экспорт `findChild` из `resourceimport_styles.go`.

---

## 4. Файловая карта изменений

| Файл | Изменения | Объём |
|---|---|---|
| `pkg/docx/oxml/styles_custom.go` | `DocDefaultsRPr()`, `DocDefaultsPPr()` | +12 строк |
| `pkg/docx/resourceimport.go` | Поля `docDefDelta`, `srcDocDefaultsPPr/RPr`, `copiedClones`, тип `copiedStyleEntry` | +20 строк |
| `pkg/docx/resourceimport_styles.go` | `docDefaultsDelta` struct, `computeDocDefaultsDelta()`, `diffProperties()`, `compensateDocDefaults()`, `compensationStrategy()`, `injectDelta()`, `findBuiltinByName()`, `stripDefaultFlag()`, `fixupCopiedStyles()`, изменения в `mergeOneStyle`, `copyStyleToTarget`, `importStyles`, `resolveStyleChain` | +280 строк |
| `pkg/docx/resourceimport_styles_test.go` | Тесты на все новые функции | +300 строк |

**Новые файлы:** 0 (все изменения в существующих файлах, в соответствии с архитектурой проекта).

---

## 5. Тест-план

### 5.1 Unit-тесты (resourceimport_styles_test.go)

```
=== diffProperties ===

TestDiffProperties_BothNil
    source и target оба nil → return nil

TestDiffProperties_SourceNil_TargetHasSz
    source nil, target имеет sz=22 →
    delta содержит sz=20 (OOXML implicit default)

TestDiffProperties_TargetNil_SourceHasSz
    source имеет sz=24, target nil →
    delta = nil (target не задаёт sz → source value уже правильный)

TestDiffProperties_Different_AttributeLevel
    source: spacing after=0 line=240 lineRule=auto
    target: spacing after=160 line=259 lineRule=auto
    → delta: spacing after=0 line=240 (lineRule совпадает — НЕ включается)

TestDiffProperties_Identical
    оба одинаковы → return nil

TestDiffProperties_RFonts_SourceExplicit_TargetTheme
    source: rFonts ascii=TNR eastAsia=Batang
    target: rFonts asciiTheme=minorHAnsi
    → delta содержит rFonts с source-атрибутами

=== computeDocDefaultsDelta ===

TestComputeDocDefaultsDelta_BothEmpty
    source и target без docDefaults → delta == nil

TestComputeDocDefaultsDelta_SourceEmpty_TargetHasSpacingAndSz
    source: pPrDefault пустой, rPrDefault пустой
    target: spacing after=160 line=259, sz=22
    → delta.pPr = spacing after=0 line=240
    → delta.rPr = sz=20 szCs=20 rFonts=TNR

TestComputeDocDefaultsDelta_Identical
    оба идентичны → docDefDelta == nil

=== compensateDocDefaults ===

TestCompensateDocDefaults_RootStyle
    Normal без basedOn, delta имеет sz=20, spacing →
    clone получает sz=20 + spacing injected

TestCompensateDocDefaults_StyleWithExplicitProp
    стиль уже имеет sz=24, delta.rPr имеет sz=20 →
    clone сохраняет sz=24 (explicit > delta)

TestCompensateDocDefaults_ParentCopied
    стиль с basedOn, parent в copiedClones →
    clone НЕ изменяется (parent покроет)

TestCompensateDocDefaults_ParentInTarget
    стиль с basedOn, parent замаплен в target →
    resolve source chain → inject uncovered delta only

TestCompensateDocDefaults_TableStyle_Skipped
    table style → compensateDocDefaults = no-op

TestCompensateDocDefaults_NumberingStyle_Skipped
    numbering style → compensateDocDefaults = no-op

TestCompensateDocDefaults_CharacterStyle
    character style → компенсация rPr delta, pPr игнорируется

=== resolveStyleChain (обновлённый) ===

TestResolveStyleChain_IncludesDocDefaults
    source docDefaults: sz=24, spacing after=100
    style chain: Normal (rFonts=Arial)
    → resolved: rFonts=Arial + sz=24 + spacing after=100

TestResolveStyleChain_ChainOverridesDocDefaults
    source docDefaults: sz=24
    Normal: sz=20
    → resolved: sz=20 (chain overrides docDefaults)

TestResolveStyleChain_NoDocDefaults
    srcDocDefaultsPPr/RPr = nil → поведение как раньше

=== fixupCopiedStyles ===

TestFixupCopiedStyles_RemapsBasedOn
    child копируется раньше parent (BFS order) →
    после fixup: basedOn ссылается на правильный target styleId

TestFixupCopiedStyles_RemapsAndCompensates
    root style + child style both copied →
    root получает delta, child пропускается

=== findBuiltinByName ===

TestFindBuiltinByName_LocalizedNormal
    source "Normal" (id=Normal) vs target "Normal" (id=a) → match

TestFindBuiltinByName_NoMatch
    source "Achievement" (custom) → nil (IsBuiltin check in caller)

TestFindBuiltinByName_TypeMismatch
    source paragraph "Normal" vs target character "Normal" → nil

TestFindBuiltinByName_CustomWithSameName
    target has custom style named "Normal" → nil (CustomStyle check)

=== stripDefaultFlag ===

TestStripDefaultFlag_Present
    clone с w:default="1" → атрибут удалён

TestStripDefaultFlag_Absent
    clone без default → no-op

=== mergeOneStyle ===

TestMergeOneStyle_NameFallback_UseDestination
    source Normal (id="Normal") → target "a" (same name) →
    styleMap["Normal"] = "a", НЕ копирует

TestMergeOneStyle_NameFallback_KeepSource
    source Normal matched by name → expandStyles["Normal"] = srcNormal
    styleMap["Normal"] = targetDefaultParaStyleId
```

### 5.2 Integration-тесты

```
TestReplaceWithContent_DocDefaultsDelta_SzPreserved
    source: docDefaults нет sz (10pt), стиль Normal: Arial
    target: docDefaults sz=22 (11pt)
    → output: скопированный Normal содержит explicit sz=20

TestReplaceWithContent_DocDefaultsDelta_SpacingPreserved
    source: pPrDefault пустой (single spacing)
    target: pPrDefault spacing after=160, line=259
    → output: скопированные стили содержат explicit spacing=single

TestReplaceWithContent_ExpandPath_IncludesDocDefaults
    KeepSourceFormatting mode, source Normal conflicts with target →
    expanded paragraphs contain sz from source docDefaults

TestReplaceWithContent_LocalizedBuiltin_NoDuplicate
    source: Normal (id="Normal"), target: Normal (id="a")
    → output: один Normal (id="a"), styleMap["Normal"]="a"

TestReplaceWithContent_NoDuplicateDefaultFlag
    → output: ровно один w:default="1" per type

TestReplaceWithContent_BFS_BasedOnRemapping
    source: BodyText basedOn Normal, Normal copies as "Normal",
    BodyText processed before Normal in BFS →
    after fixup: BodyText basedOn correctly remapped

TestReplaceWithContent_MultiSource_TwoMarks
    template с двумя метками (<mark1>) и (<mark2>)
    source1 и source2 — разные документы с разными docDefaults
    → output: оба source корректно компенсированы
```

### 5.3 Visual regression test

Обновить `visual-regtest/replace-user-mark-batch`:
- Результат `out/1_page.docx` визуально совпадает с mark `in/mark/1_page.docx`
- Межстрочные интервалы, размеры шрифтов, отступы идентичны source
- Проверить все 6 вариантов: default, keep_src, keep_src_force, keep_diff, keep_diff_force, keep_num

---

## 6. Порядок реализации

```
Phase 1: Инфраструктура
  1. DocDefaultsRPr/DocDefaultsPPr в oxml/styles_custom.go
  2. docDefaultsDelta struct + computeDocDefaultsDelta + diffProperties
  3. Новые поля в ResourceImporter (copiedClones, srcDocDefaults*, docDefDelta)
  4. Unit-тесты на diffProperties + computeDocDefaultsDelta

Phase 2: Двухпроходная архитектура
  5. Изменить copyStyleToTarget — defer remap/compensate, collect clones
  6. Реализовать fixupCopiedStyles (Pass 2)
  7. Изменить importStyles — вызов fixupCopiedStyles после merge loop
  8. Unit-тесты на fixupCopiedStyles (BFS remap correctness)

Phase 3: Name-based style resolution
  9. findBuiltinByName
  10. Изменить mergeOneStyle — name fallback + targetId
  11. stripDefaultFlag в copyStyleToTarget
  12. Unit-тесты на name resolution + no-duplicate default

Phase 4: Delta compensation (copy path — Fix B)
  13. compensateDocDefaults + compensationStrategy + injectDelta
  14. Unit-тесты на compensation strategies

Phase 5: Expand path fix (Fix A)
  15. Изменить resolveStyleChain — source docDefaults как база
  16. Unit-тесты на resolveStyleChain с docDefaults

Phase 6: Integration + Regression
  17. Integration-тесты (все из §5.2)
  18. go test ./...
  19. visual-regtest/replace-user-mark-batch — все 6 вариантов
```

---

## 7. Инварианты (что гарантировать)

1. **Идемпотентность:** повторный вызов importStyles() — no-op (существующий styleDone флаг).

2. **Source immutability:** source document никогда не модифицируется. Все delta-операции на Copy(). `srcDocDefaultsPPr`/`srcDocDefaultsRPr` — readonly ссылки, Copy() при использовании в resolveStyleChain.

3. **Multi-source корректность:** каждый ResourceImporter имеет свой docDefDelta и srcDocDefaults. При вставке mark1 и mark2 с разными docDefaults каждый получает свою компенсацию.

4. **Backward compatibility:** если source и target docDefaults одинаковы — delta == nil, compensateDocDefaults — no-op. resolveStyleChain начинает с source docDefaults, но если они пустые (nil) — поведение идентично текущему.

5. **Один default per type:** после import, output styles.xml содержит ровно один `w:default="1"` для каждого типа стиля.

6. **basedOn chain integrity:** Pass 2 выполняет remap когда styleMap полный — все ссылки корректны.

7. **Type safety:** delta compensation применяется только к paragraph и character стилям. Table, numbering, и другие типы пропускаются.

---

## 8. Что НЕ входит в этот fix

| Исключение | Причина |
|---|---|
| Добавление docDefaults в codegen schema | Избыточно. Readonly accessor через `FindElement` достаточен. Кодогенерация docDefaults — отдельный task |
| Theme resolution (themeFont → actual font name) | В плане ENTERPRISE уже отмечено как "не нужно". Aspose тоже не делает. Мы инжектируем source explicit значения; theme-атрибуты не трогаем |
| Модификация target docDefaults | Target docDefaults принадлежат template. Менять их — нарушение контракта |
| Merge docDefaults (combine source + target) | Невозможно корректно: один documentDefaults на документ. Aspose не мержит |
| Компенсация свойств mapped (не скопированных) стилей в UseDestinationStyles | Name-based matching находит target built-in стиль и использует его определение. Explicit свойства source стиля (напр. rFonts=Arial в source Normal) заменяются target определением. **Это штатное поведение UseDestinationStyles** — режим явно предполагает использование target стилей при конфликте. Для сохранения source форматирования используйте KeepSourceFormatting |

---

## 9. Сложные edge-cases

### 9.1 Стиль без basedOn, не root (standalone)

```xml
<!-- Source: JobTitle — no basedOn, not based on Normal -->
<w:style w:styleId="JobTitle">
  <w:pPr><w:spacing w:after="60" w:line="220" w:lineRule="atLeast"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Arial Black" .../></w:rPr>
</w:style>
```

JobTitle не имеет basedOn → compensateAll. Delta для `sz=20` инжектируется (sz не задан). Spacing не инжектируется (clone уже задаёт after, line, lineRule — explicit > delta).

### 9.2 Цепочка A → B → C, где B в target, A и C копируются

```
C basedOn B basedOn A
B найден в target по name → styleMap[B_src] = B_tgt
A и C копируются
```

**Pass 2 обработка:**

A — root (no basedOn), compensateAll → delta injected в A.
C — basedOn B. B не в copiedClones (mapped) → compensateUncovered.
Resolve source chain (C → B_src → A_src): свойство X определено в B_src → chain покрывает → delta для X НЕ инжектируется в C.

**Важно:** в target цепочка C → B_tgt. B_tgt может определять другое значение X чем B_src. Это **intended behavior UseDestinationStyles** — mapped стили используют target определение.

### 9.3 w:spacing — multi-attribute element

`w:spacing` имеет несколько атрибутов: `w:after`, `w:before`, `w:line`, `w:lineRule`. `diffProperties` работает с **attribute-level** гранулярностью:

```
source effective: <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
target:           <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
delta:            <w:spacing w:after="0" w:line="240"/>
                  (lineRule совпадает — НЕ включается в delta)
```

При инжекции в стиль, который уже имеет `<w:spacing w:after="60" w:line="220">`:
- `w:after` — clone задаёт 60 → findChild("spacing") exists → skip (explicit > delta)
- Результат: ничего не меняется. Корректно.

Стиль без spacing (compensateAll):
- spacing не в clone → inject `<w:spacing w:after="0" w:line="240"/>` из delta.

### 9.4 KeepSourceFormatting с name match — expand path (Fix A)

```
Source: Normal (id="Normal") — basedOn=none, rPr: rFonts=Arial, no sz
Target: "a" (name="Normal") — matched by name
Source docDefaults: rPr: rFonts=TNR, no sz (implicit 10pt)
```

В KeepSourceFormatting: expandStyles["Normal"] = srcNormal.

`expandDirectFormatting` вызывает `resolveStyleChain(srcNormal)`:
- **С Fix A:** chain начинает с source docDefaults (rFonts=TNR, implicit sz=20).
- Override: Normal rPr (rFonts=Arial) перезаписывает docDefaults rFonts.
- Resolved rPr: rFonts=Arial (from Normal) + sz=20 (from docDefaults, not overridden).

`expandParagraphStyle` merges resolved properties в direct attrs → абзац получает rFonts=Arial и sz=20. Корректно.

**Без Fix A:** resolved rPr = только rFonts=Arial. sz отсутствует → наследуется от target docDefaults (22 = 11pt). Неверно.

### 9.5 BFS порядок и Pass 2 (fix BUG 1)

```
BFS closure: [TableGrid, Objective, Achievement, Normal, BodyText]
                ↑ children first          parents last ↑
```

**Pass 1 (merge loop):** все стили обрабатываются, styleMap заполняется. Normal matched by name → styleMap["Normal"]="a". BodyText не в target → copied.

**Pass 2 (fixup):** BodyText clone имеет basedOn="Normal". `remapStyleRefsInElement` меняет на basedOn="a" (используя полный styleMap). Затем `compensateDocDefaults`: BodyText basedOn Normal, Normal mapped (не copied) → compensateUncovered.

Без Pass 2 (старый код): при обработке BodyText в merge loop, Normal ещё не в styleMap → basedOn="Normal" не ремапится → битая ссылка в target.
