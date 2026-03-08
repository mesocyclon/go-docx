# Проект: Шаблонизатор документов для go-docx

**Статус:** Проектирование
**Версия:** 1.0
**Дата:** 2026-03-07

---

## 1. Краткое описание

Внедрение в библиотеку go-docx шаблонизатора документов, поддерживающего три типа маркеров:

| Маркер | Формат | Действие |
|--------|--------|----------|
| Текстовый | `[[ключ]]` | Замена маркера на строку текста |
| Табличный | `\|\|ключ\|\|` | Замена маркера на таблицу |
| Пользовательский | `((ключ))` | Вставка содержимого документа-источника (без header/footer) с сохранением стилей. Поддерживает рекурсию: вставленный документ может содержать любые виды маркеров |

---

## 2. Требования

### 2.1 Функциональные

| ID | Требование |
|----|-----------|
| FR-1 | Текстовые маркеры `[[ключ]]` заменяются на строковое значение |
| FR-2 | Табличные маркеры `\|\|ключ\|\|` заменяются на таблицу из `TableData` |
| FR-3 | Пользовательские маркеры `((ключ))` заменяются на body-контент документа-источника |
| FR-4 | При вставке `(())` сохраняются стили, нумерация, сноски, изображения источника |
| FR-5 | При вставке `(())` header/footer из источника НЕ переносятся |
| FR-6 | Пользовательские маркеры поддерживают рекурсию: вставленный документ может содержать `[[]]`, `\|\|\|\|`, `(())` |
| FR-7 | Маркеры обрабатываются в body, headers, footers, comments, table cells |
| FR-8 | Маркеры, разбитые на несколько XML-ранов (split runs), корректно обнаруживаются |
| FR-9 | Маркеры без соответствующих данных остаются в документе как есть |
| FR-10 | Множественные вхождения одного маркера заменяются все |

### 2.2 Нефункциональные

| ID | Требование |
|----|-----------|
| NFR-1 | Защита от бесконечной рекурсии: лимит глубины (по умолчанию 10) |
| NFR-2 | Обнаружение цикла A→B→A на уровне одного прохода |
| NFR-3 | Единый `ResourceImporter` для всех `(())` замен в одном вызове |
| NFR-4 | Нулевые изменения в существующих файлах библиотеки |
| NFR-5 | API согласован с существующим стилем (variadic options, error return) |
| NFR-6 | Документ не изменяется до вызова `ExecuteTemplate` |

---

## 3. Анализ текущей архитектуры

### 3.1 Существующие механизмы замены

Шаблонизатор строится **поверх** трех существующих примитивов замены:

```
Document.ReplaceText(old, new string) (int, error)
Document.ReplaceWithTable(old string, td TableData) (int, error)
Document.ReplaceWithContent(old string, cd ContentData) (int, error)
```

Каждый из них:
- Обрабатывает body, все headers/footers (с дедупликацией по `*StoryPart`), comments
- Корректно работает с маркерами, разбитыми на несколько XML-ранов
- Рекурсивно обходит table cells
- Возвращает количество произведенных замен

### 3.2 Цепочка вызовов (существующая)

```
Document.Replace*()
  ├── Body.replace*()                      ← BlockItemContainer
  │     └── replaceTagWithElements()       ← двухфазный движок
  │           ├── Phase 1: SplitParagraphAtTags() → splice
  │           └── Phase 2: recurse into table cells
  ├── Section[].Header/Footer.replace*Dedup()  ← дедупликация по StoryPart*
  │     └── BlockItemContainer.replace*()
  └── Comments[].replace*()
        └── BlockItemContainer.replace*()
```

### 3.3 Ресурсный импорт (для ReplaceWithContent)

```
ResourceImporter (один на вызов ReplaceWithContent)
  ├── importNumbering()     ← numIdMap, absNumIdMap
  ├── importStyles()        ← styleMap, UseDestinationStyles
  ├── importFootnotes()     ← footnoteIdMap
  ├── importEndnotes()      ← endnoteIdMap
  └── remapAll()            ← единый DFS-проход
```

**Ключевое свойство:** `ResourceImporter` создается один раз и используется для body, headers, footers, comments, обеспечивая консистентность маппингов.

### 3.4 Что шаблонизатор получает "бесплатно"

| Возможность | Обеспечивает |
|-------------|-------------|
| Маркеры через run boundaries | `collectTextAtoms` + `SplitParagraphAtTags` |
| Замена в headers/footers | Цикл по sections с `applyToContainerDedup` |
| Замена в comments | Итерация `Comments.Iter()` |
| Замена в table cells | Phase 2 в `replaceTagWithElements` |
| Перенос стилей, нумерации, сносок | `ResourceImporter` |
| Исключение header/footer источника | `prepareContentElements` пропускает `w:sectPr` |
| Дедупликация изображений | SHA-256 в `WmlPackage.GetOrAddImagePart` |
| Дедупликация header/footer обработки | `seen map[*parts.StoryPart]bool` |

---

## 4. Проектирование

### 4.1 Обзор архитектуры

```
┌─────────────────────────────────────────────────────────┐
│  Template Engine (НОВЫЙ СЛОЙ)                           │
│  ExecuteTemplate() — оркестратор                        │
│  TemplateData — модель данных                           │
│  templateRunner — многопроходный движок                 │
├─────────────────────────────────────────────────────────┤
│  Domain Layer (СУЩЕСТВУЮЩИЙ, без изменений)             │
│  Document.ReplaceText()                                 │
│  Document.ReplaceWithTable()                            │
│  Document.ReplaceWithContent()                          │
├─────────────────────────────────────────────────────────┤
│  Parts / OXML / OPC (СУЩЕСТВУЮЩИЕ, без изменений)      │
└─────────────────────────────────────────────────────────┘
```

Шаблонизатор — **тонкий оркестрационный слой**, который:
1. Формирует полные строки маркеров из ключей (`key` → `[[key]]`)
2. Управляет порядком и многопроходностью выполнения
3. Делегирует всю работу существующим `Replace*` методам

### 4.2 API

#### 4.2.1 TemplateData — модель данных

```go
// TemplateData содержит данные для подстановки в шаблон документа.
//
// Ключи — это идентификаторы маркеров без разделителей.
// Движок автоматически оборачивает их в соответствующие разделители:
//   - "name"  → ищет [[name]] в документе
//   - "items" → ищет ||items|| в документе
//   - "header" → ищет ((header)) в документе
type TemplateData struct {
    texts    map[string]string
    tables   map[string]TableData
    content  map[string]*Document
    options  templateOptions
}

type templateOptions struct {
    maxDepth int  // default: 10
    // Управляет поведением при обнаружении маркеров без данных.
    // false (default): маркеры без данных остаются в документе.
    // true: возвращать ошибку при обнаружении маркера без данных.
    strictMode bool
}
```

#### 4.2.2 Builder API

```go
// NewTemplateData создает пустой набор данных для шаблона.
func NewTemplateData() *TemplateData

// Text добавляет текстовую замену: [[key]] → value.
func (td *TemplateData) Text(key, value string) *TemplateData

// Table добавляет табличную замену: ||key|| → table.
func (td *TemplateData) Table(key string, data TableData) *TemplateData

// Content добавляет пользовательскую замену: ((key)) → body-контент документа.
// Header и footer документа-источника НЕ переносятся.
// Документ-источник может содержать маркеры любого типа (рекурсия).
func (td *TemplateData) Content(key string, doc *Document) *TemplateData

// MaxDepth устанавливает максимальную глубину рекурсии для (()) маркеров.
// По умолчанию: 10. Значение 0 запрещает рекурсию (один проход).
func (td *TemplateData) MaxDepth(depth int) *TemplateData

// StrictMode включает строгий режим: ошибка при обнаружении маркера без данных.
func (td *TemplateData) StrictMode() *TemplateData
```

#### 4.2.3 Метод Document

```go
// ExecuteTemplate подставляет данные из TemplateData в маркеры документа.
//
// Порядок обработки:
//  1. Пользовательские маркеры (()) — многопроходно до maxDepth
//  2. Табличные маркеры ||
//  3. Текстовые маркеры [[]]
//
// При ошибке документ может быть частично модифицирован (consistent
// с поведением существующих Replace* методов).
func (d *Document) ExecuteTemplate(data *TemplateData) error
```

#### 4.2.4 Пример использования

```go
// Открыть шаблон
tpl, err := docx.OpenFile("template.docx")
if err != nil {
    log.Fatal(err)
}

// Открыть документ-источник для пользовательской метки
header, err := docx.OpenFile("header.docx")
if err != nil {
    log.Fatal(err)
}

// Подготовить данные
data := docx.NewTemplateData().
    Text("client_name", "ООО Рога и Копыта").
    Text("date", "07.03.2026").
    Text("total", "1 500 000 руб.").
    Table("items", docx.TableData{
        Rows: [][]string{
            {"Услуга", "Кол-во", "Цена"},
            {"Разработка", "160 ч", "800 000"},
            {"Тестирование", "80 ч", "400 000"},
            {"Внедрение", "1", "300 000"},
        },
        Style: docx.StyleName("Table Grid"),
    }).
    Content("header", header).
    MaxDepth(5)

// Выполнить подстановку
if err := tpl.ExecuteTemplate(data); err != nil {
    log.Fatal(err)
}

// Сохранить результат
if err := tpl.SaveFile("result.docx"); err != nil {
    log.Fatal(err)
}
```

### 4.3 Алгоритм выполнения

```
ExecuteTemplate(data)
│
├─ validate(data)                     ← проверка входных данных
│
├─ Phase 1: Content markers (())     ← многопроходная рекурсия
│  │
│  └─ for depth := 0; depth < maxDepth; depth++
│     │
│     │  totalHits := 0
│     │
│     │  for key, source := range data.content
│     │  │  marker := "((" + key + "))"
│     │  │  n, err := doc.ReplaceWithContent(marker, ContentData{Source: source})
│     │  │  totalHits += n
│     │  │
│     │  if totalHits == 0 → break   ← ранний выход: маркеров больше нет
│     │
│     └─ if depth == maxDepth-1 && totalHits > 0
│        └─ return ErrMaxDepthExceeded (если strict) или warning
│
├─ Phase 2: Table markers ||
│  │
│  └─ for key, tableData := range data.tables
│     │  marker := "||" + key + "||"
│     │  _, err := doc.ReplaceWithTable(marker, tableData)
│     │
│
├─ Phase 3: Text markers [[]]
│  │
│  └─ for key, text := range data.texts
│     │  marker := "[[" + key + "]]"
│     │  _, err := doc.ReplaceText(marker, text)
│     │
│
└─ return nil
```

### 4.4 Обоснование порядка фаз

```
(()) → |||| → [[]]
```

| Порядок | Обоснование |
|---------|-------------|
| `(())` первыми | Вставленный контент может содержать `\|\|\|\|` и `[[]]` маркеры, которые должны быть обработаны в следующих фазах |
| `\|\|\|\|` вторыми | Таблицы не содержат вложенных маркеров (ячейки — plain text из `TableData.Rows`). Но табличный маркер может стоять рядом с текстовым — замена таблицы перестраивает XML, поэтому лучше сделать до текстовой замены |
| `[[]]` последними | Чистая текстовая замена, не создает новых элементов. Самая безопасная для финализации |

### 4.5 Рекурсия пользовательских маркеров

#### 4.5.1 Многопроходный алгоритм

```
Шаблон: "Начало ((A)) Конец"
A.docx:  "Часть A ((B)) продолжение"
B.docx:  "Часть B [[name]]"

Pass 0: Replace ((A)) → "Начало Часть A ((B)) продолжение Конец"
        Replace ((B)) → в этом же проходе! Результат:
        "Начало Часть A Часть B [[name]] продолжение Конец"
        totalHits = 2

Pass 1: Replace ((A)) → 0 hits (маркера нет)
        Replace ((B)) → 0 hits (маркера нет)
        totalHits = 0 → break

Phase 2: нет табличных маркеров
Phase 3: Replace [[name]] → подставляет значение

Результат: "Начало Часть A Часть B Иванов продолжение Конец"
```

#### 4.5.2 Защита от циклов

```
A.docx содержит ((B))
B.docx содержит ((A))

Pass 0: Replace ((A)) → вставляет A (с ((B)))
        Replace ((B)) → вставляет B (с ((A)))
        totalHits = 2

Pass 1: Replace ((A)) → 1 hit (вставленный из B содержал ((A)))
        Replace ((B)) → 1 hit
        totalHits = 2

... (продолжает расти экспоненциально)

Pass maxDepth-1: totalHits > 0 → return ErrMaxDepthExceeded
```

Глубина по умолчанию (10) ограничивает расширение. В реальных сценариях 2-3 уровня — максимум.

#### 4.5.3 Оптимизация: порядок обработки content-ключей

Внутри одного прохода порядок обработки content-ключей влияет на количество проходов:

```go
// Сортируем ключи для детерминированного поведения.
// Это не влияет на корректность, но делает результат предсказуемым.
keys := sortedKeys(data.content)
for _, key := range keys {
    marker := "((" + key + "))"
    n, err := doc.ReplaceWithContent(marker, ContentData{Source: data.content[key]})
    totalHits += n
}
```

Поскольку `ReplaceWithContent` сканирует весь документ, маркеры, внесенные ранней заменой в том же проходе, будут обнаружены последующими заменами. Сортировка обеспечивает детерминизм (Go map iteration order нестабилен).

### 4.6 Диаграмма взаимодействия

```
┌────────┐     ┌───────────────┐     ┌──────────────────┐
│  User  │────>│ExecuteTemplate│────>│TemplateData      │
└────────┘     └───────┬───────┘     │ .texts            │
                       │             │ .tables           │
                       │             │ .content          │
                       │             │ .options          │
                       │             └──────────────────┘
                       │
          ┌────────────┼────────────────┐
          │            │                │
          v            v                v
   ┌──────────┐ ┌────────────┐  ┌────────────┐
   │ Phase 1  │ │  Phase 2   │  │  Phase 3   │
   │ Content  │ │  Tables    │  │   Text     │
   │ (())     │ │  ||||      │  │   [[]]     │
   └────┬─────┘ └─────┬──────┘  └─────┬──────┘
        │              │               │
        v              v               v
   ┌──────────────────────────────────────────┐
   │          Document (существующий)          │
   │  .ReplaceWithContent()                    │
   │  .ReplaceWithTable()                      │
   │  .ReplaceText()                           │
   ├──────────────────────────────────────────┤
   │  BlockItemContainer.replaceTagWithElements│
   │  ResourceImporter                         │
   │  prepareContentElements                   │
   │  SplitParagraphAtTags                     │
   └──────────────────────────────────────────┘
```

### 4.7 Обработка ошибок

```go
var (
    // ErrMaxDepthExceeded возвращается когда рекурсия (()) маркеров
    // достигла максимальной глубины и необработанные маркеры остались.
    ErrMaxDepthExceeded = errors.New("docx: template content marker " +
        "recursion exceeded maximum depth")

    // ErrEmptyKey возвращается при пустом ключе маркера.
    ErrEmptyKey = errors.New("docx: template key must not be empty")

    // ErrNilSource возвращается при nil документе-источнике для (()) маркера.
    ErrNilSource = errors.New("docx: content source document must not be nil")

    // ErrMarkerNotFound возвращается в strict mode когда маркер
    // из данных не найден в документе.
    ErrMarkerNotFound = errors.New("docx: template marker not found in document")
)
```

**Стратегия ошибок:**

| Ситуация | Поведение (default) | Поведение (strict) |
|----------|--------------------|--------------------|
| Маркер в документе, нет данных | Оставить как есть | Оставить как есть (данные задает пользователь, а не документ) |
| Данные есть, маркер не найден | Игнорировать | `ErrMarkerNotFound` |
| Ошибка `ReplaceWith*` | Вернуть error | Вернуть error |
| Рекурсия > maxDepth | Остановить, вернуть nil | `ErrMaxDepthExceeded` |
| Пустой ключ | `ErrEmptyKey` | `ErrEmptyKey` |
| nil source в Content | `ErrNilSource` | `ErrNilSource` |

### 4.8 Потокобезопасность

Наследует ограничение библиотеки: **не потокобезопасен**. Один `Document` — один поток. Для параллельной обработки шаблонов: открывать отдельный `Document` из `[]byte` в каждой горутине (существующий паттерн из `replace-user-mark-batch`).

---

## 5. Детальный план реализации

### 5.1 Новые файлы

```
pkg/docx/
├── template.go              ← TemplateData + ExecuteTemplate  (~200 строк)
├── template_test.go         ← unit-тесты                     (~500 строк)
│
visual-regtest/
├── template/                ← визуальный регрессионный тест
│   ├── main.go
│   ├── builder.go
│   └── replacements.go
```

**Изменения в существующих файлах: НЕТ.**

Шаблонизатор — чистый orchestration layer поверх существующего API. Никакие существующие файлы не модифицируются.

### 5.2 Структура `template.go`

```go
package docx

import (
    "errors"
    "fmt"
    "sort"
)

// --- Константы маркеров ---

const (
    textMarkerOpen    = "[["
    textMarkerClose   = "]]"
    tableMarkerOpen   = "||"
    tableMarkerClose  = "||"
    contentMarkerOpen = "(("
    contentMarkerClose = "))"

    defaultMaxDepth = 10
)

// --- Ошибки ---

var (
    ErrMaxDepthExceeded = errors.New("docx: template content marker " +
        "recursion exceeded maximum depth")
    ErrEmptyKey   = errors.New("docx: template key must not be empty")
    ErrNilSource  = errors.New("docx: content source document must not be nil")
    ErrMarkerNotFound = errors.New("docx: template marker not found in document")
)

// --- templateOptions ---

type templateOptions struct {
    maxDepth   int
    strictMode bool
}

// --- TemplateData ---

// TemplateData содержит данные для подстановки в маркеры документа-шаблона.
type TemplateData struct {
    texts   map[string]string
    tables  map[string]TableData
    content map[string]*Document
    options templateOptions
}

// NewTemplateData создает пустой набор данных для шаблона.
func NewTemplateData() *TemplateData {
    return &TemplateData{
        texts:   make(map[string]string),
        tables:  make(map[string]TableData),
        content: make(map[string]*Document),
        options: templateOptions{maxDepth: defaultMaxDepth},
    }
}

// Text добавляет текстовую замену: [[key]] -> value.
func (td *TemplateData) Text(key, value string) *TemplateData {
    td.texts[key] = value
    return td
}

// Table добавляет табличную замену: ||key|| -> table.
func (td *TemplateData) Table(key string, data TableData) *TemplateData {
    td.tables[key] = data
    return td
}

// Content добавляет пользовательскую замену: ((key)) -> body-контент документа.
func (td *TemplateData) Content(key string, doc *Document) *TemplateData {
    td.content[key] = doc
    return td
}

// MaxDepth устанавливает максимальную глубину рекурсии для (()) маркеров.
func (td *TemplateData) MaxDepth(depth int) *TemplateData {
    td.options.maxDepth = depth
    return td
}

// StrictMode включает строгий режим.
func (td *TemplateData) StrictMode() *TemplateData {
    td.options.strictMode = true
    return td
}

// --- Валидация ---

func (td *TemplateData) validate() error {
    for key := range td.texts {
        if key == "" {
            return ErrEmptyKey
        }
    }
    for key := range td.tables {
        if key == "" {
            return ErrEmptyKey
        }
    }
    for key, doc := range td.content {
        if key == "" {
            return ErrEmptyKey
        }
        if doc == nil {
            return fmt.Errorf("%w: key %q", ErrNilSource, key)
        }
    }
    return nil
}

// --- ExecuteTemplate ---

// ExecuteTemplate подставляет данные из TemplateData в маркеры документа.
func (d *Document) ExecuteTemplate(data *TemplateData) error {
    if err := data.validate(); err != nil {
        return err
    }

    // Phase 1: Content markers (()) — многопроходная рекурсия.
    if err := d.executeContentMarkers(data); err != nil {
        return err
    }

    // Phase 2: Table markers ||.
    if err := d.executeTableMarkers(data); err != nil {
        return err
    }

    // Phase 3: Text markers [[]].
    if err := d.executeTextMarkers(data); err != nil {
        return err
    }

    return nil
}

// --- Фаза 1: Content ---

func (d *Document) executeContentMarkers(data *TemplateData) error {
    if len(data.content) == 0 {
        return nil
    }

    keys := sortedKeys(data.content)
    maxDepth := data.options.maxDepth

    for depth := 0; depth <= maxDepth; depth++ {
        totalHits := 0
        for _, key := range keys {
            marker := contentMarkerOpen + key + contentMarkerClose
            n, err := d.ReplaceWithContent(marker, ContentData{
                Source: data.content[key],
            })
            if err != nil {
                return fmt.Errorf("docx: template content %q: %w", key, err)
            }
            totalHits += n
        }
        if totalHits == 0 {
            return nil // Все маркеры обработаны.
        }
        if depth == maxDepth {
            if data.options.strictMode {
                return ErrMaxDepthExceeded
            }
            return nil // Тихо прекратить.
        }
    }
    return nil
}

// --- Фаза 2: Tables ---

func (d *Document) executeTableMarkers(data *TemplateData) error {
    keys := sortedKeys(data.tables)
    for _, key := range keys {
        marker := tableMarkerOpen + key + tableMarkerClose
        n, err := d.ReplaceWithTable(marker, data.tables[key])
        if err != nil {
            return fmt.Errorf("docx: template table %q: %w", key, err)
        }
        if n == 0 && data.options.strictMode {
            return fmt.Errorf("%w: ||%s||", ErrMarkerNotFound, key)
        }
    }
    return nil
}

// --- Фаза 3: Text ---

func (d *Document) executeTextMarkers(data *TemplateData) error {
    keys := sortedKeys(data.texts)
    for _, key := range keys {
        marker := textMarkerOpen + key + textMarkerClose
        n, err := d.ReplaceText(marker, data.texts[key])
        if err != nil {
            return fmt.Errorf("docx: template text %q: %w", key, err)
        }
        if n == 0 && data.options.strictMode {
            return fmt.Errorf("%w: [[%s]]", ErrMarkerNotFound, key)
        }
    }
    return nil
}

// --- Утилиты ---

// sortedKeys возвращает отсортированные ключи map для детерминированного порядка.
func sortedKeys[V any](m map[string]V) []string {
    keys := make([]string, 0, len(m))
    for k := range m {
        keys = append(keys, k)
    }
    sort.Strings(keys)
    return keys
}
```

### 5.3 Фазы реализации

#### Фаза 1: Ядро (~3-4 часа)

| Шаг | Описание | Файл |
|-----|----------|------|
| 1.1 | Создать `TemplateData` с builder API | `template.go` |
| 1.2 | Реализовать `validate()` | `template.go` |
| 1.3 | Реализовать `ExecuteTemplate()` с тремя фазами | `template.go` |
| 1.4 | Реализовать многопроходную рекурсию для `(())` | `template.go` |
| 1.5 | Реализовать `sortedKeys` generic helper | `template.go` |
| 1.6 | Определить sentinel-ошибки | `template.go` |

#### Фаза 2: Тесты (~4-5 часов)

| Шаг | Описание | Файл |
|-----|----------|------|
| 2.1 | Тесты текстовых маркеров `[[]]` | `template_test.go` |
| 2.2 | Тесты табличных маркеров `\|\|\|\|` | `template_test.go` |
| 2.3 | Тесты пользовательских маркеров `(())` | `template_test.go` |
| 2.4 | Тесты рекурсии: 2-3 уровня вложенности | `template_test.go` |
| 2.5 | Тесты защиты: цикл A↔B, maxDepth | `template_test.go` |
| 2.6 | Тесты strict mode | `template_test.go` |
| 2.7 | Тесты edge cases: пустые значения, маркеры в header/footer/comments/cells | `template_test.go` |
| 2.8 | Тесты TemplateData builder и validate | `template_test.go` |
| 2.9 | Тест round-trip: execute → save → open → verify | `template_test.go` |

#### Фаза 3: Визуальная регрессия (~2-3 часа)

| Шаг | Описание | Файл |
|-----|----------|------|
| 3.1 | Builder для шаблонного документа | `visual-regtest/template/builder.go` |
| 3.2 | Набор замен (тексты, таблицы, документы) | `visual-regtest/template/replacements.go` |
| 3.3 | Оркестратор | `visual-regtest/template/main.go` |
| 3.4 | Интеграция в Makefile | `visual-regtest/Makefile` |

---

## 6. Тест-план

### 6.1 Unit-тесты

```
TestExecuteTemplate_TextMarkers
├── single_marker           — [[name]] → "Иван"
├── multiple_markers        — [[first]], [[last]] → "Иван", "Петров"
├── same_marker_twice       — [[name]] встречается 2 раза
├── marker_in_header        — [[name]] в header
├── marker_in_footer        — [[name]] в footer
├── marker_in_comment       — [[name]] в комментарии
├── marker_in_table_cell    — [[name]] в ячейке таблицы
├── marker_across_runs      — [[na|me]] разбита на 2 рана
├── replace_with_empty      — [[name]] → ""
├── cyrillic_marker         — [[имя]] → "Иванов"
├── no_match_default        — нет маркера, ошибки нет
└── no_match_strict         — нет маркера, ErrMarkerNotFound

TestExecuteTemplate_TableMarkers
├── single_table            — ||items|| → 3x3 таблица
├── table_with_style        — ||items|| → таблица со стилем
├── table_in_cell           — ||items|| внутри другой таблицы
├── table_in_header         — ||items|| в header
└── multiple_tables         — ||items1||, ||items2||

TestExecuteTemplate_ContentMarkers
├── single_content          — ((header)) → содержимое header.docx
├── styles_preserved        — стили источника переносятся
├── images_preserved        — изображения из источника переносятся
├── no_header_footer        — header/footer источника не переносятся
├── numbering_preserved     — нумерация списков переносится
├── content_in_header       — ((mark)) в header документа
├── content_in_cell         — ((mark)) в ячейке таблицы
└── multiple_sources        — ((a)), ((b)) из разных документов

TestExecuteTemplate_Recursion
├── two_levels              — A содержит ((B)), B — plain text
├── three_levels            — A→B→C
├── content_with_text_marks — A содержит [[name]]
├── content_with_table_marks— A содержит ||items||
├── all_marker_types        — A содержит все три типа
├── cycle_ab                — A↔B, maxDepth=3, завершается
├── self_reference          — A содержит ((A)), maxDepth=2
├── max_depth_zero          — один проход, без рекурсии
├── max_depth_strict        — превышение, ErrMaxDepthExceeded
└── early_exit              — маркеров нет после 1-го прохода

TestExecuteTemplate_Validation
├── empty_key_text          — ErrEmptyKey
├── empty_key_table         — ErrEmptyKey
├── empty_key_content       — ErrEmptyKey
├── nil_source              — ErrNilSource
└── empty_data              — нет ошибки, документ не изменен

TestExecuteTemplate_PhaseOrder
├── content_brings_text     — ((a)) вставляет "[[name]]", Phase 3 заменяет
├── content_brings_table    — ((a)) вставляет "||items||", Phase 2 заменяет
└── content_brings_content  — ((a)) вставляет "((b))", Phase 1 pass 2 заменяет

TestTemplateData_Builder
├── chaining                — td.Text().Table().Content() fluent API
├── overwrite               — повторный Text("k", "v2") перезаписывает
└── defaults                — maxDepth=10, strict=false
```

### 6.2 Визуальная регрессия

| Сценарий | Описание |
|----------|----------|
| VR-1 | Шаблон с `[[]]` маркерами → текстовая замена |
| VR-2 | Шаблон с `\|\|\|\|` маркерами → вставка таблиц |
| VR-3 | Шаблон с `(())` маркерами → вставка документов |
| VR-4 | Шаблон со всеми тремя типами маркеров одновременно |
| VR-5 | Рекурсия: `(())` источник содержит `[[]]` маркеры |
| VR-6 | Маркеры в headers/footers |

---

## 7. Рассмотренные альтернативы

### 7.1 Regex-based парсинг маркеров vs прямая строка

**Выбрано: прямая строка.**

Regex-парсинг (`\[\[(\w+)\]\]`) потребовал бы:
- Предварительного сканирования документа для обнаружения маркеров
- Маппинга найденных ключей к данным
- Дополнительной обработки cross-run маркеров (regex работает с plain text, а маркеры могут быть разбиты на XML-раны)

Прямая строка проще и использует уже отлаженный механизм `collectTextAtoms` + `findOccurrences` в существующем коде. Ключи заранее известны из `TemplateData`, поэтому "обнаружение" маркеров не требуется.

### 7.2 Единый ResourceImporter vs per-call

**Выбрано: per-call (существующее поведение).**

Каждый `ReplaceWithContent` создает свой `ResourceImporter`. Для разных source-документов это правильно — у каждого свой набор стилей и нумерации. Объединение в один `ResourceImporter` для разных sources создало бы сложность без выигрыша.

### 7.3 Template как отдельный тип vs метод Document

**Выбрано: метод `Document.ExecuteTemplate()`.**

Альтернатива — тип `Template` с методом `Execute(doc, data)`:

```go
tmpl := docx.NewTemplate(doc)
tmpl.Execute(data)
```

Отвергнуто, потому что:
- Не добавляет ценности (Template не хранит pre-computed состояние)
- Нарушает идиому библиотеки (все операции — методы Document)
- Создает лишний тип, который нужно документировать

### 7.4 Dependency-ordered content expansion vs multi-pass

**Выбрано: multi-pass.**

Dependency ordering потребовал бы:
- Парсинга content-документов для обнаружения маркеров
- Построения DAG зависимостей
- Топологической сортировки
- Обработки циклов в графе

Multi-pass проще, корректен, и в реальных сценариях (1-3 уровня) завершается за 2-4 итерации.

### 7.5 Сканирование документа для strict mode

**Рассмотрено и отложено** на будущее.

Полноценная проверка "маркер есть в документе, но данных нет" потребовала бы:
- Парсинга всего текста документа
- Regex-поиска маркеров по паттернам `\[\[.*?\]\]`, `\|\|.*?\|\|`, `\(\(.*?\)\)`
- Сравнения с ключами из `TemplateData`

Это добавляет сложность и может давать false positives (текст `[[comment]]` в обычном тексте). В текущей версии strict mode проверяет только "данные есть, маркер не найден" (обратная проверка).

---

## 8. Граничные случаи

| Случай | Поведение |
|--------|----------|
| Один ключ для разных типов маркеров: `Text("x", ...)` + `Table("x", ...)` | Допустимо — разные разделители (`[[x]]` vs `\|\|x\|\|`), конфликта нет |
| Маркер `[[]]` (пустой ключ) | `ErrEmptyKey` при валидации |
| Маркер `[[foo` (незакрытый) | Не является валидным маркером, игнорируется |
| Пустой `TemplateData` (без данных) | `ExecuteTemplate` возвращает nil, документ не изменен |
| Source document содержит `(())` маркер для себя | Расширяется до maxDepth, затем останавливается |
| `TableData` с пустым `Rows` | Ошибка из существующего `ReplaceWithTable` (валидация в `buildTableElement`) |
| Маркер `\|\|foo\|\|` где `foo` — ключ | Корректно: `\|\|foo\|\|` парсится как marker open `\|\|` + key `foo` + marker close `\|\|` |
| Маркеры вложены: `[[tab\|\|le]]` | Каждый тип маркеров ищется как целая строка; вложенные разделители — просто символы текста |

---

## 9. Расширяемость (будущее)

Текущий дизайн оставляет точки расширения для будущих версий:

### 9.1 Условные блоки (v2)

```go
td.Condition("show_discount", true)
// В документе: {{#show_discount}} ... {{/show_discount}}
```

Потребует нового типа маркеров и логики удаления/сохранения блоков XML-элементов.

### 9.2 Циклы (v2)

```go
td.Repeat("items", []map[string]string{
    {"name": "Item 1", "price": "100"},
    {"name": "Item 2", "price": "200"},
})
```

Потребует клонирования секции документа для каждой итерации.

### 9.3 Пользовательские разделители (v1.1)

```go
td.SetDelimiters("${", "}", "#{", "}", "#{", "}") // text, table, content
```

Тривиальная замена констант на параметры в `templateOptions`.

### 9.4 Callback-маркеры (v2)

```go
td.TextFunc("today", func() string {
    return time.Now().Format("02.01.2006")
})
```

Ленивое вычисление значений.

---

## 10. Метрики успешности

| Метрика | Цель |
|---------|------|
| Новый код | < 250 строк продакшн-кода |
| Тесты | > 30 тест-кейсов |
| Покрытие | > 90% для `template.go` |
| Изменения в существующих файлах | 0 |
| Новые зависимости | 0 |
| Тесты проходят | `go test ./... ` — 0 failures |
| Визуальная регрессия | SSIM > 0.99 для всех сценариев |

---

## 11. Риски и митигация

| Риск | Вероятность | Влияние | Митигация |
|------|------------|---------|-----------|
| Маркеры разделителей `\|\|` конфликтуют с таблицами Markdown в тексте документа | Низкая | Низкое | Пользователь выбирает ключи; v1.1 — кастомные разделители |
| Экспоненциальный рост при циклических `(())` | Средняя | Высокое | maxDepth=10 по умолчанию; документировать |
| Потеря стилей при многоуровневой `(())` рекурсии | Низкая | Среднее | Каждый `ReplaceWithContent` запускает полный `ResourceImporter` |
| Performance на больших шаблонах с множеством маркеров | Низкая | Среднее | Каждый `Replace*` — линейный скан; N маркеров × M элементов = O(NM). Приемлемо для типичных документов |
| Порядок обработки ключей внутри фазы влияет на результат | Средняя | Низкое | Сортировка ключей обеспечивает детерминизм |
