# Аудит плана: Шаблонизатор документов go-docx

**Дата:** 2026-03-07
**Цель:** Проверка плана на enterprise-надежность реализации

---

## Вердикт

План архитектурно здоров: тонкий оркестрационный слой поверх отлаженных Replace*-примитивов — правильный подход. Однако обнаружены **4 критических**, **6 серьезных** и **5 умеренных** проблем, которые необходимо устранить до начала реализации.

---

## КРИТИЧЕСКИЕ ПРОБЛЕМЫ (блокируют production)

### K1. Экспоненциальный рост документа при циклических (()) маркерах — нет защиты по памяти

**Суть:** Цикл A↔B с `maxDepth=10` создает 2^10 = 1024 копии контента. При 100 КБ контента — 100 МБ. При `maxDepth=10` (default) это гарантированный OOM в enterprise-сценариях.

**Текущий план:** Только `maxDepth` лимит. Этого недостаточно.

**Решение:**

```go
type templateOptions struct {
    maxDepth         int   // default: 10
    maxTotalInserts  int   // default: 1000 — лимит на общее число замен (())
    strictMode       bool
}
```

В `executeContentMarkers` добавить счетчик:

```go
grandTotal := 0
for depth := 0; depth <= maxDepth; depth++ {
    // ...
    grandTotal += totalHits
    if grandTotal > data.options.maxTotalInserts {
        return ErrInsertLimitExceeded
    }
}
```

Это защищает от экспоненциального роста независимо от глубины. `MaxTotalInserts(1000)` — builder-метод.

---

### K2. NFR-3 ложное — ResourceImporter НЕ единый

**Суть:** NFR-3 заявляет: *"Единый ResourceImporter для всех (()) замен в одном вызове"*. Это неверно. Каждый `ReplaceWithContent` создает **новый** `ResourceImporter` (строка `document.go:569`). В шаблонизаторе N content-ключей × D проходов = N×D независимых ResourceImporter-ов.

**Влияние:** Для разных source-документов это корректное поведение (у каждого свой набор стилей). Но NFR-3 вводит в заблуждение и может привести к ложным ожиданиям при реализации.

**Решение:** Удалить NFR-3. Заменить на:

> NFR-3: Каждый `ReplaceWithContent` создает собственный `ResourceImporter`, привязанный к конкретному source-документу. Это обеспечивает корректный импорт стилей/нумерации из каждого source.

---

### K3. NFR-2 ложное — обнаружения циклов нет

**Суть:** NFR-2 заявляет: *"Обнаружение цикла A→B→A на уровне одного прохода"*. Реализация этого не делает — только `maxDepth` лимит. Цикл A↔B **не обнаруживается**, просто прерывается по лимиту. Это разные вещи.

**Влияние:** Без обнаружения цикла пользователь не получит осмысленную ошибку. `ErrMaxDepthExceeded` не объясняет **почему** — переполнение глубины из-за цикла или из-за легитимно глубокой иерархии.

**Решение:** Удалить NFR-2. Заменить на:

> NFR-2: Защита от бесконечной рекурсии через двойной лимит: maxDepth (вложенность) + maxTotalInserts (абсолютный объем). При срабатывании — детальная ошибка с указанием текущей глубины и количества замен.

```go
var ErrMaxDepthExceeded = errors.New(
    "docx: template content marker recursion exceeded maximum depth; " +
    "possible circular reference between content sources")
```

---

### K4. Ключи с символами разделителей создают невалидные маркеры

**Суть:** Нет валидации ключей на содержание символов разделителей. Ключ `"a]]b"` создает маркер `"[[a]]b]]"`, который будет обнаружен как `"[[a]]"` + мусор `"b]]"`. Ключ `"a))inner"` создает `"((a))inner))"` — аналогично.

**Влияние:** Тихая ошибка — маркер не будет найден или будет найден частично. Отладка невозможна без знания внутренней механики.

**Решение:** Добавить валидацию в `validate()`:

```go
func validateKey(key, openDelim, closeDelim string) error {
    if key == "" {
        return ErrEmptyKey
    }
    if strings.Contains(key, openDelim) || strings.Contains(key, closeDelim) {
        return fmt.Errorf("%w: key %q contains delimiter characters %q/%q",
            ErrInvalidKey, key, openDelim, closeDelim)
    }
    return nil
}
```

Вызывать для каждого типа маркеров с соответствующими разделителями.

---

## СЕРЬЕЗНЫЕ ПРОБЛЕМЫ (влияют на enterprise-качество)

### С1. Нет `context.Context` — невозможно отменить долгую операцию

**Суть:** `ExecuteTemplate` может работать долго (рекурсивные (()) + большие документы). Enterprise-приложения требуют timeout и cancellation. Без `context.Context` вызывающий код не может прервать операцию.

**Текущий план:** Отсутствует.

**Решение:** Добавить вариант с контекстом:

```go
func (d *Document) ExecuteTemplate(data *TemplateData) error {
    return d.ExecuteTemplateWithContext(context.Background(), data)
}

func (d *Document) ExecuteTemplateWithContext(ctx context.Context, data *TemplateData) error {
    // Проверка ctx.Err() между фазами и между итерациями depth-loop
}
```

Это не требует изменений в нижележащих `Replace*` — достаточно проверять `ctx.Err()` в ключевых точках оркестрации (между вызовами Replace*, между проходами).

---

### С2. Нет результата выполнения — невозможен аудит

**Суть:** `ExecuteTemplate` возвращает только `error`. Enterprise-системы требуют аудит: сколько маркеров обработано, сколько осталось, какая глубина рекурсии использована.

**Текущий план:** Отсутствует.

**Решение:**

```go
// TemplateResult содержит статистику выполнения шаблона.
type TemplateResult struct {
    TextReplacements    int            // Количество текстовых замен
    TableReplacements   int            // Количество табличных замен
    ContentReplacements int            // Количество контентных замен
    ContentDepth        int            // Фактическая глубина рекурсии (())
    Unresolved          []string       // Оставшиеся маркеры (только в strict mode)
}

func (d *Document) ExecuteTemplate(data *TemplateData) (*TemplateResult, error)
```

Это ломает сигнатуру из плана, но соответствует паттерну библиотеки (все Replace* возвращают `(int, error)`).

---

### С3. Двойное оборачивание ошибок

**Суть:** План оборачивает ошибки как `fmt.Errorf("docx: template content %q: %w", key, err)`. Но `ReplaceWithContent` внутри уже оборачивает с `"docx: ..."`. Результат:

```
docx: template content "header": docx: importing styles: docx: ...
```

Тройной префикс `docx:` — нечитаемо.

**Решение:** Использовать контекст без повторного префикса:

```go
return fmt.Errorf("template content %q: %w", key, err)
```

Или ввести `TemplateError` тип, оборачивающий цепочку с чистым форматированием:

```go
type TemplateError struct {
    Phase   string // "content", "table", "text"
    Key     string
    Cause   error
}

func (e *TemplateError) Error() string {
    return fmt.Sprintf("docx: template %s %q: %s", e.Phase, e.Key, e.Cause)
}

func (e *TemplateError) Unwrap() error { return e.Cause }
```

---

### С4. Тип ошибок несовместим с существующей иерархией

**Суть:** Библиотека использует `DocxError` с `Unwrap()` (см. `errors.go`). План использует `errors.New()` для sentinel-ошибок. Это нарушает единообразие и не поддерживает `errors.As(*DocxError)`.

**Решение:** Для enterprise consistency определить ошибки шаблонизатора через существующий `DocxError`:

```go
var (
    ErrMaxDepthExceeded = NewDocxError(nil,
        "docx: template content marker recursion exceeded maximum depth")
    ErrEmptyKey = NewDocxError(nil,
        "docx: template key must not be empty")
    // ...
)
```

Или ввести `TemplateError` как подтип `DocxError` (аналогично `InvalidXmlError`, `InvalidSpanError`).

---

### С5. `||` как разделитель — высокий риск коллизий

**Суть:** `||` — одна из самых часто встречающихся комбинаций символов:
- Юридические документы: `"параграф || пункт"`
- Таблицы в тексте (Markdown-стиль)
- Программисты: `"a || b"` (logical OR)
- Формулы: `||x||` (норма вектора)

Если пользователь поставит `||total||` как маркер в документе, но где-то в другом месте текста случайно окажется `||total||` (например в примере кода) — ложное срабатывание.

**Решение:** Использовать более уникальные разделители по умолчанию:

```go
const (
    textMarkerOpen    = "[["
    textMarkerClose   = "]]"
    tableMarkerOpen   = "{{tbl:"    // или: "|["
    tableMarkerClose  = "}}"        // или: "]|"
    contentMarkerOpen = "(("
    contentMarkerClose = "))"
)
```

Или **сразу реализовать кастомные разделители** (отложено в плане на v1.1, но для enterprise это day-one требование):

```go
func (td *TemplateData) Delimiters(
    textOpen, textClose,
    tableOpen, tableClose,
    contentOpen, contentClose string,
) *TemplateData
```

---

### С6. Нет атомарности — частичная модификация без возможности отката

**Суть:** Если Phase 1 (content) успешна, Phase 2 (tables) падает — документ содержит вставленный контент, но без таблиц. Enterprise-системы требуют хотя бы информирования о том, на какой фазе произошел сбой.

**Текущий план:** Упоминает проблему ("документ может быть частично модифицирован"), но не предлагает решения.

**Решение:** Полная атомарность невозможна без снапшотов XML-дерева (дорого). Прагматичный подход:

1. Включить в `TemplateResult` поле `CompletedPhase`:

```go
type TemplateResult struct {
    // ...
    CompletedPhase int // 0=none, 1=content, 2=tables, 3=text (done)
}
```

2. Документировать в godoc: `ExecuteTemplate` не является атомарной операцией. При ошибке `result.CompletedPhase` указывает последнюю завершенную фазу.

3. Рекомендовать пользователю: работать с копией документа (OpenBytes → modify → SaveToBytes).

---

## УМЕРЕННЫЕ ПРОБЛЕМЫ (улучшения для зрелости)

### У1. Нет бенчмарков в тест-плане

**Суть:** Enterprise-библиотека требует baseline performance данных. В тест-плане только функциональные тесты.

**Решение:** Добавить:

```go
BenchmarkExecuteTemplate_TextOnly      // N текстовых маркеров
BenchmarkExecuteTemplate_TableOnly     // N табличных маркеров
BenchmarkExecuteTemplate_ContentOnly   // N контентных маркеров
BenchmarkExecuteTemplate_Mixed         // все типы
BenchmarkExecuteTemplate_Recursion2    // 2 уровня (())
BenchmarkExecuteTemplate_Recursion5    // 5 уровней (())
BenchmarkExecuteTemplate_LargeDoc      // 100+ страниц шаблон
```

---

### У2. Lifecycle source-документов не документирован

**Суть:** Все content-документы (`*Document`) должны оставаться открытыми на время всего `ExecuteTemplate`. При maxDepth=10 и рекурсивных (()) это может быть значительное время. Для enterprise с десятками content-документов — значительное потребление памяти.

**Решение:** Добавить в godoc `Content()`:

```go
// Content добавляет пользовательскую замену: ((key)) → body-контент документа.
// Документ-источник doc должен оставаться открытым до завершения ExecuteTemplate.
// После завершения doc можно безопасно закрыть или переиспользовать.
```

И в `TemplateData` добавить документацию о lifecycle.

---

### У3. MaxDepth(0) — неоднозначная семантика

**Суть:** Документация говорит: *"Значение 0 запрещает рекурсию (один проход)"*. Но "один проход" — это не "запрещает", а "один раз". Пользователь может ожидать, что 0 = пропустить фазу content вообще.

**Решение:** Уточнить:

```go
// MaxDepth устанавливает максимальное количество проходов для рекурсивного
// раскрытия (()) маркеров.
//   - 0: один проход (маркеры раскрываются, но вставленный контент не сканируется повторно)
//   - 1: два прохода (контент первого уровня проверяется на вложенные маркеры)
//   - N: N+1 проходов
//   - -1: пропустить content-фазу полностью
```

---

### У4. Нет теста на concurrent safety documentation

**Суть:** Документация заявляет "не потокобезопасен", но нет теста с `-race`, который бы верифицировал, что concurrent вызовы на РАЗНЫХ Document-ах безопасны (а на одном — нет).

**Решение:** Добавить тест:

```go
func TestExecuteTemplate_ConcurrentDifferentDocuments(t *testing.T) {
    // Параллельное выполнение на разных Document-ах должно быть безопасно
    // go test -race
}
```

---

### У5. Нет валидации TableData при добавлении в TemplateData

**Суть:** `td.Table("key", TableData{Rows: nil})` — невалидный TableData принимается без ошибки. Ошибка возникнет только при `ExecuteTemplate` из глубины `buildTableElement`. Стек ошибки будет неочевиден.

**Решение:** Валидировать `TableData` в `validate()`:

```go
for key, td := range td.tables {
    if len(td.Rows) == 0 {
        return fmt.Errorf("%w: table %q has no rows", ErrInvalidTableData, key)
    }
    cols := len(td.Rows[0])
    for i, row := range td.Rows {
        if len(row) != cols {
            return fmt.Errorf("%w: table %q row %d has %d columns, expected %d",
                ErrInvalidTableData, key, i, len(row), cols)
        }
    }
}
```

---

## КОРРЕКТНЫЕ РЕШЕНИЯ В ПЛАНЕ

Для баланса — решения, которые проверены и правильны:

| Решение | Обоснование |
|---------|-------------|
| Порядок фаз `(()) → \|\|\|\| → [[]]` | Корректен: content может нести вложенные маркеры других типов |
| Multi-pass вместо dependency ordering | Корректен: проще, корректен, достаточен для 1-3 уровней |
| Метод Document вместо отдельного типа Template | Корректен: соответствует идиоме библиотеки |
| Sorted keys для детерминизма | Корректен: Go map iteration нестабилен |
| Deep-copy source elements | Подтверждено кодом: `prepareContentElements` делает `el.Copy()` на строке 289 |
| Отсутствие изменений в существующих файлах | Подтверждено: все необходимое API уже публично |
| Прямая строка вместо regex | Корректен: `collectTextAtoms` + `findOccurrences` уже обрабатывает cross-run |
| Per-call ResourceImporter | Подтверждено кодом: `document.go:569` создает новый для каждого source |
| Early exit при totalHits=0 | Корректен: оптимизирует типичный сценарий (1-2 прохода) |
| Strict mode как opt-in | Корректен: по умолчанию пропускает неизвестные маркеры |

---

## СВОДНАЯ ТАБЛИЦА

| # | Тип | Проблема | Приоритет |
|---|-----|----------|-----------|
| K1 | Критическая | Нет лимита на total inserts → OOM при циклах | P0 |
| K2 | Критическая | NFR-3 ложное (ResourceImporter не единый) | P0 |
| K3 | Критическая | NFR-2 ложное (цикло-детекция отсутствует) | P0 |
| K4 | Критическая | Нет валидации ключей на символы разделителей | P0 |
| С1 | Серьезная | Нет context.Context | P1 |
| С2 | Серьезная | Нет TemplateResult (аудит, статистика) | P1 |
| С3 | Серьезная | Двойное оборачивание ошибок docx:docx: | P1 |
| С4 | Серьезная | Sentinel errors вместо DocxError иерархии | P1 |
| С5 | Серьезная | `\|\|` разделитель — высокий риск коллизий | P1 |
| С6 | Серьезная | Нет информации о фазе при partial failure | P1 |
| У1 | Умеренная | Нет бенчмарков | P2 |
| У2 | Умеренная | Lifecycle source-документов не документирован | P2 |
| У3 | Умеренная | MaxDepth(0) неоднозначная семантика | P2 |
| У4 | Умеренная | Нет race-condition теста | P2 |
| У5 | Умеренная | Нет ранней валидации TableData | P2 |

---

## ИСПРАВЛЕННАЯ АРХИТЕКТУРА

С учетом всех замечаний, ключевые изменения в API:

### Исправленная сигнатура

```go
// ExecuteTemplate подставляет данные из TemplateData в маркеры документа.
// Возвращает статистику выполнения и ошибку (если произошла).
// При ошибке result.CompletedPhase указывает последнюю успешную фазу.
func (d *Document) ExecuteTemplate(data *TemplateData) (*TemplateResult, error)

// ExecuteTemplateWithContext — вариант с поддержкой отмены.
func (d *Document) ExecuteTemplateWithContext(
    ctx context.Context, data *TemplateData,
) (*TemplateResult, error)
```

### Исправленный TemplateData

```go
type TemplateData struct {
    texts   map[string]string
    tables  map[string]TableData
    content map[string]*Document
    options templateOptions
}

type templateOptions struct {
    maxDepth        int  // default: 10
    maxTotalInserts int  // default: 1000, защита от экспоненциального роста
    strictMode      bool
    // Кастомные разделители (nil = defaults).
    delimiters      *markerDelimiters
}

type markerDelimiters struct {
    textOpen, textClose       string
    tableOpen, tableClose     string
    contentOpen, contentClose string
}
```

### Исправленный TemplateResult

```go
type TemplateResult struct {
    TextReplacements    int
    TableReplacements   int
    ContentReplacements int
    ContentDepth        int // фактическая глубина рекурсии
    CompletedPhase      int // 0=none, 1=content, 2=tables, 3=text
}
```

### Исправленная валидация

```go
func (td *TemplateData) validate() error {
    d := td.effectiveDelimiters()
    for key := range td.texts {
        if err := validateKey(key, d.textOpen, d.textClose); err != nil {
            return err
        }
    }
    for key, tbl := range td.tables {
        if err := validateKey(key, d.tableOpen, d.tableClose); err != nil {
            return err
        }
        if err := validateTableData(key, tbl); err != nil {
            return err
        }
    }
    for key, doc := range td.content {
        if err := validateKey(key, d.contentOpen, d.contentClose); err != nil {
            return err
        }
        if doc == nil {
            return fmt.Errorf("%w: key %q", ErrNilSource, key)
        }
    }
    return nil
}
```

### Исправленный executeContentMarkers

```go
func (d *Document) executeContentMarkers(
    ctx context.Context, data *TemplateData, result *TemplateResult,
) error {
    if len(data.content) == 0 {
        return nil
    }

    delim := data.effectiveDelimiters()
    keys := sortedKeys(data.content)
    maxDepth := data.options.maxDepth
    maxInserts := data.options.maxTotalInserts

    for depth := 0; depth <= maxDepth; depth++ {
        // Проверка отмены.
        if err := ctx.Err(); err != nil {
            return fmt.Errorf("template cancelled at content depth %d: %w", depth, err)
        }

        passHits := 0
        for _, key := range keys {
            marker := delim.contentOpen + key + delim.contentClose
            n, err := d.ReplaceWithContent(marker, ContentData{
                Source: data.content[key],
            })
            if err != nil {
                return &TemplateError{Phase: "content", Key: key, Cause: err}
            }
            passHits += n
        }
        if passHits == 0 {
            result.ContentDepth = depth
            return nil
        }

        result.ContentReplacements += passHits
        result.ContentDepth = depth + 1

        // Защита от экспоненциального роста.
        if result.ContentReplacements > maxInserts {
            return fmt.Errorf("%w: %d inserts exceeded limit %d "+
                "(possible circular reference)",
                ErrInsertLimitExceeded, result.ContentReplacements, maxInserts)
        }

        if depth == maxDepth {
            if data.options.strictMode {
                return ErrMaxDepthExceeded
            }
            return nil
        }
    }
    return nil
}
```

---

## ИСПРАВЛЕННЫЕ ТРЕБОВАНИЯ

### NFR (замена)

| ID | Было | Стало |
|----|------|-------|
| NFR-2 | Обнаружение цикла A→B→A на уровне одного прохода | Защита от бесконечной рекурсии через двойной лимит: maxDepth + maxTotalInserts |
| NFR-3 | Единый ResourceImporter для всех (()) замен | Каждый ReplaceWithContent создает собственный ResourceImporter для своего source-документа |

### NFR (новые)

| ID | Требование |
|----|-----------|
| NFR-7 | Поддержка context.Context для отмены и timeout |
| NFR-8 | Возврат TemplateResult со статистикой выполнения |
| NFR-9 | Валидация ключей на содержание символов разделителей |
| NFR-10 | Лимит на общее количество content-вставок (maxTotalInserts, default 1000) |
| NFR-11 | Ранняя валидация TableData при вызове validate() |

---

## ИСПРАВЛЕННЫЙ ТЕСТ-ПЛАН (дополнения)

```
TestExecuteTemplate_Protection
├── cycle_ab_totalInsertLimit     — A↔B, maxTotalInserts=5, ErrInsertLimitExceeded
├── exponential_growth_stopped    — self-ref ((A))→A, maxTotalInserts=50
├── context_cancelled             — ctx с timeout, отмена между проходами
├── context_cancelled_mid_phase   — отмена между ключами внутри одной фазы

TestExecuteTemplate_KeyValidation
├── key_contains_open_delim       — key "a[[b" → ErrInvalidKey
├── key_contains_close_delim      — key "a]]b" → ErrInvalidKey
├── key_contains_table_delim      — key "a||b" → ErrInvalidKey
├── key_contains_content_delim    — key "a))b" → ErrInvalidKey

TestExecuteTemplate_Result
├── result_counts                 — TextReplacements, TableReplacements, ContentReplacements
├── result_depth                  — ContentDepth отражает фактическую глубину
├── result_completed_phase_success— CompletedPhase = 3 при успехе
├── result_completed_phase_error  — CompletedPhase = 1 при ошибке в Phase 2

TestExecuteTemplate_TableDataValidation
├── empty_rows                    — ErrInvalidTableData при validate()
├── inconsistent_columns          — ErrInvalidTableData при validate()

BenchmarkExecuteTemplate_TextOnly_10
BenchmarkExecuteTemplate_TextOnly_100
BenchmarkExecuteTemplate_ContentRecursion_Depth3
BenchmarkExecuteTemplate_LargeDocument
BenchmarkExecuteTemplate_ManyMarkers

TestExecuteTemplate_ConcurrentDifferentDocuments
├── parallel_execution            — 10 горутин, каждая свой Document, -race
```

---

## ЗАКЛЮЧЕНИЕ

План архитектурно верен в основе. Ключевое решение (оркестрация поверх Replace*-примитивов) — правильное и минимально инвазивное. После устранения 4 критических и 6 серьезных проблем, план готов к enterprise-реализации.

**Основные действия перед реализацией:**
1. Добавить `maxTotalInserts` лимит (K1)
2. Исправить NFR-2, NFR-3 на корректные формулировки (K2, K3)
3. Добавить валидацию ключей на разделители (K4)
4. Добавить `context.Context` (С1)
5. Добавить `TemplateResult` (С2)
6. Решить вопрос с `||` разделителем (С5) — либо другой default, либо кастомные разделители сразу
7. Интегрировать ошибки в существующую `DocxError` иерархию (С4)
