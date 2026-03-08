# Code Review: go-docx

**Дата:** 2026-03-07
**Проект:** go-docx — библиотека для создания, чтения и модификации Microsoft Word (.docx) документов на Go
**Объем:** ~62 000 строк Go-кода, 234 файла (.go), 90 тестовых файлов
**Зависимости:** `github.com/beevik/etree`, `gopkg.in/yaml.v3` (минималистично)

---

## 1. Архитектура

### 1.1 Слоистая архитектура (оценка: 9/10)

Проект реализует чистую **4-слойную архитектуру**, зеркалирующую python-docx:

```
┌──────────────────────────────────────────────────┐
│  Domain Layer (pkg/docx/)                        │
│  Document, Paragraph, Run, Table, Section, etc.  │
├──────────────────────────────────────────────────┤
│  Parts Layer (pkg/docx/parts/)                   │
│  DocumentPart, StylesPart, NumberingPart, etc.   │
├──────────────────────────────────────────────────┤
│  OXML Layer (pkg/docx/oxml/)                     │
│  CT_* типы — generated (zz_gen_*) + custom       │
├──────────────────────────────────────────────────┤
│  OPC Layer (pkg/docx/opc/)                       │
│  ZIP I/O, relationships, формат-агностичный      │
└──────────────────────────────────────────────────┘
```

**Зависимости строго однонаправленные** — каждый слой зависит только от нижележащего:

```
docx → parts → oxml → opc → stdlib + etree
  ↓       ↓       ↓
 enum   enum    enum
image
```

**Циклических зависимостей не обнаружено.** OPC-слой полностью формат-агностичен и теоретически пригоден для xlsx/pptx.

### 1.2 Пакетная структура

| Пакет | Файлов | Назначение | Оценка |
|-------|--------|------------|--------|
| `pkg/docx/` | ~128 | Domain API: Document, Paragraph, Table, Run... | Хорошо |
| `pkg/docx/oxml/` | ~64 | XML-структуры (CT_* типы) | Отлично |
| `pkg/docx/opc/` | ~22 | Open Packaging Convention | Отлично |
| `pkg/docx/parts/` | ~19 | Document parts (Story, Styles, Numbering...) | Отлично |
| `pkg/docx/enum/` | 8 | Перечисления с bidirectional XML-маппингом | Отлично |
| `pkg/docx/image/` | 14 | Парсинг изображений (JPEG, PNG, GIF, BMP, TIFF) | Отлично |
| `pkg/docx/templates/` | — | Embed-шаблоны (default.docx, XML) | Хорошо |
| `internal/codegen/` | 5 | Кодогенерация из YAML-схем | Отлично |
| `schema/` | 16 | YAML-определения элементов OOXML | Отлично |
| `visual-regtest/` | ~20 | Визуальное регрессионное тестирование | Отлично |

### 1.3 Кодогенерация (оценка: 10/10)

Проект использует **двухуровневую стратегию** для OXML-типов:

- **Генерируемый уровень** (`zz_gen_*.go`): boilerplate GetOrAdd/Add/Remove/List-методы из YAML-схем
- **Кастомный уровень** (`*_custom.go`): hand-written бизнес-логика

Разделение чистое — генерируемый код никогда не конфликтует с кастомным. Схемы в `/schema/*.yaml` — единый источник правды для XML-структуры.

### 1.4 Иерархия Parts

```
BasePart
├─ ImagePart (binary, non-XML)
└─ XmlPart
   ├─ StylesPart
   ├─ SettingsPart
   ├─ NumberingPart
   ├─ CorePropertiesPart
   └─ StoryPart (base для body-like контента)
      ├─ DocumentPart
      ├─ HeaderPart / FooterPart
      ├─ CommentsPart
      └─ FootnotesPart / EndnotesPart
```

Паттерн наследования через композицию — идиоматичный Go.

---

## 2. Качество кода

### 2.1 Общая оценка: 7.5/10

Код зрелый, продакшн-ready. Основные находки — рефакторинг-рекомендации, не баги.

### 2.2 Положительные паттерны

**Type Safety через sealed-интерфейсы:**
```go
type StyleRef interface {
    isStyleRef() // sealed — запрещает внешние реализации
}
```
Используется для: `StyleRef`, `UnderlineVal`, `LineSpacingVal`, `InlineItem`.

**Options-паттерн для API:**
```go
func (d *Document) AddParagraph(text string, style ...StyleRef)
```

**Итеративный DFS вместо рекурсии:**
Все обходы дерева элементов используют явный стек, избегая stack overflow на глубоко вложенных документах. В `table.go:398-444` есть комментарий о рефакторинге рекурсии в итерацию — отличная практика.

**Идемпотентный импорт ресурсов:**
ResourceImporter использует done-флаги для всех фаз (styles, numbering, footnotes, endnotes), что делает вызовы безопасными для повторения.

**Дедупликация изображений через SHA-256:**
`WmlPackage.GetOrAddImagePart()` предотвращает дублирование при вставке одинаковых изображений.

**Нет panic() в библиотечном коде** (только в init()-инициализации namespace-маппингов).

**Нет TODO/FIXME/HACK комментариев** — код чист от технического долга.

### 2.3 Проблемы: God Object

**`document.go` (667 строк, 29+ публичных методов)** — приближается к God Object:

| Ответственность | Примеры методов |
|----------------|-----------------|
| Добавление контента | `AddParagraph`, `AddTable`, `AddHeading`, `AddPageBreak`, `AddPicture` |
| Получение контента | `Paragraphs`, `Tables`, `IterInnerContent`, `InlineShapes`, `Comments` |
| Замена текста | `ReplaceText`, `ReplaceWithTable`, `ReplaceWithContent` |
| Свойства | `Sections`, `Styles`, `Settings`, `CoreProperties` |
| Персистенция | `Save`, `SaveFile` |

Нарушает Single Responsibility Principle — можно декомпозировать на отдельные фасады (Builder, Searcher, Replacer), но для библиотеки-порта python-docx это осознанное решение в пользу API-совместимости.

### 2.4 Проблемы: Большие файлы с несколькими ответственностями

| Файл | Строк | Проблема |
|------|-------|----------|
| `oxml/table_custom.go` | 1210 | Смешаны: создание таблиц, cell merging, свойства ячеек |
| `section.go` | 722 | 5 концепций: Section, Sections, baseHeaderFooter, Header, Footer |
| `document.go` | 667 | God Object (см. выше) |
| `blkcntnr.go` | 500+ | Базовый класс + движок замены + splice-логика |

**Рекомендация:** `section.go` разбить на `section.go` + `header_footer.go`. `table_custom.go` — на `table_custom.go` + `table_merge_custom.go`.

### 2.5 Проблемы: Дублирование кода

**1. Итерация по дочерним элементам параграфа (4 места):**

В `blkcntnr.go` (IterInnerContent, Paragraphs, Tables) и `section.go` (collectBlockItems) используется идентичный паттерн:

```go
for _, child := range c.element.ChildElements() {
    if child.Space == "w" && child.Tag == "p" {
        p := &oxml.CT_P{Element: oxml.WrapElement(child)}
        result = append(result, newParagraph(p, c.part))
    }
}
```

Можно вынести в общий итератор.

**2. Header/Footer dedup-логика (3 места):**

В `section.go` три метода (`replaceTextDedup`, `replaceWithTableDedup`, `replaceWithContentDedup`) реализуют почти идентичный паттерн дедупликации по `relId`.

**3. Error wrapping** — `fmt.Errorf("docx: %s: %w", context, err)` повторяется 40+ раз без централизованного хелпера.

### 2.6 Проблемы: Magic Numbers

| Значение | Где | Проблема |
|----------|-----|----------|
| `"04A0"` | `oxml/table_custom.go:36` | Хардкод table look без комментария |
| `914400` | `shape_custom.go:29,107` | Хардкод 1 inch в EMU без ссылки на константу |
| `"continue"` | `text_paragraph_custom.go:146` | Строковый литерал для vMerge — вместо константы |

Константы `EmuPerInch`, `EmuPerPt`, `EmuPerTwip` определены в `shared.go`, но не всегда используются — встречаются сырые значения.

### 2.7 Проблемы: Нейминг

Несколько сокращений снижают читаемость:

| Имя | Что значит | Лучше |
|-----|-----------|-------|
| `bic` | BlockItemContainer | `container` |
| `hdrFtrOps` | Header/Footer Operations | `HeaderFooterOps` |
| `rPrOwner` | Run Properties Owner | `RunPropertiesOwner` |
| `sr.sectPrEl` | Section Properties Element | `sectionPropsElement` |

Префикс `CT_*` в oxml — наследие OOXML-спецификации (Complex Type), оправдан контекстом.

---

## 3. Обработка ошибок (оценка: 9/10)

### Сильные стороны

- **Кастомные типы ошибок:** `DocxError`, `InvalidXmlError`, `PackageNotFoundError`, `InvalidSpanError`
- **Консистентное оборачивание:** `fmt.Errorf("context: %w", err)` по всему коду
- **Единые префиксы:** `"docx:"`, `"parts:"`, `"opc:"` — облегчают трассировку
- **errors.Is / errors.As** работают корректно (есть тесты в `domain_test.go`)

### Незначительные замечания

- `_, _ = rand.Read(b)` в `resourceimport_num.go:213` — безопасно, документировано комментарием "infallible since Go 1.22"
- Нет подавленных ошибок в продакшн-коде

---

## 4. Тестирование (оценка: 8.5/10)

### 4.1 Покрытие

- **90 тестовых файлов** по всем слоям (docx, parts, oxml, opc, image, enum, codegen)
- **324 использования** `t.Parallel()` — тесты параллелизированы
- **3 бенчмарка** (OpenBytes, SaveToBytes, RoundTrip)
- Соотношение тест-файлов к продакшн-файлам: ~38%

### 4.2 Паттерны тестирования

**Table-driven tests** — используются массово и качественно:
```go
tests := []struct {
    name     string
    innerXml string
    expected string
}{
    {"empty_p", ``, ""},
    {"simple_text", `<w:r><w:t>foo</w:t></w:r>`, "foo"},
}
for _, tt := range tests {
    t.Run(tt.name, func(t *testing.T) { ... })
}
```

**Тестовые хелперы** хорошо организованы в `testutil_test.go` и `domain_test.go`:
- `mustParseXml()`, `makeP()`, `makeR()`, `makeTbl()` — фабрики элементов
- `compareBoolPtr()` — nil-safe сравнение
- Все помечены `t.Helper()`

**Round-trip тесты:**
- `TestRoundTrip_DefaultDocx` — Open → Save → Re-open
- `TestIntegration_ModifyAndRoundTrip` — Open → Modify → Save → Re-open → Verify
- `TestIntegration_BlobContentPreserved` — бинарная стабильность

### 4.3 Покрытие edge-cases

**Вертикальное слияние таблиц:**
- 2-строчное (basic), 6-строчное (deep chain), malformed documents, multi-column partial merge, gridSpan+vMerge комбинации

**Комментарии (11 round-trip тестов):**
- Single/multiple/multi-run comment ranges, метаданные (author, initials, timestamp), пустой текст

**Ошибки:**
- `error_propagation_test.go` — проверяет, что ошибки corrupt XML не проглатываются

### 4.4 Visual Regression Testing

Отдельная инфраструктура в `/visual-regtest/`:
- **Docker** + LibreOffice headless + pdftoppm для рендеринга
- **SSIM** (Structural Similarity) сравнение original vs round-trip
- HTML-отчеты с side-by-side thumbnails и heatmap различий
- Настраиваемые параметры: `SSIM_THRESHOLD`, `DPI`, `WORKERS`
- Отдельные сценарии: roundtrip, replace-txt, replace-tbl, replace-content, gen-files

Это продвинутый уровень тестирования, редко встречающийся в Go-библиотеках.

### 4.5 Тестовые анти-паттерны

- **Не обнаружены**: тесты без ассертов, зависимость от внешнего состояния, flaky-тесты
- 1 `t.Skip()` в `replacetable_test.go` — единственный пропуск

---

## 5. ResourceImport (оценка: 9/10)

Система импорта ресурсов при `ReplaceWithContent()` — одна из самых сложных и одновременно чистых частей проекта.

### Архитектура

```
ResourceImporter (создается на каждый ReplaceWithContent)
├─ importNumbering()    — фаза 1 (обязательно первая)
├─ importStyles()       — фаза 2
├─ importFootnotes()    — фаза 3
├─ importEndnotes()     — фаза 4
└─ remapAll()           — единый DFS-проход для ремапа всех ID
```

### Ключевые решения

- **Idempotency flags** (`numDone`, `styleDone`, ...) — безопасный повторный вызов
- **BFS для transitive closure стилей** — basedOn, next, link зависимости
- **Детекция default style mismatch** — автоматическая материализация implicit стилей
- **Регенерация nsid** (random 8-hex) — предотвращает непреднамеренное слияние нумерации в Word
- **Shared engine для footnotes/endnotes** — `importNoteEntries()` с 10-шаговым пайплайном

### Минимальное дублирование

Каждый `resourceimport_*.go` файл содержит специализированную логику сбора и импорта — дублирование оправдано различиями в XML-структуре.

---

## 6. Безопасность и робастность

| Аспект | Статус | Комментарий |
|--------|--------|-------------|
| panic() в библиотеке | Нет | Только в init()-инициализации (допустимо) |
| Unsafe type assertions | Нет | Все с проверкой или type switch |
| Stack overflow risk | Нет | Итеративный DFS с явным стеком |
| Global mutable state | Нет | Namespace maps read-only после init() |
| Error suppression | Нет | Все ошибки обрабатываются |
| Concurrency safety | Документировано | "Package docx is not safe for concurrent use" |
| SQL/XSS/Injection | N/A | Библиотека для файлов, нет сетевого I/O |

---

## 7. Линтинг

`.golangci.yml` настроен с разумным набором линтеров:
- `govet`, `errcheck`, `staticcheck`, `revive`, `gosimple`, `ineffassign`, `unused`, `typecheck`
- Правило revive: экспортируемые символы должны иметь комментарии
- Таймаут: 5 минут

---

## 8. Итоговые оценки

| Критерий | Оценка | Комментарий |
|----------|--------|-------------|
| **Архитектура** | 9/10 | Чистая 4-слойная, нет циклических зависимостей |
| **Разделение ответственностей** | 8/10 | Отличное между слоями, Document — borderline God Object |
| **Отсутствие спагетти** | 8/10 | Несколько больших файлов, но логика структурирована |
| **Чистота проекта** | 8.5/10 | Нет TODO/FIXME, нет мертвого кода, минимальный tech debt |
| **Колхоз-фактор** | 9/10 | Минимальный — несколько magic numbers, немного сокращений |
| **Обработка ошибок** | 9/10 | Кастомные типы, обёртка, цепочки ошибок |
| **Тестирование** | 8.5/10 | 90 файлов, table-driven, round-trip, visual regression |
| **Кодогенерация** | 10/10 | YAML-схемы → Go-код, чистое разделение gen/custom |
| **API дизайн** | 8.5/10 | Sealed interfaces, options pattern, lazy properties |
| **Зрелость** | 9/10 | Продакшн-ready, зрелые паттерны |

### **Общая оценка: 8.5/10**

---

## 9. Рекомендации по улучшению

### Высокий приоритет

1. **Разбить `section.go` (722 строки)** — вынести Header/Footer в отдельный файл `header_footer.go`
2. **Разбить `oxml/table_custom.go` (1210 строк)** — cell merging в `table_merge_custom.go`
3. **Устранить дублирование итерации по block items** — извлечь общий итератор в `blkcntnr.go`

### Средний приоритет

4. **Заменить magic number `"04A0"`** в `table_custom.go:36` на именованную константу
5. **Использовать константы `EmuPerInch` etc.** вместо сырых `914400` в `shape_custom.go`
6. **Header/Footer dedup** — обобщить три dedup-метода в `section.go` через generic callback

### Низкий приоритет

7. **Переименовать сокращения** (`bic` → `container`, `hdrFtrOps` → `HeaderFooterOps`)
8. **Константа для `"continue"`** vMerge в `text_paragraph_custom.go`
9. **Централизованный error wrapper** — `func docxErr(ctx string, err error) error`

---

## 10. Заключение

**go-docx — зрелый, хорошо спроектированный проект** с чистой слоистой архитектурой, продуманной кодогенерацией и впечатляющей инфраструктурой тестирования (включая визуальную регрессию через Docker + SSIM).

Основные сильные стороны:
- Строго однонаправленные зависимости между слоями
- Нет циклических зависимостей
- Нет panic() в библиотечном коде
- Нет подавленных ошибок
- Нет глобального мутабельного состояния
- Нет технического долга (TODO/FIXME)
- Продвинутое тестирование

Основные зоны роста:
- Document как God Object (осознанный trade-off ради API-совместимости с python-docx)
- Несколько крупных файлов, которые стоит разбить
- Локальное дублирование итераторов и dedup-логики

Проект далёк от "спагетти" или "колхоза" — это хорошо структурированная, поддерживаемая кодовая база промышленного качества.
