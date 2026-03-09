# План: Парсер темы и разрешение тема-зависимых свойств при импорте

## Проблема

go-docx сохраняет `w:themeColor`, `w:themeShade`, `w:themeTint`, `w:asciiTheme` и т.д. **as-is** при импорте. Когда тема исходного документа отличается от целевого, Word пересчитывает значения по целевой теме → другие цвета/шрифты.

Aspose.Words разрешает все тема-зависимые свойства в явные значения на этапе импорта.

## Структура theme1.xml

```xml
<a:theme name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <!-- accent2..accent6, hlink, folHlink -->
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>
```

**Цветовая схема**: 12 слотов — dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink.
**Шрифтовая схема**: major/minor × latin/ea/cs = 6 значений.

## Фазы реализации

---

### Фаза 1: ThemePart инфраструктура (~100 строк)

**Сложность**: Низкая — копируем паттерн StylesPart.

#### 1.1 `pkg/docx/parts/theme.go` — новый файл

```go
type ThemePart struct {
    *opc.XmlPart
}

func NewThemePart(xp *opc.XmlPart) *ThemePart
func LoadThemePart(...) (opc.Part, error)  // PartConstructor
```

Без OXML-генерации: парсинг через etree (тема — read-only при импорте).

#### 1.2 `pkg/docx/parts/register.go` — регистрация

```go
f.Register(opc.CTOfcTheme, LoadThemePart)
```

#### 1.3 `pkg/docx/parts/document.go` — аксессор

```go
func (dp *DocumentPart) ThemePart() (*ThemePart, error)
```

По relationship RTTheme. Без GetOrAdd — тему не создаём, только читаем.

---

### Фаза 2: Парсинг темы и разрешение значений (~200 строк)

**Сложность**: Средняя — алгоритмы shade/tint требуют аккуратности.

#### 2.1 `pkg/docx/theme.go` — новый файл

**Структуры:**

```go
// ThemeData — извлечённые данные из theme1.xml (read-only).
type ThemeData struct {
    Colors map[string]string  // "accent1" → "4472C4", "dk1" → "000000"
    MajorFont FontSlot        // latin, ea, cs
    MinorFont FontSlot
}

type FontSlot struct {
    Latin, EA, CS string
}
```

**Функции парсинга:**

```go
// parseThemeData извлекает цвета и шрифты из theme element.
// Обходит a:themeElements/a:clrScheme, a:fontScheme.
// Для a:sysClr берёт @lastClr (кешированное значение), для a:srgbClr — @val.
func parseThemeData(themeEl *etree.Element) *ThemeData
```

**Функции разрешения:**

```go
// resolveThemeColor разрешает themeColor + shade/tint → 6-значный RGB hex.
// WML shade: newChannel = channel * shade / 255
// WML tint:  newChannel = 255 - (255 - channel) * (255 - tint) / 255
func (td *ThemeData) resolveThemeColor(themeColor string, shade, tint int) string

// resolveThemeFont разрешает themeFont slot ("minorHAnsi", "majorEastAsia") → имя шрифта.
// Маппинг: asciiTheme/hAnsiTheme → latin, eastAsiaTheme → ea, cstheme → cs.
// "minor*" → MinorFont, "major*" → MajorFont.
func (td *ThemeData) resolveThemeFont(themeSlot string) string
```

**Алгоритмы shade/tint** (WML, не DML):

| Модификатор | Формула (для каждого R,G,B канала) |
|-------------|-------------------------------------|
| `w:themeShade` (0–255) | `new = old * shade / 255` |
| `w:themeTint` (0–255) | `new = 255 - (255 - old) * (255 - tint) / 255` |

Ни HSL, ни sRGB gamma — простое линейное в RGB пространстве (по спеку WML §17.3.2.6).

---

### Фаза 3: Интеграция в import pipeline (~150 строк)

**Сложность**: Средняя — нужно пройти по всем элементам.

#### 3.1 `pkg/docx/theme_resolve.go` — новый файл

```go
// resolveThemeProperties обходит элементы DFS и заменяет
// тема-зависимые свойства на явные значения.
//
// Обрабатывает:
//   w:color/@w:themeColor → w:val = resolved RGB, удалить themeColor/Shade/Tint
//   w:rFonts/@w:asciiTheme → w:ascii = resolved font, удалить asciiTheme
//   w:rFonts/@w:hAnsiTheme → w:hAnsi = resolved font, удалить hAnsiTheme
//   w:rFonts/@w:eastAsiaTheme → w:eastAsia = resolved font, удалить eastAsiaTheme
//   w:rFonts/@w:cstheme → w:cs = resolved font, удалить cstheme
//   w:shd/@w:themeColor (фон) → аналогично
//   w:shd/@w:themeFill → аналогично
func resolveThemeProperties(elements []*etree.Element, srcTheme *ThemeData)
```

#### 3.2 Точка вызова — `contentdata.go`, `prepareContentElements`

В текущем pipeline:
```
deep-copy → sanitize → renumberBookmarks →
  materializeImplicitStyles → expandDirectFormatting →
  remapAll → importRIds → renumberDrawingIDs
```

Добавляется **перед** remapAll (после expand, чтобы expanded rPr тоже обработались):
```
deep-copy → sanitize → renumberBookmarks →
  materializeImplicitStyles → expandDirectFormatting →
  ★ resolveThemeProperties →
  remapAll → importRIds → renumberDrawingIDs
```

#### 3.3 Также обработать: footnotes/endnotes, comments, headers/footers

В `resourceimport_notes.go` → `importNoteEntries` — аналогичный вызов после expand.
В `section.go` → headers/footers (если они проходят через `prepareContentElements`, проверить).

#### 3.4 ResourceImporter — хранение ThemeData

```go
type ResourceImporter struct {
    ...
    srcTheme *ThemeData  // lazily parsed from source ThemePart
}

func (ri *ResourceImporter) sourceTheme() *ThemeData  // lazy accessor
```

---

### Фаза 4: docDefaults delta — учёт темных шрифтов (~50 строк)

**Сложность**: Низкая.

В `ooxmlImplicitRPr` добавить тема-атрибуты для rFonts:

```go
"rFonts": {
    "w:ascii": "Times New Roman", "w:hAnsi": "Times New Roman",
    "w:eastAsia": "Times New Roman", "w:cs": "Times New Roman",
    // Тема-атрибуты: отсутствие = нет темного шрифта
    "w:asciiTheme": "", "w:hAnsiTheme": "",
    "w:eastAsiaTheme": "", "w:cstheme": "",
},
```

И в `diffProperties` — учитывать, что тема-атрибут и явный атрибут на одном элементе (`w:rFonts`) взаимодействуют: явный шрифт перекрывает тему.

---

## Что НЕ входит в этот план

- **Копирование темы** (замена целевой темы исходной) — опасно, ломает форматирование существующего контента
- **DML shade/tint** (a:shade, a:tint в DrawingML) — другой алгоритм (HSL), используется внутри шейпов; шейпы копируются as-is
- **ThemeOverride** — редкий случай, можно добавить позже
- **VML spid перенумерация** — отдельная задача
- **Compat mode** перенос — отдельная задача

## Оценка объёма

| Фаза | Файлы | Строк | Сложность |
|-------|-------|-------|-----------|
| 1. ThemePart инфраструктура | 3 файла (новый + 2 правки) | ~100 | Низкая |
| 2. Парсинг + разрешение | 1 новый файл | ~200 | Средняя |
| 3. Интеграция в pipeline | 1 новый + 2–3 правки | ~150 | Средняя |
| 4. docDefaults delta | 1 правка | ~50 | Низкая |
| **Тесты** | 2–3 файла | ~300 | — |
| **Итого** | ~8 файлов | **~800** | **Средняя** |

## Файлы для изменения

| Файл | Действие |
|------|----------|
| `pkg/docx/parts/theme.go` | **Создать** — ThemePart |
| `pkg/docx/parts/register.go` | Добавить регистрацию CTOfcTheme |
| `pkg/docx/parts/document.go` | Добавить ThemePart() аксессор |
| `pkg/docx/theme.go` | **Создать** — ThemeData, parseThemeData, resolve* |
| `pkg/docx/theme_resolve.go` | **Создать** — resolveThemeProperties (DFS walker) |
| `pkg/docx/resourceimport.go` | Добавить srcTheme поле и lazy accessor |
| `pkg/docx/contentdata.go` | Вызов resolveThemeProperties в prepareContentElements |
| `pkg/docx/resourceimport_notes.go` | Вызов resolveThemeProperties в importNoteEntries |
| `pkg/docx/resourceimport_styles.go` | Обновить ooxmlImplicitRPr (тема-атрибуты) |

## Риски

1. **a:sysClr без lastClr** — некоторые темы могут не иметь `lastClr`. Fallback: маппинг системных цветов (windowText→000000, window→FFFFFF).
2. **Пустые шрифты в теме** (`<a:ea typeface=""/>`) — значит "не задан", нельзя разрешать как пустую строку. Пропускать разрешение.
3. **w:rFonts с ОБОИМИ explicit + theme** — explicit берёт приоритет в Word, но при разрешении нужно: установить explicit = resolved(theme), удалить theme-атрибут, не трогать существующий explicit если он уже задан.
4. **Цвет "auto"** — `w:val="auto"` означает системный цвет (обычно чёрный). Не путать с тема-цветами.
