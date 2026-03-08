package main

import (
	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
	. "github.com/vortex/go-docx/visual-regtest/internal/docfmt"
	. "github.com/vortex/go-docx/visual-regtest/internal/regtest"
)

func buildTemplate() (*docx.Document, error) {
	doc := Must(docx.New())

	// Enable first-page header/footer on section 1.
	sect := doc.Sections().Iter()[0]
	Must0(sect.SetDifferentFirstPageHeaderFooter(true))

	// ── Legend ────────────────────────────────────────────────────────
	Heading(doc, "ReplaceWithContent — Визуальный регрессионный тест", 1)
	{
		lp := Must(doc.AddParagraph(""))
		AddHighlighted(lp, "Жёлтая подсветка")
		AddPlain(lp, " = метка-заполнитель (будет заменена).")
	}
	{
		lp := Must(doc.AddParagraph(""))
		AddGreen(lp, "Зелёная подсветка")
		AddPlain(lp, " = ожидаемое содержимое.")
	}
	Note(doc, "В 02_result.docx: жёлтый текст заменён содержимым исходных документов из sources/.")
	Must(doc.AddParagraph(""))

	// ================================================================
	// A. TAG POSITION TESTS
	// ================================================================
	Heading(doc, "А. ТЕСТЫ ПОЗИЦИИ МЕТКИ", 1)

	// A1
	Heading(doc, "А1. Метка — весь текст параграфа", 2)
	SpecContent(doc, "(<ENTIRE>)", "Один абзац: «Inserted paragraph.»")
	TagPara(doc, "", "(<ENTIRE>)", "")

	// A2
	Heading(doc, "А2. Текст только перед меткой", 2)
	SpecContent(doc, "(<AFTER>)", "Абзац после текста «Перед: »")
	TagPara(doc, "Перед: ", "(<AFTER>)", "")

	// A3
	Heading(doc, "А3. Текст только после метки", 2)
	SpecContent(doc, "(<BEFORE>)", "Абзац перед текстом « :После»")
	TagPara(doc, "", "(<BEFORE>)", " :После")

	// A4
	Heading(doc, "А4. Текст с обеих сторон", 2)
	SpecContent(doc, "(<BOTH>)", "Абзац между «Слева » и « Справа»")
	TagPara(doc, "Слева ", "(<BOTH>)", " Справа")

	// A5 — cross-run tag with formatting
	Heading(doc, "А5. Cross-run метка (разбита на 3 рана с форматированием)", 2)
	SpecContent(doc, "(<CROSS>)", "Абзац; форматирование ранов сохраняется")
	Note(doc, "Метка разбита: «(<CR» жирный красный, «O» курсив синий, «SS>)» обычный.")
	{
		p := Must(doc.AddParagraph(""))
		r1, _ := p.AddRun("(<CR")
		_ = r1.SetBold(BoolPtr(true))
		_ = r1.Font().Color().SetRGB(&ColorRed)
		SetHighlightYellow(r1)
		r2, _ := p.AddRun("O")
		_ = r2.SetItalic(BoolPtr(true))
		_ = r2.Font().Color().SetRGB(&ColorBlue)
		SetHighlightYellow(r2)
		r3, _ := p.AddRun("SS>)")
		SetHighlightYellow(r3)
	}

	// A6
	Heading(doc, "А6. Две одинаковых метки в одном параграфе", 2)
	SpecContent(doc, "(<MULTI>) × 2", "Две вставки внутри одного параграфа")
	{
		p := Must(doc.AddParagraph(""))
		AddPlain(p, "Начало ")
		AddHighlighted(p, "(<MULTI>)")
		AddPlain(p, " середина ")
		AddHighlighted(p, "(<MULTI>)")
		AddPlain(p, " конец")
	}

	// A7
	Heading(doc, "А7. Метка в начале параграфа (пустой префикс)", 2)
	SpecContent(doc, "(<TAG_START>)", "Абзац; после вставки идёт « суффикс»")
	TagPara(doc, "", "(<TAG_START>)", " суффикс")

	// A8
	Heading(doc, "А8. Метка в конце параграфа (пустой суффикс)", 2)
	SpecContent(doc, "(<TAG_END>)", "Абзац; перед вставкой «префикс »")
	TagPara(doc, "префикс ", "(<TAG_END>)", "")

	// A9
	Heading(doc, "А9. Метка в параграфе с разрывом секции", 2)
	SpecContent(doc, "(<SECT>)", "Абзац; секция 2 начинается после")
	TagPara(doc, "", "(<SECT>)", "")
	Must(doc.AddSection(enum.WdSectionStartNewPage))
	Heading(doc, "— Страница 2 (секция 2 — разрыв сохранён) —", 3)
	Must(doc.AddParagraph("Если видите это на отдельной странице — разрыв секции сохранён."))

	// A10
	Heading(doc, "А10. Метка — весь текст ячейки таблицы", 2)
	SpecContent(doc, "(<CELL_FULL>)", "Абзац вставлен в ячейку [0,0]")
	{
		tbl := Must(doc.AddTable(2, 2, docx.StyleName("Table Grid")))
		SetCell(tbl, 0, 0, "(<CELL_FULL>)")
		SetCell(tbl, 0, 1, "Правая ячейка")
		SetCell(tbl, 1, 0, "Нижняя левая")
		SetCell(tbl, 1, 1, "Нижняя правая")
	}

	// A11
	Heading(doc, "А11. Метка в ячейке с текстом вокруг", 2)
	SpecContent(doc, "(<CELL_MIX>)", "Абзац между «До: » и « :После» в ячейке")
	{
		tbl := Must(doc.AddTable(1, 2, docx.StyleName("Table Grid")))
		c := Must(tbl.CellAt(0, 0))
		c.SetText("До: (<CELL_MIX>) :После")
		SetCell(tbl, 0, 1, "Соседняя ячейка")
	}

	// A12
	Heading(doc, "А12. Метка в основном верхнем колонтитуле", 2)
	SpecContent(doc, "(<HDR>)", "«Inserted paragraph.» в header")
	Note(doc, "Откройте верхний колонтитул.")
	{
		p := Must(sect.Header().AddParagraph(""))
		AddHighlighted(p, "(<HDR>)")
	}

	// A13
	Heading(doc, "А13. Метка в основном нижнем колонтитуле", 2)
	SpecContent(doc, "(<FTR>)", "«Inserted paragraph.» в footer")
	{
		p := Must(sect.Footer().AddParagraph(""))
		AddHighlighted(p, "(<FTR>)")
	}

	// A14
	Heading(doc, "А14. Метка в first-page header", 2)
	SpecContent(doc, "(<FP_HDR>)", "Абзац в колонтитуле первой страницы")
	{
		p := Must(sect.FirstPageHeader().AddParagraph(""))
		AddHighlighted(p, "(<FP_HDR>)")
	}

	// A15
	Heading(doc, "А15. Метка в first-page footer", 2)
	SpecContent(doc, "(<FP_FTR>)", "Абзац в нижнем колонтитуле первой страницы")
	{
		p := Must(sect.FirstPageFooter().AddParagraph(""))
		AddHighlighted(p, "(<FP_FTR>)")
	}

	// A16
	Heading(doc, "А16. Метка в теле комментария", 2)
	SpecContent(doc, "(<COMMENT>)", "Текст комментария заменён на абзац")
	{
		p := Must(doc.AddParagraph(""))
		rAnchor, _ := p.AddRun("Текст с комментарием")
		Must(doc.AddComment(
			[]*docx.Run{rAnchor},
			"Обзор: (<COMMENT>) — содержимое здесь",
			"Тестер", nil))
	}

	// A17
	Heading(doc, "А17. Метка в ячейке таблицы колонтитула", 2)
	SpecContent(doc, "(<HDR_CELL>)", "Абзац в ячейке таблицы header")
	Note(doc, "В верхнем колонтитуле есть таблица 1×2; метка в ячейке [0,0].")
	{
		hdr := sect.Header()
		tbl := Must(hdr.AddTable(1, 2, 9000))
		c := Must(tbl.CellAt(0, 0))
		c.SetText("(<HDR_CELL>)")
		c2 := Must(tbl.CellAt(0, 1))
		c2.SetText("Правая ячейка колонтитула")
	}

	// A18
	Heading(doc, "А18. Метка повторяется в 3 параграфах", 2)
	SpecContent(doc, "(<REPEAT>) × 3", "Три вставки (по одной на параграф)")
	TagPara(doc, "Первый: ", "(<REPEAT>)", "")
	TagPara(doc, "Второй: ", "(<REPEAT>)", "")
	TagPara(doc, "Третий: ", "(<REPEAT>)", "")

	// A19
	Heading(doc, "А19. Метка между двумя разрывами секций", 2)
	SpecContent(doc, "(<BETWEEN>)", "Абзац; sectPr по обе стороны сохранены")
	Must(doc.AddSection(enum.WdSectionStartNewPage))
	TagPara(doc, "", "(<BETWEEN>)", "")
	Must(doc.AddSection(enum.WdSectionStartNewPage))
	Heading(doc, "— Следующая секция (разрывы сохранены) —", 3)
	Must(doc.AddParagraph("Секция после (<BETWEEN>)."))

	// ================================================================
	// B. SOURCE CONTENT TYPE TESTS
	// ================================================================
	Heading(doc, "Б. ТЕСТЫ ТИПОВ ИСХОДНОГО СОДЕРЖИМОГО", 1)

	Heading(doc, "Б1. Источник: несколько абзацев", 2)
	SpecContent(doc, "(<SRC_MULTI>)", "Три абзаца на русском")
	TagPara(doc, "", "(<SRC_MULTI>)", "")

	Heading(doc, "Б2. Источник: только таблица", 2)
	SpecContent(doc, "(<SRC_TABLE>)", "Таблица 2×3 с кириллицей")
	TagPara(doc, "", "(<SRC_TABLE>)", "")

	Heading(doc, "Б3. Источник: абзац + таблица + абзац", 2)
	SpecContent(doc, "(<SRC_COMPLEX>)", "Абзац, таблица 2×2, абзац")
	TagPara(doc, "", "(<SRC_COMPLEX>)", "")

	Heading(doc, "Б4. Источник: абзац + изображение + абзац", 2)
	SpecContent(doc, "(<SRC_IMAGE>)", "Текст, градиент PNG 1.5\"×0.75\", текст")
	TagPara(doc, "", "(<SRC_IMAGE>)", "")

	Heading(doc, "Б5. Источник: стиль Heading 1", 2)
	SpecContent(doc, "(<SRC_HEADING>)", "Текст со стилем Heading 1")
	TagPara(doc, "", "(<SRC_HEADING>)", "")

	Heading(doc, "Б6. Источник: пустое тело", 2)
	SpecContent(doc, "(<SRC_EMPTY>)", "Пустой абзац (метка → пустая строка)")
	TagPara(doc, "", "(<SRC_EMPTY>)", "")

	Heading(doc, "Б7. Источник: большой контент (стресс-тест)", 2)
	SpecContent(doc, "(<SRC_LARGE>)", "5 абзацев + таблица 4×3")
	TagPara(doc, "", "(<SRC_LARGE>)", "")

	Heading(doc, "Б8. Источник: жирный + курсив", 2)
	SpecContent(doc, "(<SRC_STYLED>)", "«Жирный текст, обычный, курсив.»")
	TagPara(doc, "", "(<SRC_STYLED>)", "")

	// ================================================================
	// C. CYRILLIC TESTS
	// ================================================================
	Heading(doc, "В. ТЕСТЫ КИРИЛЛИЦЫ", 1)

	Heading(doc, "В1. Кириллическое имя метки", 2)
	SpecContent(doc, "(<КИРИЛЛИЦА>)", "Русский текст + таблица 2×2")
	TagPara(doc, "", "(<КИРИЛЛИЦА>)", "")

	Heading(doc, "В2. Кириллический текст вокруг метки", 2)
	SpecContent(doc, "(<ЗАМЕНА>)", "Русский абзац между «Текст до: » и « :текст после»")
	TagPara(doc, "Текст до: ", "(<ЗАМЕНА>)", " :текст после")

	// ================================================================
	// D. MULTI-SOURCE TESTS
	// ================================================================
	Heading(doc, "Г. ТЕСТЫ С НЕСКОЛЬКИМИ ИСТОЧНИКАМИ", 1)

	Heading(doc, "Г1. Разные метки — разные источники", 2)
	SpecContent(doc, "(<SRC_A>) и (<SRC_B>)", "«[Источник A]» и «[Источник Б]»")
	TagPara(doc, "", "(<SRC_A>)", "")
	TagPara(doc, "", "(<SRC_B>)", "")

	// ================================================================
	// F. IMPORT FORMAT MODE TESTS
	// ================================================================
	Heading(doc, "Е. ТЕСТЫ РЕЖИМОВ ИМПОРТА (ImportFormatMode)", 1)
	Note(doc, "Источник src_conflict_style.docx: Heading 1 переопределён (красный, 20pt).")
	Note(doc, "Источник src_numbered_list.docx: нумерованный список из 3 пунктов.")
	Note(doc, "Источник src_unique_style.docx: кастомный стиль MyCustomRed (14pt, красный).")
	Must(doc.AddParagraph(""))

	Heading(doc, "Е1. UseDestinationStyles + конфликтующий стиль", 2)
	SpecContent(doc, "(<IFM_USE_DEST>)", "Heading 1 с форматированием ЦЕЛЕВОГО документа (синий, стандартный)")
	Note(doc, "Стиль target побеждает: заголовок выглядит как обычный Heading 1.")
	TagPara(doc, "", "(<IFM_USE_DEST>)", "")

	Heading(doc, "Е2. KeepSourceFormatting + конфликтующий стиль", 2)
	SpecContent(doc, "(<IFM_KEEP_SRC>)", "Heading 1 с ИСХОДНЫМ форматированием (красный, 20pt) через direct attrs")
	Note(doc, "Свойства источника развёрнуты в прямые атрибуты параграфа/рана.")
	TagPara(doc, "", "(<IFM_KEEP_SRC>)", "")

	Heading(doc, "Е3. KeepSourceFormatting + ForceCopyStyles", 2)
	SpecContent(doc, "(<IFM_KEEP_SRC_FORCE>)", "Heading 1 скопирован как Heading1_0 (красный, 20pt)")
	Note(doc, "Стиль скопирован с суффиксом _0. Вставленный контент ссылается на Heading1_0.")
	TagPara(doc, "", "(<IFM_KEEP_SRC_FORCE>)", "")

	Heading(doc, "Е4. KeepDifferentStyles + конфликтующий стиль", 2)
	SpecContent(doc, "(<IFM_KEEP_DIFF>)", "Форматирование различается → развёрнуто в direct attrs")
	Note(doc, "Гибрид: стили разные → поведение как KeepSourceFormatting.")
	TagPara(doc, "", "(<IFM_KEEP_DIFF>)", "")

	Heading(doc, "Е5. KeepDifferentStyles + ForceCopyStyles", 2)
	SpecContent(doc, "(<IFM_KEEP_DIFF_FORCE>)", "Форматирование различается → стиль скопирован с суффиксом")
	Note(doc, "Гибрид + ForceCopyStyles: разные → копия с суффиксом.")
	TagPara(doc, "", "(<IFM_KEEP_DIFF_FORCE>)", "")

	Heading(doc, "Е6. KeepSourceNumbering = true", 2)
	SpecContent(doc, "(<IFM_KEEP_NUM>)", "Нумерованный список как отдельное определение (не слитый)")
	Note(doc, "Нумерация сохранена как отдельный abstractNum, начинается с 1.")
	TagPara(doc, "", "(<IFM_KEEP_NUM>)", "")

	Heading(doc, "Е7. KeepSourceNumbering = false (merge)", 2)
	SpecContent(doc, "(<IFM_MERGE_NUM>)", "Нумерованный список слит с существующим определением")
	Note(doc, "Нумерация сливается с matching target list definition.")
	TagPara(doc, "", "(<IFM_MERGE_NUM>)", "")

	Heading(doc, "Е8. KeepSourceFormatting + уникальный стиль", 2)
	SpecContent(doc, "(<IFM_UNIQUE_STYLE>)", "MyCustomRed (14pt, красный) — deep-copy в target")
	Note(doc, "Стиль отсутствует в target → копируется целиком (все 3 режима согласны).")
	TagPara(doc, "", "(<IFM_UNIQUE_STYLE>)", "")

	// ================================================================
	// E. EDGE CASES
	// ================================================================
	Heading(doc, "Д. ГРАНИЧНЫЕ СЛУЧАИ", 1)

	Heading(doc, "Д1. Метка отсутствует (no match)", 2)
	SpecContent(doc, "(<NO_MATCH>)", "0 замен — метки нет в шаблоне")
	Note(doc, "Метка (<NO_MATCH>) НЕ присутствует. Замена вернёт 0.")

	Heading(doc, "Д2. Пустая строка поиска", 2)
	SpecContent(doc, `""`, "0 замен — пустая строка всегда возвращает 0")
	Note(doc, "Тестируется программно (не визуально).")

	Heading(doc, "Д3. Само-ссылка (source == target)", 2)
	SpecContent(doc, "(<SELF>)", "Копия тела самого документа на момент вызова")
	Note(doc, "Источник — сам целевой документ. Deep-copy предотвращает бесконечный цикл.")
	TagPara(doc, "", "(<SELF>)", "")

	Heading(doc, "Д4. Параграф без метки (sanity check)", 2)
	SpecContent(doc, "(нет метки)", "Параграф ниже должен остаться без изменений")
	Must(doc.AddParagraph("Этот параграф не содержит меток и должен сохраниться как есть."))

	return doc, nil
}
