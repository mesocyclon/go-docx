// Business proposal generator — demonstrates go-docx capabilities
// (except element replacement) in a single bilingual (RU/EN) document.
package main

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"log"
	"os"
	"path/filepath"
	"runtime"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ---------------------------------------------------------------------------
// Color palette
// ---------------------------------------------------------------------------

const (
	cNavy     = "1B3A5C"
	cBlue     = "2980B9"
	cSoftBlue = "4A7FB5"
	cGreen    = "27AE60"
	cOrange   = "E67E22"
	cRed      = "C0392B"
	cPurple   = "8E44AD"
	cGray     = "666666"
	cGrayLt   = "999999"
	cGrayXLt  = "AAAAAA"
	cDark     = "333333"
)

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

func main() {
	doc := ok(docx.New())

	configure(doc)
	buildTitlePage(doc)
	buildAboutSection(doc)
	buildTeamSection(doc)
	buildServicesSection(doc)
	buildMethodologySection(doc)
	buildTimelineSection(doc)
	buildTermsSection(doc)
	buildContactSection(doc)

	out := outputPath("proposal.docx")
	must(doc.SaveFile(out))
	log.Printf("saved %s", out)
}

// ---------------------------------------------------------------------------
// Configuration: metadata, page, styles, headers/footers
// ---------------------------------------------------------------------------

func configure(doc *docx.Document) {
	// Metadata
	p := ok(doc.CoreProperties())
	p.SetTitle("Коммерческое предложение — Example Solutions")
	p.SetAuthor("Example Solutions")
	p.SetSubject("IT-услуги и разработка / IT Services & Development")
	p.SetCategory("Commercial Proposal")
	p.SetKeywords("IT, development, consulting, разработка, консалтинг")
	p.SetLanguage("ru-RU")
	now := time.Now()
	p.SetCreated(now)
	p.SetModified(now)

	// Page layout — A4
	s := ok(doc.Sections().Get(0))
	must(s.SetPageWidth(ip(docx.Mm(210).Twips())))
	must(s.SetPageHeight(ip(docx.Mm(297).Twips())))
	must(s.SetTopMargin(ip(docx.Cm(2.5).Twips())))
	must(s.SetBottomMargin(ip(docx.Cm(2).Twips())))
	must(s.SetLeftMargin(ip(docx.Cm(2.5).Twips())))
	must(s.SetRightMargin(ip(docx.Cm(2).Twips())))
	must(s.SetDifferentFirstPageHeaderFooter(true))

	// Styles
	styles := ok(doc.Styles())

	st := ok(styles.AddStyle("Quote Block", enum.WdStyleTypeParagraph, false))
	must(st.Font().SetName(sp("Georgia")))
	must(st.Font().SetSize(lp(docx.Pt(10.5))))
	must(st.Font().SetItalic(bp(true)))
	must(st.Font().Color().SetRGB(rgb("2E4057")))
	must(st.ParagraphFormat().SetLeftIndent(ip(docx.Cm(1.2).Twips())))
	must(st.ParagraphFormat().SetRightIndent(ip(docx.Cm(1.2).Twips())))
	must(st.ParagraphFormat().SetSpaceBefore(ip(docx.Pt(6).Twips())))
	must(st.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(6).Twips())))

	// Header/Footer — first page empty, rest with branding
	sect := ok(doc.Sections().Get(0))
	ok(sect.FirstPageHeader().AddParagraph(""))
	ok(sect.FirstPageFooter().AddParagraph(""))

	hp := ok(sect.Header().AddParagraph(""))
	must(hp.SetAlignment(ap(enum.WdParagraphAlignmentRight)))
	r := ok(hp.AddRun("Example Solutions"))
	must(r.Font().SetSize(lp(docx.Pt(8))))
	must(r.Font().SetSmallCaps(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGrayLt)))
	r = ok(hp.AddRun("  |  Коммерческое предложение"))
	must(r.Font().SetSize(lp(docx.Pt(8))))
	must(r.Font().SetItalic(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGrayLt)))

	fp := ok(sect.Footer().AddParagraph(""))
	must(fp.SetAlignment(ap(enum.WdParagraphAlignmentCenter)))
	r = ok(fp.AddRun("Конфиденциально / Confidential"))
	must(r.Font().SetSize(lp(docx.Pt(7.5))))
	must(r.Font().Color().SetRGB(rgb(cGrayXLt)))
}

// ---------------------------------------------------------------------------
// Page 1: Title
// ---------------------------------------------------------------------------

func buildTitlePage(doc *docx.Document) {
	// Logo
	logo := generateLogo(600, 100)
	w, h := docx.Cm(14).Emu(), docx.Cm(2.3).Emu()
	ok(doc.AddPicture(logo, &w, &h))

	emptyN(doc, 6)

	// Title
	addStyledParagraph(doc, "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ",
		enum.WdParagraphAlignmentCenter, "Calibri", 28, true, false, cNavy)

	// Subtitle
	p := ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentCenter)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(4).Twips())))
	r := ok(p.AddRun("Commercial Proposal"))
	must(r.Font().SetName(sp("Calibri")))
	must(r.Font().SetSize(lp(docx.Pt(16))))
	must(r.Font().SetItalic(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cSoftBlue)))

	// Horizontal line (thin paragraph with underline)
	p = ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentCenter)))
	must(p.ParagraphFormat().SetSpaceBefore(ip(docx.Pt(12).Twips())))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(12).Twips())))
	r = ok(p.AddRun("                                                                                       "))
	uv := docx.UnderlineStyle(enum.WdUnderlineSingle)
	must(r.Font().SetUnderline(&uv))
	must(r.Font().Color().SetRGB(rgb(cGrayXLt)))

	emptyN(doc, 1)

	// Company name
	p = ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentCenter)))
	r = ok(p.AddRun("Example Solutions"))
	must(r.Font().SetSize(lp(docx.Pt(14))))
	must(r.Font().SetSmallCaps(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cDark)))

	// Tagline
	p = ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentCenter)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(24).Twips())))
	r = ok(p.AddRun("Инженерия будущего / Engineering the Future"))
	must(r.Font().SetSize(lp(docx.Pt(10))))
	must(r.Font().SetItalic(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGrayLt)))

	// Addressee info
	addSimpleParagraph(doc, "Подготовлено для: ООО «Заказчик»", 10, false, cGray,
		enum.WdParagraphAlignmentCenter)
	addSimpleParagraph(doc, fmt.Sprintf("Дата: %s", time.Now().Format("02 января 2006")), 10, false, cGray,
		enum.WdParagraphAlignmentCenter)
	addSimpleParagraph(doc, "Ref: NT-2026-0342", 10, false, cGray,
		enum.WdParagraphAlignmentCenter)
}

// ---------------------------------------------------------------------------
// Pages 2-3: About Company
// ---------------------------------------------------------------------------

func buildAboutSection(doc *docx.Document) {
	ok(doc.AddPageBreak())
	ok(doc.AddHeading("О компании / About Us", 1))

	// Introduction with rich formatting
	p := ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(8).Twips())))
	ls := docx.LineSpacingMultiple(1.15)
	must(p.ParagraphFormat().SetLineSpacing(&ls))

	r := ok(p.AddRun("Example Solutions"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cNavy)))

	ok(p.AddRun(
		" — международная IT-компания, основанная в 2013 году в Москве. " +
			"Мы специализируемся на проектировании и разработке "))

	r = ok(p.AddRun("высоконагруженных распределённых систем"))
	uv := docx.UnderlineStyle(enum.WdUnderlineDouble)
	must(r.Font().SetUnderline(&uv))

	ok(p.AddRun(", облачной инфраструктуре и "))

	r = ok(p.AddRun("AI/ML-решениях"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().SetItalic(bp(true)))
	must(r.Font().SetHighlightColor(hlPtr(enum.WdColorIndexYellow)))

	ok(p.AddRun(
		". За 12 лет работы мы реализовали более 200 проектов для компаний " +
			"из финансового сектора, телекоммуникаций, промышленности и государственного управления."))

	// Second paragraph
	p = ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(8).Twips())))
	ok(p.AddRun(
		"Наша команда объединяет "))
	r = ok(p.AddRun("300+"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().SetSuperscript(bp(true)))
	ok(p.AddRun(
		" инженеров из России, Германии, Сингапура и Казахстана. " +
			"Мы являемся сертифицированными партнёрами AWS, Google Cloud и Yandex Cloud. " +
			"Все проекты ведутся по методологии Agile с еженедельной отчётностью и " +
			"прозрачным управлением рисками."))

	// Epigraph
	ok(doc.AddParagraph(
		"«Инженерное мастерство — это искусство делать из сложного простое» / "+
			"\"Engineering mastery is the art of making the complex simple\"",
		docx.StyleName("Quote Block"),
	))

	// Key facts
	ok(doc.AddHeading("Ключевые показатели / Key Metrics", 2))

	facts := []struct{ label, value string }{
		{"Год основания / Founded", "2013"},
		{"Штаб-квартира / Headquarters", "Москва, Россия"},
		{"Офисы / Offices", "Санкт-Петербург, Berlin, Singapore, Almaty"},
		{"Сотрудников / Employees", "320+"},
		{"Реализованных проектов / Delivered projects", "200+"},
		{"Enterprise-клиентов / Enterprise clients", "150+"},
		{"Основной стек / Primary stack", "Go, Rust, Python, TypeScript, K8s"},
		{"Сертификации / Certifications", "ISO 27001, SOC 2 Type II, AWS Partner"},
	}
	for _, f := range facts {
		addFactRow(doc, f.label, f.value)
	}

	// Visual contrast: old vs new
	emptyN(doc, 1)
	p = ok(doc.AddParagraph(""))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(6).Twips())))
	r = ok(p.AddRun("Устаревший подход: монолит + ручное развёртывание"))
	must(r.Font().SetStrike(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGrayLt)))
	must(r.Font().SetSize(lp(docx.Pt(10))))

	p = ok(doc.AddParagraph(""))
	r = ok(p.AddRun("Наш подход: "))
	must(r.Font().SetSize(lp(docx.Pt(10))))
	r = ok(p.AddRun("микросервисы + Infrastructure as Code + непрерывная доставка"))
	must(r.Font().SetAllCaps(bp(true)))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGreen)))
	must(r.Font().SetSize(lp(docx.Pt(10))))
}

// ---------------------------------------------------------------------------
// Page 3: Team & Competencies
// ---------------------------------------------------------------------------

func buildTeamSection(doc *docx.Document) {
	ok(doc.AddPageBreak())
	ok(doc.AddHeading("Команда и компетенции / Team & Expertise", 1))

	p := ok(doc.AddParagraph(
		"Наша организационная структура обеспечивает полный цикл разработки — " +
			"от бизнес-анализа до эксплуатации в production-среде. " +
			"Ниже представлено распределение специалистов по направлениям.",
	))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(12).Twips())))

	tbl := ok(doc.AddTable(8, 4, docx.StyleName("Table Grid")))
	must(tbl.SetAlignment(taPtr(enum.WdTableAlignmentCenter)))

	// Header
	setTableHeader(tbl, 0, []string{
		"Направление / Department",
		"Специалисты / Headcount",
		"Ключевые технологии / Key Technologies",
		"Средний опыт / Avg. Experience",
	})

	type teamRow struct{ dept, count, tech, exp string }
	data := []teamRow{
		{"Backend Engineering", "85", "Go, Rust, gRPC, PostgreSQL", "7+ лет / years"},
		{"Frontend & Mobile", "45", "React, TypeScript, Flutter", "5+ лет / years"},
		{"Cloud & DevOps", "50", "Kubernetes, Terraform, AWS, GCP", "6+ лет / years"},
		{"AI/ML & Data", "35", "Python, PyTorch, MLflow, Spark", "5+ лет / years"},
		{"QA & Automation", "40", "Go, Playwright, k6, Allure", "4+ лет / years"},
		{"Architecture & BA", "30", "UML, C4, Domain-Driven Design", "10+ лет / years"},
		{"Project Management", "35", "Agile, SAFe, Jira, Confluence", "8+ лет / years"},
	}
	for ri, d := range data {
		ok(tbl.CellAt(ri+1, 0)).SetText(d.dept)
		ok(tbl.CellAt(ri+1, 1)).SetText(d.count)
		must(ok(tbl.CellAt(ri+1, 1)).SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
		ok(tbl.CellAt(ri+1, 2)).SetText(d.tech)
		ok(tbl.CellAt(ri+1, 3)).SetText(d.exp)
		must(ok(tbl.CellAt(ri+1, 3)).SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
	}

	for ri := range 8 {
		row := ok(tbl.Rows().Get(ri))
		must(row.SetHeight(ip(docx.Cm(0.7).Twips())))
		must(row.SetHeightRule(rhPtr(enum.WdRowHeightRuleAtLeast)))
	}

	emptyN(doc, 1)
	// Comment on data
	cp := ok(doc.AddParagraph(""))
	cr := ok(cp.AddRun("Данные о численности и компетенциях актуальны на I квартал 2026 г."))
	must(cr.Font().SetSize(lp(docx.Pt(9))))
	must(cr.Font().SetItalic(bp(true)))
	must(cr.Font().Color().SetRGB(rgb(cGrayLt)))
	initials := "HR"
	ok(doc.AddComment([]*docx.Run{cr},
		"Обновить перед отправкой: запросить актуальные данные у HR-департамента",
		"Example QA", &initials))
}

// ---------------------------------------------------------------------------
// Page 4: Services & Pricing
// ---------------------------------------------------------------------------

func buildServicesSection(doc *docx.Document) {
	ok(doc.AddPageBreak())
	ok(doc.AddHeading("Услуги и тарифы / Services & Pricing", 1))

	p := ok(doc.AddParagraph(
		"Мы предлагаем комплексные IT-услуги в трёх основных направлениях. " +
			"Каждое направление доступно в двух тарифных планах: Standard для средних " +
			"компаний и Enterprise для крупных организаций с повышенными требованиями " +
			"к SLA и безопасности. Стоимость указана в условных единицах за месяц.",
	))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(12).Twips())))

	tbl := ok(doc.AddTable(7, 4, docx.StyleName("Table Grid")))
	must(tbl.SetAlignment(taPtr(enum.WdTableAlignmentCenter)))

	setTableHeader(tbl, 0, []string{
		"Услуга / Service", "Описание / Description",
		"Тариф / Tier", "Цена / Price",
	})

	type svcRow struct{ svc, desc, tier, price string }
	data := []svcRow{
		{"Cloud Infrastructure\nОблачная инфраструктура",
			"Проектирование Kubernetes-кластеров, настройка CI/CD-пайплайнов, " +
				"мониторинг и алертинг (Prometheus, Grafana), автоматизация через Terraform",
			"Standard", "3 500 у.е."},
		{"", "", "Enterprise", "8 200 у.е."},
		{"Backend Development\nСерверная разработка",
			"Микросервисная архитектура на Go/Rust, проектирование API-шлюзов, " +
				"интеграция с ERP/CRM-системами, высоконагруженные очереди сообщений",
			"Standard", "5 000 у.е."},
		{"", "", "Enterprise", "12 000 у.е."},
		{"AI/ML Consulting\nAI/ML-консалтинг",
			"Обучение и дообучение моделей, построение MLOps-пайплайнов, " +
				"внедрение LLM-ассистентов, аналитика данных и предиктивное моделирование",
			"Standard", "7 500 у.е."},
		{"", "", "Enterprise", "18 000 у.е."},
	}
	for ri, d := range data {
		row := ri + 1
		if d.svc != "" {
			ok(tbl.CellAt(row, 0)).SetText(d.svc)
			ok(tbl.CellAt(row, 1)).SetText(d.desc)
		}
		c2 := ok(tbl.CellAt(row, 2))
		c2.SetText(d.tier)
		must(c2.SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))

		c3 := ok(tbl.CellAt(row, 3))
		c3.SetText(d.price)
		must(c3.SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
	}

	// Vertical merges
	for _, pair := range [][2]int{{1, 2}, {3, 4}, {5, 6}} {
		ok(ok(tbl.CellAt(pair[0], 0)).Merge(ok(tbl.CellAt(pair[1], 0))))
		ok(ok(tbl.CellAt(pair[0], 1)).Merge(ok(tbl.CellAt(pair[1], 1))))
	}

	for ri := range 7 {
		row := ok(tbl.Rows().Get(ri))
		must(row.SetHeight(ip(docx.Cm(0.9).Twips())))
		must(row.SetHeightRule(rhPtr(enum.WdRowHeightRuleAtLeast)))
	}

	// Footnote
	emptyN(doc, 1)
	p = ok(doc.AddParagraph(""))
	r := ok(p.AddRun("*"))
	must(r.Font().SetSuperscript(bp(true)))
	must(r.Font().SetSize(lp(docx.Pt(8))))
	r = ok(p.AddRun(" Все цены указаны без НДС (20%). Индивидуальные условия и скидки при годовом контракте обсуждаются отдельно."))
	must(r.Font().SetSize(lp(docx.Pt(8))))
	must(r.Font().SetItalic(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGray)))

	// Differentiators
	emptyN(doc, 1)
	ok(doc.AddHeading("Отличия Enterprise-тарифа / Enterprise Tier Benefits", 2))

	benefits := []string{
		"Выделенный технический аккаунт-менеджер (TAM) с опытом в домене клиента",
		"SLA 99.9% с финансовыми гарантиями и штрафными санкциями",
		"Круглосуточная поддержка 24/7 с временем реакции не более 15 минут",
		"Приоритетный доступ к новым продуктам и бета-функциям",
		"Ежеквартальный аудит безопасности и performance review",
	}
	for _, b := range benefits {
		p = ok(doc.AddParagraph(""))
		must(p.ParagraphFormat().SetLeftIndent(ip(docx.Cm(0.5).Twips())))
		must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(3).Twips())))
		r = ok(p.AddRun("—  "))
		must(r.Font().SetBold(bp(true)))
		must(r.Font().Color().SetRGB(rgb(cNavy)))
		must(r.Font().SetSize(lp(docx.Pt(10))))
		r = ok(p.AddRun(b))
		must(r.Font().SetSize(lp(docx.Pt(10))))
	}
}

// ---------------------------------------------------------------------------
// Page 5: Project Methodology
// ---------------------------------------------------------------------------

func buildMethodologySection(doc *docx.Document) {
	ok(doc.AddPageBreak())
	ok(doc.AddHeading("Методология работы / Project Methodology", 1))

	p := ok(doc.AddParagraph(
		"Мы используем гибридный подход, сочетающий лучшие практики Agile и " +
			"классического проектного управления. Каждый проект проходит через " +
			"четыре ключевые фазы, обеспечивающие предсказуемость результатов " +
			"и управляемость рисков.",
	))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(12).Twips())))

	type phaseInfo struct {
		num, name, nameEn, desc string
	}
	phases := []phaseInfo{
		{"01", "Исследование и анализ", "Discovery & Analysis",
			"Глубокое погружение в бизнес-процессы заказчика. Формирование технического " +
				"задания, анализ рисков, определение метрик успеха. Результат: " +
				"детальный документ требований (SRS) и архитектурное решение (SAD)."},
		{"02", "Проектирование и прототипирование", "Design & Prototyping",
			"Разработка системной архитектуры, проектирование API-контрактов, " +
				"создание интерактивных прототипов UI/UX. Валидация решений " +
				"с ключевыми стейкхолдерами до начала разработки."},
		{"03", "Итеративная разработка", "Iterative Development",
			"Двухнедельные спринты с демонстрацией результатов. Непрерывная интеграция " +
				"и развёртывание (CI/CD). Автоматическое тестирование с покрытием не менее 80%. " +
				"Еженедельные статус-отчёты."},
		{"04", "Внедрение и сопровождение", "Deployment & Support",
			"Поэтапный ввод в эксплуатацию (blue-green / canary deployment). " +
				"Передача знаний команде заказчика. Гарантийная поддержка 3 месяца. " +
				"Опциональное долгосрочное сопровождение."},
	}

	for _, ph := range phases {
		// Phase number + name
		p = ok(doc.AddParagraph(""))
		must(p.ParagraphFormat().SetSpaceBefore(ip(docx.Pt(10).Twips())))
		must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(2).Twips())))
		must(p.ParagraphFormat().SetKeepWithNext(bp(true)))

		r := ok(p.AddRun(ph.num + "  "))
		must(r.Font().SetBold(bp(true)))
		must(r.Font().SetSize(lp(docx.Pt(14))))
		must(r.Font().Color().SetRGB(rgb(cBlue)))

		r = ok(p.AddRun(ph.name))
		must(r.Font().SetBold(bp(true)))
		must(r.Font().SetSize(lp(docx.Pt(11))))
		must(r.Font().Color().SetRGB(rgb(cNavy)))

		r = ok(p.AddRun("  /  " + ph.nameEn))
		must(r.Font().SetSize(lp(docx.Pt(11))))
		must(r.Font().SetItalic(bp(true)))
		must(r.Font().Color().SetRGB(rgb(cGrayLt)))

		// Description
		p = ok(doc.AddParagraph(""))
		must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
		must(p.ParagraphFormat().SetLeftIndent(ip(docx.Cm(0.7).Twips())))
		must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(6).Twips())))
		r = ok(p.AddRun(ph.desc))
		must(r.Font().SetSize(lp(docx.Pt(10))))
	}

	emptyN(doc, 1)
	ok(doc.AddParagraph(
		"«Мы не просто пишем код — мы решаем бизнес-задачи с помощью технологий» / "+
			"\"We don't just write code — we solve business problems through technology\"",
		docx.StyleName("Quote Block"),
	))
}

// ---------------------------------------------------------------------------
// Page 6: Timeline (landscape)
// ---------------------------------------------------------------------------

func buildTimelineSection(doc *docx.Document) {
	sect := ok(doc.AddSection(enum.WdSectionStartNewPage))
	must(sect.SetOrientation(enum.WdOrientationLandscape))
	must(sect.SetPageWidth(ip(docx.Mm(297).Twips())))
	must(sect.SetPageHeight(ip(docx.Mm(210).Twips())))
	must(sect.SetLeftMargin(ip(docx.Cm(2).Twips())))
	must(sect.SetRightMargin(ip(docx.Cm(2).Twips())))
	must(sect.SetTopMargin(ip(docx.Cm(2).Twips())))
	must(sect.SetBottomMargin(ip(docx.Cm(1.5).Twips())))

	ok(doc.AddHeading("Дорожная карта проекта / Project Roadmap", 1))

	p := ok(doc.AddParagraph(
		"Ориентировочный график реализации проекта. " +
			"Конкретные сроки уточняются на этапе Discovery по результатам анализа требований.",
	))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(10).Twips())))

	tbl := ok(doc.AddTable(7, 7, docx.StyleName("Table Grid")))
	must(tbl.SetAlignment(taPtr(enum.WdTableAlignmentCenter)))

	setTableHeader(tbl, 0, []string{
		"Фаза / Phase", "Q1", "Q2", "Q3", "Q4",
		"Результат / Deliverable", "Ответственный / Owner",
	})

	type phase struct {
		name        string
		q           [4]string
		deliverable string
		owner       string
	}
	phases := []phase{
		{"Discovery\nИсследование",
			[4]string{"Active", "", "", ""},
			"SRS, SAD, Risk Register",
			"Business Analysts"},
		{"Design\nПроектирование",
			[4]string{"Prep", "Active", "", ""},
			"API Contracts, UI/UX Prototype",
			"Architects"},
		{"Development\nРазработка",
			[4]string{"", "Start", "Active", "Active"},
			"Working Software (MVP + iterations)",
			"Engineering Team"},
		{"Testing\nТестирование",
			[4]string{"", "", "Active", "Active"},
			"Test Reports, Performance Benchmarks",
			"QA Engineers"},
		{"Deployment\nВнедрение",
			[4]string{"", "", "Prep", "Active"},
			"Production Release, Runbooks",
			"DevOps / SRE"},
		{"Support\nПоддержка",
			[4]string{"", "", "", "Start"},
			"SLA Reports, Knowledge Base",
			"Support Team"},
	}

	statusColors := map[string]string{
		"Active": cGreen,
		"Start":  cOrange,
		"Prep":   cSoftBlue,
	}

	for ri, ph := range phases {
		row := ri + 1
		ok(tbl.CellAt(row, 0)).SetText(ph.name)
		must(ok(tbl.CellAt(row, 0)).SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))

		for qi := range 4 {
			c := ok(tbl.CellAt(row, qi+1))
			if ph.q[qi] != "" {
				ps := c.Paragraphs()
				if len(ps) > 0 {
					must(ps[0].SetAlignment(ap(enum.WdParagraphAlignmentCenter)))
					run := ok(ps[0].AddRun(ph.q[qi]))
					must(run.Font().SetBold(bp(true)))
					must(run.Font().SetSize(lp(docx.Pt(8.5))))
					if col, found := statusColors[ph.q[qi]]; found {
						must(run.Font().Color().SetRGB(rgb(col)))
					}
				}
			}
			must(c.SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
		}

		ok(tbl.CellAt(row, 5)).SetText(ph.deliverable)
		ok(tbl.CellAt(row, 6)).SetText(ph.owner)
		must(ok(tbl.CellAt(row, 5)).SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
		must(ok(tbl.CellAt(row, 6)).SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
	}

	for ri := range 7 {
		row := ok(tbl.Rows().Get(ri))
		must(row.SetHeight(ip(docx.Cm(1.0).Twips())))
		must(row.SetHeightRule(rhPtr(enum.WdRowHeightRuleAtLeast)))
	}

	emptyN(doc, 1)
	p = ok(doc.AddParagraph(""))
	r := ok(p.AddRun("Обозначения: "))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().SetSize(lp(docx.Pt(9))))
	r = ok(p.AddRun("Active"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGreen)))
	must(r.Font().SetSize(lp(docx.Pt(9))))
	ok(p.AddRun(" — основная работа; "))
	r = ok(p.AddRun("Start"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cOrange)))
	must(r.Font().SetSize(lp(docx.Pt(9))))
	ok(p.AddRun(" — начало; "))
	r = ok(p.AddRun("Prep"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cSoftBlue)))
	must(r.Font().SetSize(lp(docx.Pt(9))))
	ok(p.AddRun(" — подготовка"))
}

// ---------------------------------------------------------------------------
// Page 7: Terms
// ---------------------------------------------------------------------------

func buildTermsSection(doc *docx.Document) {
	sect := ok(doc.AddSection(enum.WdSectionStartNewPage))
	must(sect.SetOrientation(enum.WdOrientationPortrait))
	must(sect.SetPageWidth(ip(docx.Mm(210).Twips())))
	must(sect.SetPageHeight(ip(docx.Mm(297).Twips())))
	must(sect.SetLeftMargin(ip(docx.Cm(2.5).Twips())))
	must(sect.SetRightMargin(ip(docx.Cm(2).Twips())))

	ok(doc.AddHeading("Условия сотрудничества / Terms & Conditions", 1))

	terms := []struct{ title, body string }{
		{"Срок действия / Validity",
			"Настоящее предложение действительно в течение 30 (тридцати) календарных дней " +
				"с даты составления документа. По истечении срока условия подлежат пересмотру."},
		{"Порядок оплаты / Payment Terms",
			"Предоплата 50% от стоимости этапа в течение 5 рабочих дней после подписания " +
				"договора. Оставшиеся 50% — в течение 10 рабочих дней после приёмки этапа. " +
				"При годовом контракте предоставляется скидка 15%."},
		{"Валюта расчётов / Currency",
			"Расчёты производятся в долларах США (USD) или евро (EUR) по курсу ЦБ РФ " +
				"на дату выставления счёта. Допускается расчёт в рублях по предварительному согласованию."},
		{"Конфиденциальность / Confidentiality",
			"Стороны обязуются не разглашать коммерческие, технические и финансовые условия " +
				"сотрудничества без предварительного письменного согласия другой стороны. " +
				"Срок действия обязательств — 3 года с момента подписания NDA."},
		{"Интеллектуальная собственность / IP Rights",
			"Все результаты интеллектуальной деятельности, созданные в ходе проекта, " +
				"переходят в полную собственность Заказчика после завершения окончательного расчёта. " +
				"Исполнитель сохраняет право на использование общих компонентов и библиотек."},
		{"Гарантии качества / Quality Assurance",
			"Гарантийный период составляет 3 месяца с момента приёмки. В течение этого срока " +
				"устранение дефектов выполняется безвозмездно. SLA для Enterprise-тарифа: " +
				"доступность 99.9%, время реакции на критический инцидент — не более 15 минут."},
		{"Форс-мажор / Force Majeure",
			"Стороны освобождаются от ответственности за неисполнение обязательств в случае " +
				"обстоятельств непреодолимой силы. Уведомление — не позднее 3 рабочих дней."},
	}

	for i, t := range terms {
		p := ok(doc.AddParagraph(""))
		must(p.ParagraphFormat().SetSpaceBefore(ip(docx.Pt(8).Twips())))
		must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(2).Twips())))
		must(p.ParagraphFormat().SetKeepWithNext(bp(true)))

		r := ok(p.AddRun(fmt.Sprintf("%d. ", i+1)))
		must(r.Font().SetBold(bp(true)))
		must(r.Font().Color().SetRGB(rgb(cNavy)))
		must(r.Font().SetSize(lp(docx.Pt(10.5))))

		r = ok(p.AddRun(t.title))
		must(r.Font().SetBold(bp(true)))
		must(r.Font().SetSize(lp(docx.Pt(10.5))))

		p = ok(doc.AddParagraph(""))
		must(p.SetAlignment(ap(enum.WdParagraphAlignmentJustify)))
		must(p.ParagraphFormat().SetLeftIndent(ip(docx.Cm(0.5).Twips())))
		must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(4).Twips())))
		r = ok(p.AddRun(t.body))
		must(r.Font().SetSize(lp(docx.Pt(10))))
	}
}

// ---------------------------------------------------------------------------
// Page 8: Contact & Signature
// ---------------------------------------------------------------------------

func buildContactSection(doc *docx.Document) {
	ok(doc.AddHeading("Контактная информация / Contact Information", 1))

	contacts := []struct{ label, value string }{
		{"Компания", "Example Solutions LLC / ООО «НоваТех Солюшенс»"},
		{"ИНН / Tax ID", "7712345678"},
		{"Генеральный директор", "Алексей Владимирович Петров"},
		{"Телефон", "+7 (495) 123-45-67"},
		{"Email", "proposal@Example.solutions"},
		{"Сайт / Website", "www.Example.solutions"},
		{"Юридический адрес", "123456, г. Москва, ул. Технологическая, д. 42, оф. 301"},
	}

	tabPos := docx.Cm(5).Twips()
	for _, c := range contacts {
		p := ok(doc.AddParagraph(""))
		must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(3).Twips())))
		ok(p.ParagraphFormat().TabStops().AddTabStop(
			tabPos,
			enum.WdTabAlignmentLeft,
			enum.WdTabLeaderSpaces,
		))

		r := ok(p.AddRun(c.label))
		must(r.Font().SetBold(bp(true)))
		must(r.Font().SetSize(lp(docx.Pt(10))))
		must(r.Font().Color().SetRGB(rgb(cNavy)))

		r = ok(p.AddRun(""))
		r.AddTab()

		r = ok(p.AddRun(c.value))
		must(r.Font().SetSize(lp(docx.Pt(10))))
	}

	// Signature block
	emptyN(doc, 4)

	p := ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentRight)))
	must(p.ParagraphFormat().SetKeepTogether(bp(true)))

	r := ok(p.AddRun("С уважением / Best regards,"))
	must(r.Font().SetItalic(bp(true)))
	must(r.Font().SetSize(lp(docx.Pt(10))))
	must(r.AddBreak(enum.WdBreakTypeLine))

	emptyN(doc, 1)

	p = ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(enum.WdParagraphAlignmentRight)))

	r = ok(p.AddRun("____________________________"))
	must(r.Font().SetSize(lp(docx.Pt(10))))
	must(r.Font().Color().SetRGB(rgb(cGrayLt)))
	must(r.AddBreak(enum.WdBreakTypeLine))

	r = ok(p.AddRun("Алексей Владимирович Петров"))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().SetSize(lp(docx.Pt(11))))
	must(r.AddBreak(enum.WdBreakTypeLine))

	r = ok(p.AddRun("Генеральный директор / CEO"))
	must(r.Font().SetSize(lp(docx.Pt(9))))
	must(r.Font().Color().SetRGB(rgb(cGray)))
	must(r.AddBreak(enum.WdBreakTypeLine))

	r = ok(p.AddRun("Example Solutions LLC"))
	must(r.Font().SetSize(lp(docx.Pt(9))))
	must(r.Font().SetSmallCaps(bp(true)))
	must(r.Font().Color().SetRGB(rgb(cGray)))
}

// ---------------------------------------------------------------------------
// Logo generation
// ---------------------------------------------------------------------------

func generateLogo(w, h int) *bytes.Reader {
	img := image.NewRGBA(image.Rect(0, 0, w, h))

	from := [3]float64{27, 58, 92}
	to := [3]float64{41, 128, 185}

	for x := range w {
		t := float64(x) / float64(w-1)
		cr := uint8(from[0]*(1-t) + to[0]*t)
		cg := uint8(from[1]*(1-t) + to[1]*t)
		cb := uint8(from[2]*(1-t) + to[2]*t)
		col := color.RGBA{R: cr, G: cg, B: cb, A: 255}
		for y := range h {
			img.Set(x, y, col)
		}
	}

	// Subtle accent lines
	white := color.RGBA{R: 255, G: 255, B: 255, A: 100}
	for x := w / 6; x < w*5/6; x++ {
		for _, y := range []int{h/3 - 1, h / 3, h/3 + 1} {
			img.Set(x, y, white)
		}
	}

	var buf bytes.Buffer
	_ = png.Encode(&buf, img)
	return bytes.NewReader(buf.Bytes())
}

// ---------------------------------------------------------------------------
// Reusable formatting helpers
// ---------------------------------------------------------------------------

func addFactRow(doc *docx.Document, label, value string) {
	p := ok(doc.AddParagraph(""))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(2).Twips())))
	ok(p.ParagraphFormat().TabStops().AddTabStop(
		docx.Cm(8).Twips(),
		enum.WdTabAlignmentLeft,
		enum.WdTabLeaderDots,
	))

	r := ok(p.AddRun(label))
	must(r.Font().SetBold(bp(true)))
	must(r.Font().SetSize(lp(docx.Pt(10))))
	must(r.Font().Color().SetRGB(rgb(cNavy)))

	r = ok(p.AddRun(""))
	r.AddTab()

	r = ok(p.AddRun(value))
	must(r.Font().SetSize(lp(docx.Pt(10))))
}

func addStyledParagraph(doc *docx.Document, text string,
	align enum.WdParagraphAlignment, font string, size float64,
	bold, italic bool, clr string,
) {
	p := ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(align)))
	r := ok(p.AddRun(text))
	must(r.Font().SetName(sp(font)))
	must(r.Font().SetSize(lp(docx.Pt(size))))
	must(r.Font().SetBold(bp(bold)))
	must(r.Font().SetItalic(bp(italic)))
	must(r.Font().Color().SetRGB(rgb(clr)))
}

func addSimpleParagraph(doc *docx.Document, text string, size float64,
	bold bool, clr string, align enum.WdParagraphAlignment,
) {
	p := ok(doc.AddParagraph(""))
	must(p.SetAlignment(ap(align)))
	must(p.ParagraphFormat().SetSpaceAfter(ip(docx.Pt(2).Twips())))
	r := ok(p.AddRun(text))
	must(r.Font().SetSize(lp(docx.Pt(size))))
	must(r.Font().SetBold(bp(bold)))
	must(r.Font().Color().SetRGB(rgb(clr)))
}

func setTableHeader(tbl *docx.Table, rowIdx int, headers []string) {
	for ci, h := range headers {
		c := ok(tbl.CellAt(rowIdx, ci))
		c.SetText(h)
		must(c.SetVerticalAlignment(vaPtr(enum.WdCellVerticalAlignmentCenter)))
		for _, p := range c.Paragraphs() {
			for _, r := range p.Runs() {
				must(r.Font().SetBold(bp(true)))
				must(r.Font().SetSize(lp(docx.Pt(9))))
				must(r.Font().Color().SetRGB(rgb(cNavy)))
			}
		}
	}
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

func emptyN(doc *docx.Document, n int) {
	for range n {
		ok(doc.AddParagraph(""))
	}
}

func outputPath(name string) string {
	_, src, _, _ := runtime.Caller(0)
	dir := filepath.Join(filepath.Dir(src), "out")
	_ = os.MkdirAll(dir, 0o755)
	return filepath.Join(dir, name)
}

func must(err error) {
	if err != nil {
		log.Fatal(err)
	}
}
func ok[T any](v T, err error) T { must(err); return v }

func ip(v int) *int                 { return &v }
func sp(v string) *string           { return &v }
func bp(v bool) *bool               { return &v }
func lp(v docx.Length) *docx.Length { return &v }

func ap(v enum.WdParagraphAlignment) *enum.WdParagraphAlignment          { return &v }
func taPtr(v enum.WdTableAlignment) *enum.WdTableAlignment               { return &v }
func hlPtr(v enum.WdColorIndex) *enum.WdColorIndex                       { return &v }
func rhPtr(v enum.WdRowHeightRule) *enum.WdRowHeightRule                 { return &v }
func vaPtr(v enum.WdCellVerticalAlignment) *enum.WdCellVerticalAlignment { return &v }

func rgb(hex string) *docx.RGBColor {
	c, err := docx.RGBColorFromString(hex)
	must(err)
	return &c
}
