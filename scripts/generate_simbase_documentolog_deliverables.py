#!/usr/bin/env python3
"""Generate Word TZ and executive PowerPoint for SimBASE ↔ Documentolog integration."""

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt as PptPt

OUT_DIR = Path("/workspace/deliverables")
OUT_DIR.mkdir(parents=True, exist_ok=True)
TODAY = date(2026, 9, 10).strftime("%d.%m.%Y")


def set_doc_styles(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    for level in range(1, 4):
        h = doc.styles[f"Heading {level}"]
        h.font.name = "Times New Roman"
        h.font.bold = True
        h.font.color.rgb = RGBColor(0, 51, 102)


def add_title_page(doc: Document) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "ТЕХНИЧЕСКОЕ ЗАДАНИЕ\n\n"
        "на выполнение интеграционных работ\n"
        "между платформами SimBASE и Documentolog\n\n"
        "(входящая и исходящая корреспонденция с внешним контуром)"
    )
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "Times New Roman"

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mrun = meta.add_run(f"\nВерсия: 1.1\nДата: {TODAY}\nСтатус: Проект\n")
    mrun.font.size = Pt(12)
    mrun.font.name = "Times New Roman"

    doc.add_page_break()


def add_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for p in hdr_cells[i].paragraphs:
            for r in p.runs:
                r.bold = True
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            table.rows[ri + 1].cells[ci].text = val
    doc.add_paragraph()


def build_word_tz() -> Path:
    doc = Document()
    set_doc_styles(doc)
    add_title_page(doc)

    sections = [
        (
            "1. Общие положения",
            [
                (
                    "1.1. Наименование работ",
                    "Разработка и внедрение интеграционного решения «SimBASE ↔ Documentolog» "
                    "для автоматизации входящей и исходящей официальной корреспонденции с внешними "
                    "контрагентами, включая государственные органы Республики Казахстан.",
                ),
                (
                    "1.2. Основание для разработки",
                    "Необходимость устранения дублирования делопроизводственных процедур между "
                    "SimBASE (внутренние процессы) и Documentolog (внешний обмен через шлюзы "
                    "ЕСЭДО и КЦОЭД). Основание: _________________________ (договор/приказ заказчика).",
                ),
                (
                    "1.3. Заказчик и исполнитель",
                    "Заказчик: _________________________\n"
                    "Исполнитель: _________________________\n"
                    "Владелец платформы SimBASE: Simourg LTD / партнёр _________________________\n"
                    "Поставщик Documentolog: Documentolog Business",
                ),
                (
                    "1.4. Область применения",
                    "Только входящая и исходящая корреспонденция с внешним контуром. "
                    "Внутренний документооборот SimBASE, кадровые и прочие процессы без внешней "
                    "отправки — вне рамок настоящего ТЗ.",
                ),
            ],
        ),
        (
            "2. Назначение и цели проекта",
            [
                (
                    "2.1. Назначение",
                    "Обеспечить единое рабочее место сотрудников в SimBASE при сохранении "
                    "Documentolog в роли юридически значимого шлюза обмена с внешним контуром "
                    "(ЕСЭДО, КЦОЭД, контрагенты).",
                ),
                (
                    "2.2. Цели",
                    "• Исключить повторное согласование и регистрацию одного и того же письма в двух системах.\n"
                    "• Сократить время обработки входящей и исходящей корреспонденции.\n"
                    "• Снизить операционные ошибки при переносе вложений и реквизитов.\n"
                    "• Обеспечить прозрачный контроль статусов отправки и исполнения в SimBASE.",
                ),
                (
                    "2.3. Ограничение по лицензиям",
                    "Имеется 33 лицензии Documentolog. Целевая модель: техническая учётная запись "
                    "для интеграции; персональные лицензии — только при юридической необходимости "
                    "действий от имени конкретного лица в Documentolog. Требуется согласование "
                    "с Documentolog возможности отправки в ЕСЭДО через API без рутинного входа "
                    "пользователей в UI.",
                ),
            ],
        ),
        (
            "3. Описание ситуации AS-IS",
            [
                (
                    "3.1. Роли систем",
                    "SimBASE — внутренние документы и внутренние бизнес-процессы.\n"
                    "Documentolog — только внешняя переписка (контрагенты, госорганы); встроенные "
                    "шлюзы ЕСЭДО и КЦОЭД.",
                ),
                (
                    "3.2. Причина разделения",
                    "SimBASE не имеет шлюзов ЕСЭДО/КЦОЭД; Documentolog — имеет. Внешняя "
                    "корреспонденция ведётся в Documentolog, внутренняя — в SimBASE.",
                ),
                (
                    "3.3. Исходящая корреспонденция (AS-IS)",
                    "1) Подготовка и согласование письма в SimBASE.\n"
                    "2) Перенос (вложение) в Documentolog.\n"
                    "3) Повторное согласование, подписание, регистрация в Documentolog.\n"
                    "4) Отправка контрагенту / через ЕСЭДО.",
                ),
                (
                    "3.4. Входящая корреспонденция (AS-IS)",
                    "1) Письмо поступает в Documentolog (ЕСЭДО/внешний канал).\n"
                    "2) Офис-менеджер регистрирует в Documentolog.\n"
                    "3) Вкладывает письмо в SimBASE.\n"
                    "4) Согласование в SimBASE, направление ответственному подразделению.",
                ),
                (
                    "3.5. Последствия",
                    "Дублирование согласований; риск расхождения реквизитов и вложений; "
                    "двойная нагрузка на делопроизводство; отсутствие единой картины статуса.",
                ),
            ],
        ),
        (
            "4. Целевое состояние TO-BE",
            [
                (
                    "4.1. Принципы",
                    "• SimBASE — единая точка работы пользователей по Вх/Исх корреспонденции.\n"
                    "• Documentolog — backend для юридически значимого обмена с внешним контуром.\n"
                    "• Согласование и маршрутизация — только в SimBASE.\n"
                    "• Повторное согласование в Documentolog не выполняется.",
                ),
                (
                    "4.2. Входящая корреспонденция (TO-BE)",
                    "1) Письмо принимается Documentolog (как сейчас).\n"
                    "2) Офис-менеджер в SimBASE нажимает «Запросить входящие» (опционально — "
                    "автоматический опрос по расписанию).\n"
                    "3) Интеграция забирает новые письма, метаданные и вложения в SimBASE.\n"
                    "4) Запускается БП: регистрация → согласование → ответственное подразделение.\n"
                    "5) Статусы внешнего контура отображаются в карточке SimBASE.",
                ),
                (
                    "4.3. Исходящая корреспонденция (TO-BE)",
                    "1) Подготовка, согласование, подписание и регистрация — в SimBASE.\n"
                    "2) По кнопке «Отправить через Documentolog» — передача в Documentolog и "
                    "дальнейшая отправка контрагенту / через ЕСЭДО.\n"
                    "3) Мониторинг статусов (отправлено, доставлено, зарегистрировано у адресата) — в SimBASE.",
                ),
            ],
        ),
        (
            "5. Функциональные требования",
            [],
        ),
    ]

    for title, subs in sections:
        doc.add_heading(title, level=1)
        for sub_title, body in subs:
            doc.add_heading(sub_title, level=2)
            doc.add_paragraph(body)

    doc.add_heading("5.1. Входящая корреспонденция", level=2)
    add_table(
        doc,
        ["ID", "Требование", "Критерий"],
        [
            ["FR-IN-01", "Кнопка «Запросить входящие из Documentolog»", "Доступ по роли офис-менеджера"],
            ["FR-IN-02", "Импорт только новых писем", "Идемпотентность по documentolog_id"],
            ["FR-IN-03", "Перенос метаданных и вложений", "Полнота данных в тесте"],
            ["FR-IN-04", "Автозапуск БП регистрация → согласование → исполнитель", "По утверждённой схеме"],
            ["FR-IN-05", "Связь SimBASE ↔ Documentolog", "Внешний ID в карточке"],
            ["FR-IN-06", "Плановый опрос (опционально)", "Настраиваемый интервал"],
        ],
    )

    doc.add_heading("5.2. Исходящая корреспонденция", level=2)
    add_table(
        doc,
        ["ID", "Требование", "Критерий"],
        [
            ["FR-OUT-01", "БП без этапов согласования в Documentolog", "Нет дублирующих задач"],
            ["FR-OUT-02", "Валидация перед отправкой", "Блокировка при ошибках"],
            ["FR-OUT-03", "Отправка после статуса «Готов к отправке»", "Контроль прав"],
            ["FR-OUT-04", "Передача в Documentolog (ЕСЭДО/B2B)", "Успешный ответ API"],
            ["FR-OUT-05", "Фиксация статусов отправки", "Webhook/опрос"],
            ["FR-OUT-06", "Запрет повторной отправки без явного действия", "Контроль версии"],
        ],
    )

    doc.add_heading("6. Техническая архитектура", level=1)
    doc.add_paragraph(
        "Компоненты: SimBASE (БП, формы, хранилище файлов); интеграционный сервис (REST, "
        "webhook, адаптер Documentolog); Documentolog Business (ЕСЭДО/КЦОЭД, внешний обмен).\n\n"
        "Критическая зависимость: подтверждение у Documentolog API для импорта входящей "
        "корреспонденции и отправки в ЕСЭДО без рутинной работы пользователей в UI. "
        "Публичная документация API (apibusiness.documentolog.com) ориентирована на подписание; "
        "требуется проектная спецификация методов для канцелярии."
    )

    doc.add_heading("7. Этапы работ", level=1)
    add_table(
        doc,
        ["Этап", "Содержание", "Результат"],
        [
            ["0", "Обследование", "Отчёт, ТЗ v1.2"],
            ["1", "Workshop с Documentolog (API ЕСЭДО/входящие)", "Спецификация API"],
            ["2", "Проектирование", "Проектное решение"],
            ["3", "Разработка интеграции", "Стенд DEV"],
            ["4", "Настройка SimBASE (2 БП)", "TEST"],
            ["5", "Пилот", "Протокол ОПЭ"],
            ["6", "Промышленная эксплуатация", "Акт, регламент"],
        ],
    )

    doc.add_heading("8. Критерии приёмки", level=1)
    doc.add_paragraph(
        "1. Входящее: письмо из Documentolog импортируется в SimBASE по кнопке офис-менеджера, "
        "проходит регистрацию и согласование без ручного дублирования в Documentolog.\n"
        "2. Исходящее: согласование и подписание только в SimBASE; отправка через Documentolog "
        "без второго цикла согласования в Documentolog.\n"
        "3. Статусы отправки видны в карточке SimBASE.\n"
        "4. Журнал интеграции восстанавливает цепочку по одному письму.\n"
        "5. Подтверждение делопроизводства по результатам пилота (количество кейсов — согласовать)."
    )

    doc.add_heading("9. Риски", level=1)
    add_table(
        doc,
        ["Риск", "Мера"],
        [
            ["Нет API «забрать входящие»", "Запрос в Documentolog; временный полуавтомат"],
            ["ЕСЭДО только из UI Documentolog", "Расширенный API от Documentolog"],
            ["Двойная регистрация номеров", "Учётный номер — SimBASE; Documentolog — транспорт"],
            ["33 лицензии", "Техучётка + маппинг подписантов"],
        ],
    )

    doc.add_heading("10. Открытые вопросы", level=1)
    doc.add_paragraph(
        "• Объёмы Вх/Исх в день и доля ЕСЭДО.\n"
        "• Состав подписантов исходящих в ГО.\n"
        "• Автоопрос входящих vs только кнопка.\n"
        "• Доступность файлов SimBASE для Documentolog по HTTPS.\n"
        "• Контакт Documentolog для проектного API.\n"
        "• Партнёр SimBASE и сроки настройки БП."
    )

    doc.add_heading("Приложение А. Сравнение AS-IS и TO-BE", level=1)
    add_table(
        doc,
        ["Параметр", "AS-IS", "TO-BE"],
        [
            ["Рабочее место", "SimBASE + Documentolog", "Только SimBASE"],
            ["Согласование", "Дважды", "Один раз (SimBASE)"],
            ["ЕСЭДО/КЦОЭД", "Documentolog", "Без изменений (под капотом)"],
            ["Входящие", "Ручная регистрация в обеих системах", "«Запросить» → БП в SimBASE"],
            ["Исходящие", "Согласование в обеих системах", "SimBASE → «Отправить» в Documentolog"],
        ],
    )

    out = OUT_DIR / "TZ_SimBASE_Documentolog_Vh_Ish_korrespondenciya.docx"
    doc.save(out)
    return out


def add_slide_title(prs: Presentation, title: str, subtitle: str = "") -> None:
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    if subtitle and len(slide.placeholders) > 1:
        slide.placeholders[1].text = subtitle


def add_slide_bullets(prs: Presentation, title: str, bullets: list[str]) -> None:
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    tf = slide.shapes.placeholders[1].text_frame
    tf.clear()
    tf.word_wrap = True
    for i, text in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.level = 0
        p.font.size = PptPt(20)
        p.font.name = "Calibri"


def add_slide_two_column(prs: Presentation, title: str, left_title: str, left: list[str], right_title: str, right: list[str]) -> None:
    layout = prs.slide_layouts[5]  # blank
    slide = prs.slides.add_slide(layout)
    # title
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    box.text_frame.text = title
    box.text_frame.paragraphs[0].font.size = PptPt(32)
    box.text_frame.paragraphs[0].font.bold = True
    box.text_frame.paragraphs[0].font.color.rgb = PptRGB(0, 51, 102)

    def col(x, y, w, h, head, items):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = head
        p.font.bold = True
        p.font.size = PptPt(22)
        p.font.color.rgb = PptRGB(0, 102, 153)
        for it in items:
            bp = tf.add_paragraph()
            bp.text = it
            bp.font.size = PptPt(18)
            bp.level = 0

    add_slide_two_column_col = col
    add_slide_two_column_col(Inches(0.5), Inches(1.2), Inches(4.3), Inches(5.5), left_title, left)
    add_slide_two_column_col(Inches(5.0), Inches(1.2), Inches(4.3), Inches(5.5), right_title, right)


def build_ppt() -> Path:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    add_slide_title(
        prs,
        "Интеграция SimBASE и Documentolog",
        f"Единый процесс входящей и исходящей корреспонденции\n{TODAY}",
    )

    add_slide_bullets(
        prs,
        "Зачем это нужно?",
        [
            "Сегодня сотрудники ведут одно письмо в двух системах: SimBASE и Documentolog.",
            "SimBASE — вся внутренняя работа; Documentolog — единственный канал к внешнему миру (ЕСЭДО, КЦОЭД).",
            "Из-за отсутствия шлюза ЕСЭДО в SimBASE внешняя переписка «застряла» в Documentolog.",
            "Повторное согласование и регистрация отнимают время и создают ошибки.",
        ],
    )

    add_slide_bullets(
        prs,
        "Как работает сейчас (AS-IS)",
        [
            "Исходящее: согласовали в SimBASE → вручную в Documentolog → снова согласование → отправка.",
            "Входящее: пришло в Documentolog → регистрация → вручную в SimBASE → снова согласование.",
            "Офис-менеджеры и инициаторы постоянно переключаются между системами.",
            "Нет единой картины: где письмо, на каком этапе, кто ответственный.",
        ],
    )

    add_slide_bullets(
        prs,
        "Что предлагаем (TO-BE)",
        [
            "SimBASE — единое рабочее место для всех по входящей и исходящей корреспонденции.",
            "Documentolog остаётся: это «двигатель» обмена с госорганами и контрагентами.",
            "Входящие: кнопка «Запросить» → письма попадают в SimBASE → один регламентный БП.",
            "Исходящие: полный цикл в SimBASE → кнопка «Отправить» → Documentolog уходит во внешку.",
        ],
    )

    add_slide_two_column(
        prs,
        "Что это даст организации",
        "Эффект для бизнеса",
        [
            "Меньше ручного труда делопроизводства",
            "Быстрее прохождение писем",
            "Меньше ошибок при переносе файлов и реквизитов",
            "Прозрачный контроль статусов в одной системе",
            "Сотрудники не заходят в Documentolog в рутине",
        ],
        "Что не меняется",
        [
            "33 лицензии Documentolog — используем как транспорт",
            "Юридическая значимость и ЕСЭДО — через Documentolog",
            "Внутренние процессы SimBASE — без изменений",
            "Отказ от Documentolog не требуется",
        ],
    )

    add_slide_bullets(
        prs,
        "Ключевые сценарии",
        [
            "Входящее: Documentolog принял письмо → офис-менеджер «Запросить» → регистрация и маршрут в SimBASE.",
            "Исходящее: согласование и подписание в SimBASE → «Отправить через Documentolog» → мониторинг статуса.",
            "Повторное согласование в Documentolog исключается по дизайну.",
        ],
    )

    add_slide_bullets(
        prs,
        "Риски и как их снимаем",
        [
            "API Documentolog для входящих и ЕСЭДО — согласуем на старте с вендором (workshop).",
            "Двойные номера — учётный номер в SimBASE, Documentolog только для транспорта.",
            "Подписание ЭЦП — по возможности из SimBASE (iframe/API), без второго маршрута.",
            "Пилот на ограниченном потоке перед полным отключением рутинного UI Documentolog.",
        ],
    )

    add_slide_bullets(
        prs,
        "Этапы и решение для руководства",
        [
            "Этап 0–1: обследование + согласование API с Documentolog (2–4 недели — уточнить с исполнителем).",
            "Этап 2–4: проектирование, разработка интеграции, настройка двух БП в SimBASE.",
            "Этап 5: пилот с делопроизводством и офис-менеджментом.",
            "Решение: утвердить концепцию TO-BE, назначить владельца процесса и выделить рабочую группу (ИТ, канцелярия, Documentolog, SimBASE).",
        ],
    )

    add_slide_title(
        prs,
        "Итог",
        "Одна система для людей — SimBASE.\n"
        "Один канал во внешний мир — Documentolog.\n"
        "Один раз согласуем — один раз отправим.",
    )

    out = OUT_DIR / "Prezentaciya_SimBASE_Documentolog_dlya_rukovodstva.pptx"
    prs.save(out)
    return out


if __name__ == "__main__":
    w = build_word_tz()
    p = build_ppt()
    print(f"Word: {w}")
    print(f"PPT:  {p}")
