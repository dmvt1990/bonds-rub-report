"""
Слайд «Рублевые облигации — ОФЗ-ИН (линкеры)».

Структура повторяет bond_table_ofz.py. Отличия — в колонках:

  * «Индекс. номинал, ₽» — опубликованный MOEX (`facevalue`), а не наш
    расчёт: именно его клиент видит в выписке;
  * «Реальн. дох-ть» — `yield_close`; у линкеров доходность MOEX и есть
    реальная (сверено с собственным расчётом, расхождение ≤0.01 п.п.);
  * «Индексация до DD.MM» — прирост номинала по уже опубликованному ИПЦ.
    Этого числа нет ни у MOEX, ни у Минфина;
  * «Вмен. инфл.» — breakeven против кривой ОФЗ-ПД. Пока только для
    внутреннего использования, см. INCLUDE_BREAKEVEN.
"""
from datetime import date as _date
from typing import List, Optional

from pptx.enum.text import PP_ALIGN
from pptx.presentation import Presentation as PresentationType
from pptx.util import Cm

from ..constants import (
    FONT_PRIMARY,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    TABLE_SLIDE_TITLE_FONT_SIZE,
    TABLE_SLIDE_LEFT, TABLE_SLIDE_TOP,
    TABLE_SLIDE_SOURCE_TOP, TABLE_SLIDE_SOURCE_FONT_SIZE,
    TABLE_SLIDE_FOOTNOTE_TOP, TABLE_SLIDE_FOOTNOTE_FONT_SIZE,
    TABLE_OFZ_IN_COL_WIDTHS_CM,
)
from bonds_report_core.formatting import (
    format_date_dmy, format_percent, format_number,
)
from bonds_report_core.models import Bond
from bonds_report_core.tables import TableColumn, render_bond_table
from bonds_report_core.fonts import set_run_font

from .. import ofz_in as oi

_TABLE_SLIDE_LAYOUT_INDEX = 3
SLIDE_TITLE = "Рублевые облигации"
SECTION_TITLE_OFZ_IN = "ОФЗ-ИН — с индексируемым на инфляцию номиналом"

# Вменённая инфляция считается всегда и пишется в лог, но на слайд пока не
# идёт: кривая breakeven круто растёт к длинному концу (≈5% к 2028 против
# ≈9.6% к 2033), и пока не отделён премий за ликвидность в длинных линкерах,
# показывать это клиенту как «ожидания рынка» нельзя. Включать осознанно.
INCLUDE_BREAKEVEN = False


def _columns(locked_until) -> List[TableColumn]:
    w = TABLE_OFZ_IN_COL_WIDTHS_CM
    cols = [
        TableColumn("№", w[0], align="center"),
        TableColumn("Выпуск", w[1], align="center"),
        TableColumn("ISIN", w[2], align="center"),
        TableColumn("Дата\nпогашения", w[3], align="center"),
        TableColumn("Купон,\n% год.", w[4], align="center"),
        TableColumn("Индекс.\nноминал, ₽", w[5], align="center"),
        TableColumn("Цена, %\nот ном.", w[6], align="center"),
        TableColumn("Реальн.\nдох-ть,\n% год.", w[7], align="center"),
        TableColumn(f"Индексация\nдо {locked_until:%d.%m}, %", w[8], align="center"),
    ]
    if INCLUDE_BREAKEVEN:
        cols.append(TableColumn("Вмен.\nинфл., %", w[9], align="center"))
    cols.append(TableColumn("В\nперечне", w[-1], align="center"))
    return cols


def _bond_to_row(i, b: Bond, highlighted, locked_pct, breakeven) -> List[str]:
    row = [
        str(i),
        b.shortname or "—",
        b.isin,
        format_date_dmy(b.matdate),
        format_percent(b.couponpercent, decimals=2),
        format_number(float(b.facevalue), decimals=2) if b.facevalue else "—",
        format_number(b.close, decimals=2),
        format_percent(b.yield_close, decimals=2),
        format_percent(locked_pct, decimals=2) if locked_pct is not None else "—",
    ]
    if INCLUDE_BREAKEVEN:
        row.append(format_percent(breakeven, decimals=2) if breakeven is not None else "—")
    row.append("✓" if b.isin in highlighted else "—")
    return row


def _fill_title_placeholder(slide, text: str) -> bool:
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = text
            set_run_font(run, FONT_PRIMARY)
            run.font.size = TABLE_SLIDE_TITLE_FONT_SIZE
            run.font.color.rgb = COLOR_TEXT_PRIMARY
            return True
    return False


def _blank_other_placeholders(slide) -> None:
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            continue
        try:
            tf = ph.text_frame
            tf.clear()
            tf.paragraphs[0].add_run().text = ""
        except Exception:
            pass


def _add_source_line(slide, report_date, locked_until) -> None:
    box = slide.shapes.add_textbox(Cm(18.00), TABLE_SLIDE_SOURCE_TOP, Cm(14.70), Cm(0.50))
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = (f"Источник: Cbonds, {format_date_dmy(report_date)}; "
                f"ИПЦ Росстата — номинал определён до {locked_until:%d.%m.%Y}")
    set_run_font(run, FONT_PRIMARY)
    run.font.size = TABLE_SLIDE_SOURCE_FONT_SIZE
    run.font.color.rgb = COLOR_TEXT_SECONDARY


def _add_footnote(slide) -> None:
    box = slide.shapes.add_textbox(Cm(1.13), TABLE_SLIDE_FOOTNOTE_TOP, Cm(31.60), Cm(0.40))
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = (
        "Номинал ОФЗ-ИН индексируется на индекс потребительских цен Росстата с лагом; "
        "колонка «Индексация до …» — прирост номинала по уже опубликованному ИПЦ, он "
        "не зависит от будущей инфляции. Доходность указана реальная, сверх индексации. "
        "Внимание! Приведенные котировки и расчеты являются индикативными и подлежат "
        "регулярному обновлению."
    )
    set_run_font(run, FONT_PRIMARY)
    run.font.size = TABLE_SLIDE_FOOTNOTE_FONT_SIZE
    run.font.color.rgb = COLOR_TEXT_MUTED


def render_ofz_in(
    prs: PresentationType,
    bonds: List[Bond],
    cpi,
    issue_dates: dict,
    nominal_curve=None,
    section_title: str = SECTION_TITLE_OFZ_IN,
    report_date=None,
    highlighted: Optional[set] = None,
) -> bool:
    """Добавляет слайд с таблицей ОФЗ-ИН. False — если ИПЦ недоступен."""
    if cpi is None:
        return False
    highlighted = highlighted or set()
    today = report_date or _date.today()
    locked_pct = cpi.locked_in_pct(today)

    layout = prs.slide_layouts[_TABLE_SLIDE_LAYOUT_INDEX]
    slide = prs.slides.add_slide(layout)
    _fill_title_placeholder(slide, SLIDE_TITLE)
    _blank_other_placeholders(slide)

    sorted_bonds = sorted(bonds, key=lambda b: (b.matdate or _date.max))
    rows = []
    for i, b in enumerate(sorted_bonds, start=1):
        issue = issue_dates.get(b.secid) or issue_dates.get(b.isin)
        if issue:
            oi.check_nominal(cpi, b, issue, today)
        be = oi.breakeven_pct(b.yield_close, b.matdate, nominal_curve)
        if be is not None:
            # В таблицу breakeven пока не идёт (INCLUDE_BREAKEVEN=False),
            # но в cron-лог пишем всегда — это и есть «внутреннее
            # использование», о котором речь в шапке модуля.
            print(f"ОФЗ-ИН {b.shortname or b.isin}: реальн. {b.yield_close:.2f}%, "
                  f"вмен. инфляция {be:.2f}%")
        rows.append(_bond_to_row(i, b, highlighted, locked_pct, be))

    render_bond_table(
        slide,
        left=TABLE_SLIDE_LEFT,
        top=TABLE_SLIDE_TOP,
        columns=_columns(cpi.locked_until),
        rows=rows,
        section_title=section_title,
    )
    _add_source_line(slide, today, cpi.locked_until)
    _add_footnote(slide)
    return True
