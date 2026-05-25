"""
Рендеринг таблицы облигаций в pptx.

Строит таблицу с корпоративным стилем Газпромбанка:
- шапка-цвет (TABLE_HEADER_BG)
- шапка-секция (опционально, TABLE_SECTION_BG)
- zebra-строки (чередование TABLE_ROW_EVEN_BG / TABLE_ROW_ODD_BG)
- без видимых границ между ячейками
- шрифт Cera CY

Модуль не знает про конкретные колонки таблицы ОФЗ — принимает готовый
список заголовков и строк. Форматирование чисел в строки — ответственность
вызывающего кода (обычно в слайдах).
"""
from dataclasses import dataclass
from typing import List, Optional

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Emu, Pt
from lxml import etree

from .constants import (
    FONT_PRIMARY,
    TABLE_HEADER_BG, TABLE_HEADER_TEXT_COLOR, TABLE_HEADER_FONT_SIZE, TABLE_HEADER_HEIGHT,
    TABLE_SECTION_BG, TABLE_SECTION_TEXT_COLOR, TABLE_SECTION_FONT_SIZE, TABLE_SECTION_HEIGHT,
    TABLE_ROW_EVEN_BG, TABLE_ROW_ODD_BG, TABLE_ROW_TEXT_COLOR, TABLE_ROW_FONT_SIZE, TABLE_ROW_HEIGHT,
)
from .fonts import set_run_font


@dataclass
class TableColumn:
    """Описание одной колонки таблицы."""
    header: str                  # заголовок, может содержать переносы \n
    width_cm: float              # ширина колонки в см
    align: str = "left"          # 'left' / 'center' / 'right'


def _set_cell_fill(cell, rgb: RGBColor) -> None:
    """Заливка фона ячейки."""
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb


def _remove_cell_borders(cell) -> None:
    """
    Убирает видимые границы ячейки. По умолчанию python-pptx рисует
    тонкие линии; нам нужна таблица без сетки, только с заливкой.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    # Добавляем пустые элементы для всех сторон — это говорит PowerPoint
    # "границ нет" (в отличие от отсутствия элементов, которое = "по умолчанию")
    for side in ("lnL", "lnR", "lnT", "lnB"):
        # убираем существующие
        for existing in tcPr.findall(qn(f"a:{side}")):
            tcPr.remove(existing)
        ln = etree.SubElement(tcPr, qn(f"a:{side}"))
        ln.set("w", "0")
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        noFill = etree.SubElement(ln, qn("a:noFill"))


def _style_cell(
    cell,
    text: str,
    *,
    font_size,
    font_color: RGBColor,
    bg_color: RGBColor,
    bold: bool = False,
    align: str = "left",
    anchor=MSO_ANCHOR.MIDDLE,
) -> None:
    """Заполнение и стилизация одной ячейки.

    ВАЖНО по шрифтам: в ячейках таблицы мы НЕ вызываем set_run_font()
    и не задаём run.font.name — шрифт наследуется из слайд-мастера
    корпоративного шаблона Газпромбанка. Там прописана пара шрифтов
    (Latin + Cyrillic), и PowerPoint корректно подбирает кириллицу.

    Если принудительно записать латинский шрифт (Arial, Calibri) через
    python-pptx, в XML попадёт только тэг <a:latin>, а PowerPoint на
    некоторых системах не сможет найти кириллический fallback и отобразит
    UTF-8 байты как Latin-1 («ÐÐ¤Ð» вместо «ОФЗ»). Поэтому шрифт не трогаем.
    """
    _set_cell_fill(cell, bg_color)
    _remove_cell_borders(cell)

    # Вертикальное выравнивание
    cell.vertical_anchor = anchor

    # Узкие внутренние отступы — тогда в таблицу вмещается больше данных
    cell.margin_left = Cm(0.08)
    cell.margin_right = Cm(0.08)
    cell.margin_top = Cm(0.02)
    cell.margin_bottom = Cm(0.02)

    tf = cell.text_frame
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.alignment = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }[align]

    # Если текст содержит '\n' — делим на несколько параграфов
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            para = p
        else:
            para = tf.add_paragraph()
            para.alignment = p.alignment
        run = para.add_run()
        run.text = line
        # Шрифт НЕ задаём — наследуется из master/theme
        run.font.size = font_size
        run.font.color.rgb = font_color
        run.font.bold = bold


def render_bond_table(
    slide,
    *,
    left: Emu,
    top: Emu,
    columns: List[TableColumn],
    rows: List[List[str]],
    section_title: Optional[str] = None,
):
    """
    Строит таблицу облигаций на слайде.

    Args:
        slide: слайд, куда рисуем
        left, top: координаты верхнего левого угла таблицы
        columns: список TableColumn с заголовками и ширинами
        rows: данные — список строк, каждая строка — список значений (str)
              по числу колонок
        section_title: если задано, перед шапкой с колонками добавляется
              цветная строка с этим текстом (например, "Суверенные
              облигации с фиксированным купоном")
    """
    n_cols = len(columns)
    if any(len(r) != n_cols for r in rows):
        raise ValueError(f"Число значений в строках должно быть {n_cols}")

    # Количество строк таблицы: шапка + [опционально секция] + данные
    n_table_rows = 1 + (1 if section_title else 0) + len(rows)

    # Общая ширина = сумма ширин колонок
    total_width_cm = sum(c.width_cm for c in columns)
    total_width = Cm(total_width_cm)

    # Высота всей таблицы (для инициализации — потом подкорректируем высоты строк)
    total_height = (
        TABLE_HEADER_HEIGHT
        + (TABLE_SECTION_HEIGHT if section_title else 0)
        + TABLE_ROW_HEIGHT * len(rows)
    )

    table_shape = slide.shapes.add_table(n_table_rows, n_cols, left, top, total_width, total_height)
    tbl = table_shape.table

    # Ширины колонок
    for i, col in enumerate(columns):
        tbl.columns[i].width = Cm(col.width_cm)

    # Индекс текущей строки в таблице
    row_idx = 0

    # --- Шапка-секция (опционально) ---
    if section_title:
        sec_row = tbl.rows[row_idx]
        sec_row.height = TABLE_SECTION_HEIGHT

        # Настоящий merge: первая ячейка сливается с последней.
        # В результате получается одна широкая ячейка на всю ширину таблицы.
        first_cell = sec_row.cells[0]
        last_cell = sec_row.cells[n_cols - 1]
        first_cell.merge(last_cell)

        _style_cell(
            first_cell,
            text=section_title,
            font_size=TABLE_SECTION_FONT_SIZE,
            font_color=TABLE_SECTION_TEXT_COLOR,
            bg_color=TABLE_SECTION_BG,
            bold=True,
            align="center",
        )
        row_idx += 1

    # --- Шапка с названиями колонок ---
    header_row = tbl.rows[row_idx]
    header_row.height = TABLE_HEADER_HEIGHT
    for j, col in enumerate(columns):
        _style_cell(
            header_row.cells[j],
            text=col.header,
            font_size=TABLE_HEADER_FONT_SIZE,
            font_color=TABLE_HEADER_TEXT_COLOR,
            bg_color=TABLE_HEADER_BG,
            bold=True,
            align="center",
        )
    row_idx += 1

    # --- Строки данных (zebra) ---
    for i, data_row in enumerate(rows):
        bg = TABLE_ROW_EVEN_BG if i % 2 == 0 else TABLE_ROW_ODD_BG
        table_row = tbl.rows[row_idx]
        table_row.height = TABLE_ROW_HEIGHT
        for j, val in enumerate(data_row):
            _style_cell(
                table_row.cells[j],
                text=val,
                font_size=TABLE_ROW_FONT_SIZE,
                font_color=TABLE_ROW_TEXT_COLOR,
                bg_color=bg,
                align=columns[j].align,
            )
        row_idx += 1

    return table_shape
