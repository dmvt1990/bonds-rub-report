"""
Корпоративные константы стиля Газпромбанк Private.

Значения выверены по образцу PDF и шаблону template.pptx.
Если дизайн-гайд Газпромбанка обновится - правки делаются только здесь.
"""
from pptx.util import Pt, Cm, Emu
from pptx.dml.color import RGBColor

# --- Шрифты ---
# В оригинальном шаблоне Газпромбанка используется Cera CY, но этот шрифт
# есть не на всех системах. При открытии на Windows без Cera CY PowerPoint
# подставляет fallback для кириллицы из азиатских шрифтов — получаются
# иероглифы. Поэтому используем Arial: он есть везде, поддерживает
# кириллицу, визуально достаточно нейтрален для корпоративного отчёта.
FONT_PRIMARY = "Arial"
FONT_FALLBACK = "Arial"

# --- Цвета ---
# Выверено по образцу PDF и shape в шаблоне (525252 - подзаголовок)
COLOR_TEXT_PRIMARY = RGBColor(0x26, 0x26, 0x26)      # почти чёрный, для заголовков
COLOR_TEXT_SECONDARY = RGBColor(0x52, 0x52, 0x52)    # серый, для дат и подписей
COLOR_TEXT_MUTED = RGBColor(0x8C, 0x8C, 0x8C)        # светло-серый, для дисклеймеров
COLOR_ACCENT_OLIVE = RGBColor(0xB9, 0xA7, 0x7A)      # оливково-золотой корпоративный
COLOR_LINE_SEPARATOR = RGBColor(0x8C, 0x8C, 0x8C)    # для разделительных линий

# --- Размеры шрифта ---
# По образцу обложки: большой заголовок, мелкая дата, совсем мелкий дисклеймер
FONT_SIZE_COVER_TITLE = Pt(36)
FONT_SIZE_COVER_DATE = Pt(18)
FONT_SIZE_COVER_DISCLAIMER = Pt(9)

# --- Геометрия обложки ---
# Координаты placeholder'а заголовка из layout 0 шаблона:
# pos: (1.13, 7.38) cm,  size: 14.51 x 4.30 cm
# Дату и дисклеймер размещаем под заголовком, в той же колонке.
COVER_TITLE_LEFT = Cm(1.13)
COVER_TITLE_TOP = Cm(7.38)
COVER_TITLE_WIDTH = Cm(14.51)
COVER_TITLE_HEIGHT = Cm(4.30)

# Дата с разделительной линией и дисклеймер — в нижней части слайда
# (слайд 19.05 см высотой; в образце PDF это нижняя треть).
COVER_DATE_LEFT = Cm(1.13)
COVER_DATE_TOP = Cm(15.70)
COVER_DATE_WIDTH = Cm(10.00)
COVER_DATE_HEIGHT = Cm(0.80)

# Тонкая горизонтальная линия над дисклеймером (как в образце)
COVER_SEPARATOR_LEFT = Cm(1.13)
COVER_SEPARATOR_TOP = Cm(16.70)
COVER_SEPARATOR_WIDTH = Cm(6.00)
COVER_SEPARATOR_HEIGHT = Cm(0.03)

COVER_DISCLAIMER_LEFT = Cm(1.13)
COVER_DISCLAIMER_TOP = Cm(16.85)
COVER_DISCLAIMER_WIDTH = Cm(10.00)
COVER_DISCLAIMER_HEIGHT = Cm(0.50)

# --- Стиль таблицы облигаций (образец стр. 2–11) ---

# Шапка таблицы: оливково-бежевый фон, белый жирный текст
TABLE_HEADER_BG = RGBColor(0xB9, 0xA7, 0x7A)       # корпоративный оливковый
TABLE_HEADER_TEXT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_HEADER_FONT_SIZE = Pt(7)
TABLE_HEADER_HEIGHT = Cm(1.10)

# Шапка секции (строка с названием "Суверенные облигации с фикс. купоном")
# — светло-бежевый фон, серый жирный текст
TABLE_SECTION_BG = RGBColor(0xE8, 0xE0, 0xCC)
TABLE_SECTION_TEXT_COLOR = RGBColor(0x52, 0x52, 0x52)
TABLE_SECTION_FONT_SIZE = Pt(9)
TABLE_SECTION_HEIGHT = Cm(0.55)

# Строки данных: zebra (чётные/нечётные). Серый текст, без границ.
TABLE_ROW_EVEN_BG = RGBColor(0xFF, 0xFF, 0xFF)      # белый
TABLE_ROW_ODD_BG = RGBColor(0xF7, 0xF5, 0xEF)       # почти-белый с тёплым оттенком
TABLE_ROW_TEXT_COLOR = RGBColor(0x26, 0x26, 0x26)
TABLE_ROW_FONT_SIZE = Pt(7)
TABLE_ROW_HEIGHT = Cm(0.42)

# --- Геометрия слайда «таблица облигаций» ---
# По мотивам стр. 2 образца: заголовок сверху, таблица ниже,
# снизу источник данных + дисклеймер + номер страницы.
TABLE_SLIDE_TITLE_LEFT = Cm(1.13)
TABLE_SLIDE_TITLE_TOP = Cm(2.20)
TABLE_SLIDE_TITLE_WIDTH = Cm(20.00)
TABLE_SLIDE_TITLE_HEIGHT = Cm(1.20)
TABLE_SLIDE_TITLE_FONT_SIZE = Pt(24)

TABLE_SLIDE_LEFT = Cm(1.13)
TABLE_SLIDE_TOP = Cm(5.40)         # под заголовком, с запасом
TABLE_SLIDE_TOTAL_WIDTH = Cm(31.60)  # по ширине слайда минус отступы

# Источник и дисклеймер снизу
TABLE_SLIDE_SOURCE_TOP = Cm(17.20)
TABLE_SLIDE_SOURCE_FONT_SIZE = Pt(8)

TABLE_SLIDE_FOOTNOTE_TOP = Cm(17.90)
TABLE_SLIDE_FOOTNOTE_FONT_SIZE = Pt(6)

# Ширины колонок таблицы ОФЗ (11 колонок: 10 данных + "В перечне")
TABLE_OFZ_COL_WIDTHS_CM = [
    1.19,  # #
    2.99,  # Выпуск
    4.18,  # ISIN
    3.29,  # Дата погашения
    2.69,  # Купон, % год.
    2.69,  # Цена, % от ном.
    3.29,  # Дох-ть к погаш.
    2.09,  # Мод. дюр.
    4.18,  # Объём выпуска, млн.
    3.29,  # Период. купона
    1.70,  # В перечне
]

# Ширины колонок таблицы флоатеров (10 колонок: 9 данных + "В перечне")
TABLE_FLOATERS_COL_WIDTHS_CM = [
    1.03,  # #
    3.78,  # Выпуск
    4.81,  # ISIN
    3.09,  # Дата погашения
    3.09,  # Дата оферты
    4.81,  # Купон (формула)
    2.40,  # Цена, % от ном.
    4.46,  # Объём выпуска, млн ₽
    2.40,  # Период. купона
    1.70,  # В перечне
]

# Ширины колонок таблицы корпоративных фиксов (12 колонок: 11 данных + "В перечне")
TABLE_CORP_FIXED_COL_WIDTHS_CM = [
    0.92,  # #
    3.36,  # Выпуск
    4.27,  # ISIN
    2.74,  # Дата погашения
    2.74,  # Дата оферты
    2.44,  # Купон, % год.
    2.44,  # Цена, % от ном.
    3.05,  # Дох-ть к погаш.
    1.83,  # Мод. дюр.
    3.66,  # Объём выпуска, млн ₽
    2.44,  # Период. купона
    1.70,  # В перечне
]
