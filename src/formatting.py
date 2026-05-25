"""
Утилиты форматирования для презентации.
"""
from datetime import date, datetime
from typing import Union

from .settings import RU_MONTHS_GENITIVE


def format_russian_date(d: Union[date, datetime, str]) -> str:
    """
    Форматирует дату в формате '16 апреля 2026 г.' как в образце.

    >>> format_russian_date(date(2026, 4, 16))
    '16 апреля 2026 г.'
    """
    if isinstance(d, str):
        d = datetime.strptime(d, "%Y-%m-%d").date()
    elif isinstance(d, datetime):
        d = d.date()

    month_name = RU_MONTHS_GENITIVE[d.month]
    return f"{d.day} {month_name} {d.year} г."


def format_number(value, decimals: int = 2) -> str:
    """
    Форматирует число в русском стиле: пробел как разделитель тысяч,
    запятая как десятичный разделитель.

    >>> format_number(364091)
    '364 091,00'
    >>> format_number(12.97)
    '12,97'
    >>> format_number(None)
    '—'
    """
    if value is None or value == "":
        return "—"
    try:
        num = float(value)
    except (TypeError, ValueError):
        return "—"

    # Разделитель тысяч - неразрывный пробел (U+00A0)
    formatted = f"{num:,.{decimals}f}".replace(",", "\u00A0").replace(".", ",")
    return formatted


def format_percent(value, decimals: int = 2) -> str:
    """
    >>> format_percent(12.97)
    '12,97%'
    """
    if value is None or value == "":
        return "—"
    return f"{format_number(value, decimals)}%"


def format_date_dmy(d) -> str:
    """
    Формат даты для таблиц: DD.MM.YYYY (как в образце Газпромбанка).

    >>> format_date_dmy(date(2026, 9, 16))
    '16.09.2026'
    """
    if d is None:
        return "—"
    if isinstance(d, str):
        try:
            d = datetime.strptime(d, "%Y-%m-%d").date()
        except ValueError:
            return d
    return d.strftime("%d.%m.%Y")


def format_integer(value) -> str:
    """
    Целое число с пробелом-разделителем тысяч. Для 'Объём выпуска, млн'.

    >>> format_integer(364091)
    '364 091'
    """
    if value is None or value == "":
        return "—"
    try:
        num = int(round(float(value)))
    except (TypeError, ValueError):
        return "—"
    return f"{num:,}".replace(",", "\u00A0")


def format_duration(value) -> str:
    """
    Дюрация в формате '0,4' — одна десятая, запятая как разделитель.

    >>> format_duration(0.4)
    '0,4'
    """
    if value is None or value == "":
        return "—"
    try:
        return f"{float(value):.1f}".replace(".", ",")
    except (TypeError, ValueError):
        return "—"
