"""
Вычисление предыдущего торгового дня для отчёта.

Логика простая:
- В понедельник-пятницу предыдущий торговый день = вчера
- В субботу предыдущий торговый день = пятница (минус 1 день)
- В воскресенье предыдущий торговый день = пятница (минус 2 дня)

Праздники РФ пока не учитываются — при необходимости список в виде
констант можно добавить в settings.py, и расширить prev_trading_day().

Пример использования:
    from datetime import date
    from src.trading_day import previous_trading_day
    report_date = previous_trading_day()
"""
from datetime import date, timedelta
from typing import Optional


# Понедельник=0, Вторник=1, ..., Пятница=4, Суббота=5, Воскресенье=6
_SATURDAY = 5
_SUNDAY = 6


def previous_trading_day(today: Optional[date] = None) -> date:
    """
    Возвращает дату предыдущего торгового дня относительно `today`.

    Args:
        today: от какой даты отсчитывать. По умолчанию — date.today().
               Параметр полезен для тестирования и для cron-запусков,
               когда запуск идёт "за предыдущий день".

    Returns:
        date — предыдущий рабочий день (пн-пт).

    Примеры:
        >>> previous_trading_day(date(2026, 4, 22))  # среда → вторник
        datetime.date(2026, 4, 21)
        >>> previous_trading_day(date(2026, 4, 20))  # понедельник → пятница
        datetime.date(2026, 4, 17)
        >>> previous_trading_day(date(2026, 4, 18))  # суббота → пятница
        datetime.date(2026, 4, 17)
        >>> previous_trading_day(date(2026, 4, 19))  # воскресенье → пятница
        datetime.date(2026, 4, 17)
    """
    if today is None:
        today = date.today()

    # Шаг назад от today, пока не попадём в будний день.
    candidate = today - timedelta(days=1)
    while candidate.weekday() in (_SATURDAY, _SUNDAY):
        candidate -= timedelta(days=1)
    return candidate
