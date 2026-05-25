"""
Загрузка списка ISIN, которые должны быть выделены в таблицах
(колонка "В перечне").

Список хранится в общем файле /opt/config/highlighted_isins.csv —
один ISIN в строку, без заголовка. Формат CRLF или LF — без разницы.
Пустые строки и комментарии (#) игнорируются.

Файл редактируется вручную — добавляй/убирай ISIN и пересобирай отчёт.
"""
import os
from pathlib import Path
from typing import Set


# Дефолтный путь к общему файлу. На сервере он один и тот же для
# обоих проектов (bonds_rub и bonds_usd).
DEFAULT_PATH = Path(
    os.environ.get("BONDS_HIGHLIGHTED_ISINS", "/opt/config/highlighted_isins.csv")
)


def load_highlighted_isins(path: Path = None) -> Set[str]:
    """
    Возвращает множество ISIN, которые должны быть помечены «✓» в таблицах.

    Если файл не существует — возвращаем пустое множество (все бумаги
    без галочки). Это безопасный дефолт: отчёт соберётся даже без списка.

    Игнорируются:
      - пустые строки
      - комментарии (строки, начинающиеся с #)
      - повторы (множество уберёт дубли само)

    Args:
        path: путь к файлу. Если не задан — берётся из DEFAULT_PATH.

    Returns:
        Множество ISIN, в верхнем регистре, без пробелов и CRLF.
    """
    if path is None:
        path = DEFAULT_PATH

    path = Path(path)
    if not path.exists():
        return set()

    result: Set[str] = set()
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            # На всякий случай — приводим к верхнему регистру
            # (ISIN по стандарту ISO 6166 только в верхнем)
            result.add(line.upper())
    return result
