"""
Чтение bonds_data.csv с публичного URL и преобразование в объекты Bond.

Особенности:
- CSV публикуется moex_loader'ом на http://153.80.196.164/moex/bonds_data.csv
- pandas умеет читать URL напрямую, но мы используем requests с таймаутом,
  чтобы лучше контролировать ошибки сети.
- Все числа приходят как str или NaN (pandas). Преобразуем в Optional[float] /
  Optional[int] / Optional[date].
"""
from datetime import date, datetime
from io import StringIO
from typing import List, Optional

import pandas as pd
import requests

from .models import Bond
from .sources import BONDS_DATA_URL, FETCH_TIMEOUT


# Имена колонок в CSV (как их пишет bonds_update.py).
# Завязаны на контракт с лоадером — если лоадер переименует колонку,
# поменять нужно тут.
COL_ISIN = "ISIN"
COL_SECID = "SECID"
COL_SHORTNAME = "SHORTNAME"
COL_NAME = "NAME"
COL_TYPE = "TYPE"
COL_MATDATE = "MATDATE"
COL_OFFERDATE = "OFFERDATE"
COL_FACEVALUE = "FACEVALUE"
COL_FACEUNIT = "FACEUNIT"
COL_ISSUESIZE = "ISSUESIZE"
COL_LOTVALUE = "LOTVALUE"
COL_COUPONPERCENT = "COUPONPERCENT"
COL_COUPONFREQUENCY = "COUPONFREQUENCY"
COL_TRADEDATE = "TRADEDATE"
COL_CLOSE = "CLOSE"
COL_YIELDCLOSE = "YIELDCLOSE"
COL_DURATION_YEARS = "DURATION_YEARS"


def _to_float(v) -> Optional[float]:
    if v is None or pd.isna(v) or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _to_int(v) -> Optional[int]:
    f = _to_float(v)
    return int(f) if f is not None else None


def _to_date(v) -> Optional[date]:
    if v is None or pd.isna(v) or v == "":
        return None
    try:
        return datetime.strptime(str(v), "%Y-%m-%d").date()
    except ValueError:
        return None


def _to_str(v) -> Optional[str]:
    if v is None or pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def fetch_csv(source: str = BONDS_DATA_URL, timeout: int = FETCH_TIMEOUT) -> pd.DataFrame:
    """
    Возвращает DataFrame с содержимым CSV.

    Args:
        source: URL (http://...) или путь к локальному файлу.
                Если начинается на http:// или https:// — скачивается
                по сети с указанным таймаутом. Иначе читается с диска.

    Кодировка определяется автоматически: пробуем UTF-8, потом cp1251,
    потом latin-1 — это покрывает практически все варианты, в которых
    приходит CSV с русским текстом. Это предотвращает появление
    "крокозябр" вида ÐÐ¤Ð вместо "ОФЗ".
    """
    if source.startswith(("http://", "https://")):
        resp = requests.get(source, timeout=timeout)
        resp.raise_for_status()
        raw_bytes = resp.content
    else:
        # локальный файл (удобно для отладки и ручного запуска)
        with open(source, "rb") as f:
            raw_bytes = f.read()

    # Пробуем кодировки по очереди. UTF-8 — первый кандидат, потому что это
    # стандарт де-факто. cp1251 — вторая по частоте у русских CSV (особенно
    # сохранённых из Excel или скачанных с MOEX ISS в "legacy"-виде).
    # latin-1 всегда успешно декодирует любые байты (он однобайтовый и
    # покрывает весь диапазон 0-255), поэтому ставим последним как страховку.
    for encoding in ("utf-8", "cp1251", "latin-1"):
        try:
            text = raw_bytes.decode(encoding)
            # Эвристика: если после декодирования видно типичную "крокозябру"
            # (Ð в начале слов), значит кодировка угадана неверно —
            # пробуем следующую.
            if encoding == "latin-1" and "Ð" in text[:2000]:
                continue
            return pd.read_csv(StringIO(text))
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue

    # Последняя надежда: отдать pandas как есть, он может попробовать сам
    return pd.read_csv(StringIO(raw_bytes.decode("latin-1", errors="replace")))


def row_to_bond(row: pd.Series) -> Bond:
    """Превращает одну строку DataFrame в объект Bond."""
    return Bond(
        isin=str(row[COL_ISIN]).strip(),
        secid=_to_str(row.get(COL_SECID)) or "",
        shortname=_to_str(row.get(COL_SHORTNAME)) or "",
        name=_to_str(row.get(COL_NAME)),
        type_=_to_str(row.get(COL_TYPE)),
        matdate=_to_date(row.get(COL_MATDATE)),
        offerdate=_to_date(row.get(COL_OFFERDATE)),
        facevalue=_to_float(row.get(COL_FACEVALUE)),
        faceunit=_to_str(row.get(COL_FACEUNIT)),
        issuesize=_to_float(row.get(COL_ISSUESIZE)),
        lotvalue=_to_float(row.get(COL_LOTVALUE)),
        couponpercent=_to_float(row.get(COL_COUPONPERCENT)),
        couponfrequency=_to_int(row.get(COL_COUPONFREQUENCY)),
        tradedate=_to_date(row.get(COL_TRADEDATE)),
        close=_to_float(row.get(COL_CLOSE)),
        yield_close=_to_float(row.get(COL_YIELDCLOSE)),
        duration_years=_to_float(row.get(COL_DURATION_YEARS)),
    )


def load_bonds(isins: List[str], source: str = BONDS_DATA_URL) -> List[Bond]:
    """
    Скачивает полный CSV и возвращает только те облигации, чьи ISIN
    есть в переданном списке. Порядок результата соответствует порядку isins.

    Если какого-то ISIN нет в CSV — он просто пропускается (с предупреждением
    в консоль). Отчёт должен собираться, даже если пара бумаг выпали.

    Args:
        source: URL или путь к локальному CSV (см. fetch_csv).
    """
    df = fetch_csv(source)

    # Индекс по ISIN для быстрого доступа
    by_isin = {str(r[COL_ISIN]).strip(): r for _, r in df.iterrows()}

    bonds = []
    for isin in isins:
        row = by_isin.get(isin)
        if row is None:
            print(f"[warn] ISIN {isin} не найден в CSV — пропускаю")
            continue
        bonds.append(row_to_bond(row))

    return bonds
