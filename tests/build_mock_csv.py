"""
Мок-CSV с данными ровно из образца PDF (стр. 2, таблица суверенных ОФЗ).

Используется для отладки рендерера в среде без доступа к URL с реальным CSV.
После отладки будет выброшен — в продакшене рендерер читает URL.
"""
import csv

# Колонки должны совпадать с теми, что пишет patched bonds_update.py
COLUMNS = [
    "TRADEDATE", "ISIN", "SECID", "SHORTNAME", "NAME", "BOARDID", "TYPE",
    "MATDATE", "OFFERDATE", "FACEVALUE", "FACEUNIT", "ISSUESIZE", "LOTVALUE",
    "COUPONPERCENT", "COUPONVALUE", "COUPONFREQUENCY", "COUPONPERIOD",
    "OPEN", "LOW", "HIGH", "CLOSE", "WAPRICE", "VOLUME", "VALUE", "NUMTRADES",
    "ACCINT", "YIELDCLOSE", "YIELDATWAP", "DURATION_DAYS", "DURATION_YEARS",
    "STATUS",
]

# Данные ровно из образца: stр.2 таблица ОФЗ.
# issue_mln — это ОБЪЁМ в млн рублей (как в образце),
# здесь пересчитываем в штуки = млн * 1_000_000 / facevalue(1000).
OFZ_ROWS = [
    # (isin, shortname, matdate, coupon%, price%, ytm%, dur_years, issue_mln)
    ("RU000A0JWM07", "ОФЗ 26219", "2026-09-16", 7.75, 98.07, 12.97, 0.4, 364_091),
    ("RU000A0ZZYW2", "ОФЗ 26226", "2026-10-07", 7.95, 97.95, 12.91, 0.5, 368_568),
    ("RU000A0JS3W6", "ОФЗ 26207", "2027-02-03", 8.15, 96.83, 12.77, 0.8, 371_814),
    ("RU000A0JTK38", "ОФЗ 26212", "2028-01-19", 7.05, 91.67, 12.83, 1.7, 350_000),
    ("RU000A1038Z7", "ОФЗ 26237", "2029-03-14", 6.70, 85.62, 13.17, 2.7, 419_119),
    ("RU000A105RV3", "ОФЗ 26242", "2029-08-29", 9.00, 89.80, 13.22, 2.9, 529_357),
    ("RU000A100A82", "ОФЗ 26228", "2030-04-10", 7.65, 83.75, 13.43, 3.5, 514_950),
    ("RU000A0JVW48", "ОФЗ 26218", "2031-09-17", 8.50, 81.22, 13.94, 4.3, 413_387),
    ("RU000A105FZ9", "ОФЗ 26241", "2032-11-17", 9.50, 82.14, 14.16, 4.7, 750_000),
    ("RU000A1074G2", "ОФЗ 26244", "2034-03-15", 11.25, 86.85, 14.55, 5.2, 887_469),
    ("RU000A108EG6", "ОФЗ 26245", "2035-09-26", 12.00, 89.29, 14.58, 5.7, 922_311),
    ("RU000A108EF8", "ОФЗ 26247", "2039-05-11", 12.25, 89.23, 14.56, 6.1, 1_000_000),
    ("RU000A10D517", "ОФЗ 26253", "2038-10-06", 13.00, 93.26, 14.67, 5.9, 398_502),
    ("RU000A106E90", "ОФЗ 26243", "2038-05-19", 9.80, 75.51, 14.56, 6.3, 750_000),
    ("RU000A108EH4", "ОФЗ 26248", "2040-05-16", 12.25, 89.04, 14.54, 6.3, 1_000_000),
    ("RU000A103BR0", "ОФЗ 26240", "2036-07-30", 7.00, 62.74, 14.39, 6.5, 529_786),
    ("RU000A100EF5", "ОФЗ 26230", "2039-03-16", 7.70, 63.04, 14.43, 7.1, 444_326),
    ("RU000A1038V6", "ОФЗ 26238", "2041-05-15", 7.10, 59.61, 13.85, 7.4, 808_385),
]

# Флоатеры: (isin, shortname, matdate, offerdate, price%, issue_mln, freq)
# YTM и дюрация для флоатеров не заполняются (бессмысленны).
FLOATER_ROWS = [
    ("RU000A10DYH7", "ГПБ Финанс 1Р7Р",      "2027-06-17", "2026-05-05",  100.15, 20_000, 12),
    ("RU000A10AFX9", "Металлоинвест 1Р9",     "2026-06-17", None,           100.31, 10_000, 12),
    ("RU000A10C5F9", "Группа Черкизово Б2Р2", "2027-07-12", "2026-07-17",  100.34, 10_000, 12),
    ("RU000A10B0A2", "Магнит 5Р3",             "2026-08-27", None,           100.40, 84_800, 12),
    ("RU000A109K40", "ФосАгро БОП02",          "2026-09-08", None,           100.14, 35_000, 12),
    ("RU000A107WL0", "ИКС5 Финанс 3Р4",        "2034-02-28", "2026-09-08",  100.12, 10_000, 4),
    ("RU000A109S91", "Россети МосРегион 1Р7",  "2026-10-01", None,           100.35, 5_000,  12),
    ("RU000A109VM6", "Камаз БОП13",            "2026-10-15", None,           99.32,  5_000,  12),
    ("RU000A1075S4", "ИКС5 Финанс 3Р2",        "2026-10-17", None,           100.09, 20_000, 12),
    ("RU000A109R19", "ГПБ Финанс 1Р2Р",        "2027-10-24", "2026-10-25",  99.29,  10_000, 4),
    ("RU000A10A349", "РусГидро 2Р1",           "2026-11-04", None,           100.68, 40_000, 12),
    ("RU000A108G05", "Евраз Хол. 3Р01",        "2026-11-07", None,           99.97,  40_000, 12),
]


def build_row(isin, shortname, matdate, coupon, price, ytm, dur, issue_mln):
    facevalue = 1000
    issuesize = issue_mln * 1_000_000 / facevalue  # штук
    duration_days = round(dur * 365)

    return {
        "TRADEDATE": "2026-04-16",
        "ISIN": isin,
        "SECID": "",
        "SHORTNAME": shortname,
        "NAME": shortname,
        "BOARDID": "TQOB",
        "TYPE": "ofz_bond",
        "MATDATE": matdate,
        "OFFERDATE": "",
        "FACEVALUE": facevalue,
        "FACEUNIT": "SUR",
        "ISSUESIZE": int(issuesize),
        "LOTVALUE": 1000,
        "COUPONPERCENT": coupon,
        "COUPONVALUE": "",
        "COUPONFREQUENCY": 2,
        "COUPONPERIOD": "",
        "OPEN": "", "LOW": "", "HIGH": "",
        "CLOSE": price,
        "WAPRICE": price,
        "VOLUME": "", "VALUE": "", "NUMTRADES": "",
        "ACCINT": "",
        "YIELDCLOSE": ytm,
        "YIELDATWAP": ytm,
        "DURATION_DAYS": duration_days,
        "DURATION_YEARS": dur,
        "STATUS": "OK",
    }


def build_floater_row(isin, shortname, matdate, offerdate, price, issue_mln, freq):
    """
    Мок-строка для флоатера.
    Отличия от ОФЗ: COUPONPERCENT = 0 (MOEX не считает для флоатеров),
    YIELDCLOSE и DURATION_YEARS пустые, есть OFFERDATE.
    """
    facevalue = 1000
    issuesize = issue_mln * 1_000_000 / facevalue

    return {
        "TRADEDATE": "2026-04-16",
        "ISIN": isin,
        "SECID": "",
        "SHORTNAME": shortname,
        "NAME": shortname,
        "BOARDID": "TQCB",
        "TYPE": "exchange_bond",
        "MATDATE": matdate,
        "OFFERDATE": offerdate or "",
        "FACEVALUE": facevalue,
        "FACEUNIT": "SUR",
        "ISSUESIZE": int(issuesize),
        "LOTVALUE": 1000,
        "COUPONPERCENT": 0,     # признак флоатера
        "COUPONVALUE": "",
        "COUPONFREQUENCY": freq,
        "COUPONPERIOD": "",
        "OPEN": "", "LOW": "", "HIGH": "",
        "CLOSE": price,
        "WAPRICE": price,
        "VOLUME": "", "VALUE": "", "NUMTRADES": "",
        "ACCINT": "",
        "YIELDCLOSE": "",       # для флоатеров не считаем
        "YIELDATWAP": "",
        "DURATION_DAYS": "",
        "DURATION_YEARS": "",
        "STATUS": "OK",
    }


if __name__ == "__main__":
    import sys
    out_path = sys.argv[1] if len(sys.argv) > 1 else "/tmp/bonds_mock.csv"
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        for args in OFZ_ROWS:
            w.writerow(build_row(*args))
        for args in FLOATER_ROWS:
            w.writerow(build_floater_row(*args))

        # Дополняем мок-данными все ISIN из bonds.yaml, для которых у нас нет
        # «настоящих» тестовых значений. Это нужно, чтобы проверить пагинацию
        # и полный рендер без предупреждений о пропущенных бумагах.
        import random
        from pathlib import Path
        import yaml

        random.seed(42)  # детерминированный мок

        yaml_path = Path(__file__).resolve().parent.parent / "config" / "bonds.yaml"
        if yaml_path.exists():
            existing_isins = {
                r[0] for r in OFZ_ROWS
            } | {
                r[0] for r in FLOATER_ROWS
            }
            with open(yaml_path, encoding="utf-8") as yf:
                cfg = yaml.safe_load(yf)

            # Для каждого ISIN из yaml, которого ещё нет в моке, генерим строку.
            for sec_id, sec in (cfg.get("sections") or {}).items():
                is_floater_section = sec_id == "floaters"
                for entry in sec.get("bonds", []):
                    isin = entry["isin"] if isinstance(entry, dict) else entry
                    if isin in existing_isins:
                        continue
                    existing_isins.add(isin)
                    # Случайная, но правдоподобная матдата в 2026-2032
                    year = random.randint(2026, 2032)
                    month = random.randint(1, 12)
                    day = random.randint(1, 28)
                    matdate = f"{year}-{month:02d}-{day:02d}"
                    if is_floater_section:
                        w.writerow(build_floater_row(
                            isin,
                            f"Bond {isin[-5:]}",
                            matdate,
                            None,
                            round(random.uniform(95, 102), 2),
                            random.choice([5_000, 10_000, 20_000, 50_000]),
                            random.choice([4, 12]),
                        ))
                    else:
                        w.writerow(build_row(
                            isin,
                            f"Bond {isin[-5:]}",
                            matdate,
                            round(random.uniform(6, 12), 2),
                            round(random.uniform(70, 100), 2),
                            round(random.uniform(12, 15), 2),
                            round(random.uniform(0.5, 7), 1),
                            random.randint(100_000, 1_000_000),
                        ))

    total = len(OFZ_ROWS) + len(FLOATER_ROWS)
    print(f"Wrote to {out_path}")
