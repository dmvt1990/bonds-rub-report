"""
Расчёты по ОФЗ-ИН для слайда линкеров.

Два числа, которых нет ни в CSV пайплайна, ни у MOEX:

  * заблокированная индексация — насколько номинал ещё вырастет по уже
    ОПУБЛИКОВАННОМУ ИПЦ. Из-за лага индексации это известно на ~2 месяца
    вперёд, и держатель эту часть уже заработал;
  * вменённая инфляция (breakeven) — реальная доходность линкера против
    номинальной доходности ОФЗ-ПД того же срока.

Индексируемый номинал и реальную доходность НЕ пересчитываем: в отчёт идут
опубликованные MOEX значения (`facevalue`, `yield_close`) — именно их клиент
видит в своей выписке, и конкурирующая цифра в презентации была бы риском.
Свой расчёт номинала используется только как тихая сверка: расхождение
больше TOLERANCE_PCT пишется в лог.

Соглашение индексации (установлено эмпирически 2026-11-03):

    Номинал(t) = 1000 × I(t) / I(дата начала торгов)
    I(дата в месяце M, день D) = idx[M-4] + (D-1)/дней_в_M × (idx[M-3] - idx[M-4])

Лаг — ЧЕТЫРЕ месяца в этой индексации, не три, как описывает Минфин:
подогнано по шагам номинала из истории MOEX (средняя ошибка 0.022 п.п.
против 0.456 при лаге 3). НЕ «исправлять» на 3 — проверено на всех четырёх
живых выпусках, расхождение с номиналом MOEX не больше 0.01%.

ИПЦ Росстата НЕ пересматривается, поэтому вход индексации окончателен:
пересчитанный номинал задним числом не поедет.

Источник ИПЦ — публичная выгрузка одного ряда с rustata.ru: без токена и
без лимитов. Сеть может не ответить — тогда функции возвращают None, и
слайд линкеров просто не рисуется (суточный отчёт из-за этого не падает).
"""
from __future__ import annotations

import calendar
import csv
import datetime as dt
import io
import logging
import urllib.request

log = logging.getLogger(__name__)

CPI_URL = "https://rustata.ru/series/s1_mrzs6sdgzuzlzvrv/export.csv"
CPI_REF = "fedstat/indicator_31074/CPI.643.M.PCT.1.MOM"
LAG = 4
TIMEOUT = 20
TOLERANCE_PCT = 0.05
# UA обязан начинаться с «Mozilla», иначе rustata отдаёт 503 (как minfin.gov.ru)
_HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; bonds-rub-report/1.0)"}


def _shift(ym: str, k: int) -> str:
    y, m = int(ym[:4]), int(ym[5:])
    m += k
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return f"{y}-{m:02d}"


class CpiIndex:
    """Цепной базисный индекс ИПЦ. База произвольна — в индексации участвуют
    только отношения, а они от базы не зависят."""

    def __init__(self, idx: dict, latest_month: str):
        self.idx = idx
        self.latest_month = latest_month

    @property
    def locked_until(self) -> dt.date:
        """Последняя дата, на которую номинал определён опубликованным ИПЦ."""
        ym = _shift(self.latest_month, LAG - 1)
        y, m = int(ym[:4]), int(ym[5:])
        return dt.date(y, m, calendar.monthrange(y, m)[1])

    def on(self, d: dt.date):
        ym = f"{d.year}-{d.month:02d}"
        a, b = self.idx.get(_shift(ym, -LAG)), self.idx.get(_shift(ym, -LAG + 1))
        if a is None or b is None:
            return None
        n = calendar.monthrange(d.year, d.month)[1]
        return a + (d.day - 1) / n * (b - a)

    def nominal(self, issue: dt.date, on: dt.date):
        i_now, i_base = self.on(on), self.on(issue)
        if i_now is None or i_base is None:
            return None
        return 1000.0 * i_now / i_base

    def locked_in_pct(self, today: dt.date):
        """Прирост номинала от today до locked_until, в процентах.
        От даты выпуска не зависит — это отношение двух значений индекса."""
        a, b = self.on(today), self.on(self.locked_until)
        if a is None or b is None:
            return None
        return (b / a - 1) * 100.0


def load_cpi():
    """Качает ИПЦ м/м и строит цепной индекс. None — если не получилось."""
    try:
        req = urllib.request.Request(CPI_URL, headers=_HEADERS)
        raw = urllib.request.urlopen(req, timeout=TIMEOUT).read().decode("utf-8")
        rows = list(csv.DictReader(io.StringIO(raw)))
        if not rows:
            log.warning("ОФЗ-ИН: пустая выгрузка ИПЦ")
            return None
        got = (rows[0].get("series_ref") or "").strip()
        if got != CPI_REF:
            # Непрозрачный s1_-id мог быть переназначен — лучше отказаться,
            # чем молча считать индексацию по чужому ряду.
            log.error("ОФЗ-ИН: series_ref не совпал (%s, ждали %s) — слайд пропущен",
                      got, CPI_REF)
            return None
        mom = {r["period_source"]: float(r["value"])
               for r in rows if r.get("value") and len(r.get("period_source", "")) == 7}
        if not mom:
            return None
        idx, lvl = {}, 100.0
        for m in sorted(mom, reverse=True):
            idx[m] = lvl
            lvl /= mom[m] / 100.0
        latest = max(mom)
        # print, не log.info: в отчёте логирование не настроено, и уровень
        # INFO не доходит до cron-лога. Предупреждения (WARNING и выше)
        # попадают туда через last-resort handler — их оставляем на logging.
        print(f"ОФЗ-ИН: ИПЦ до {latest}, номинал определён до "
              f"{CpiIndex(idx, latest).locked_until}")
        return CpiIndex(idx, latest)
    except Exception as exc:                                  # noqa: BLE001
        log.warning("ОФЗ-ИН: ИПЦ недоступен (%s) — слайд линкеров пропущен", exc)
        return None


def check_nominal(cpi: CpiIndex, bond, issue: dt.date, today: dt.date) -> None:
    """Тихая сверка своего расчёта с опубликованным номиналом MOEX."""
    calc = cpi.nominal(issue, today)
    if calc is None or not bond.facevalue:
        return
    diff = (calc / float(bond.facevalue) - 1) * 100.0
    if abs(diff) > TOLERANCE_PCT:
        log.warning("ОФЗ-ИН %s: расчётный номинал %.2f против MOEX %.2f (%+.3f%%) — "
                    "проверьте свежесть ИПЦ и дату начала индексации",
                    bond.shortname or bond.isin, calc, float(bond.facevalue), diff)


def breakeven_pct(real_yield_pct, matdate, nominal_curve):
    """Вменённая инфляция: (1+ном)/(1+реальн) - 1, в процентах.

    Отношение, а не разность: на текущих уровнях доходностей разность
    занижает результат больше чем на процентный пункт.

    nominal_curve — [(дата погашения, доходность %)] по ОФЗ-ПД; между
    соседними выпусками интерполируем линейно по сроку.
    """
    if real_yield_pct is None or matdate is None or not nominal_curve:
        return None
    pts = sorted((d, y) for d, y in nominal_curve if d and y is not None)
    lo = [p for p in pts if p[0] <= matdate]
    hi = [p for p in pts if p[0] >= matdate]
    if not lo or not hi:
        return None
    (d1, y1), (d2, y2) = lo[-1], hi[0]
    nom = y1 if d1 == d2 else y1 + (matdate - d1).days / (d2 - d1).days * (y2 - y1)
    return ((1 + nom / 100.0) / (1 + real_yield_pct / 100.0) - 1) * 100.0
