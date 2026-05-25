"""
Модели данных — плоские dataclasses, которые передаются между слоями.

Все значения — Optional, потому что MOEX может не отдать часть полей
(например, YTM и дюрация для флоатеров). Рендерер в таком случае
рисует прочерк.
"""
from dataclasses import dataclass
from datetime import date
from typing import Optional


@dataclass
class Bond:
    """Одна облигация на дату отчёта."""

    # --- идентификация ---
    isin: str
    secid: str
    shortname: str              # «ОФЗ 26219», «Магнит 5Р3»
    name: Optional[str]         # полное название из MOEX
    type_: Optional[str]        # ofz_bond / corporate_bond / exchange_bond / ...

    # --- параметры выпуска ---
    matdate: Optional[date]
    offerdate: Optional[date]
    facevalue: Optional[float]
    faceunit: Optional[str]     # SUR / USD / EUR
    issuesize: Optional[float]  # в штуках!
    lotvalue: Optional[float]   # минимальный лот в рублях
    couponpercent: Optional[float]
    couponfrequency: Optional[int]  # раз в год: 2/4/12

    # --- рыночные данные на дату ---
    tradedate: Optional[date]
    close: Optional[float]              # цена закрытия, % от номинала
    yield_close: Optional[float]        # YTM на закрытие, % годовых
    duration_years: Optional[float]     # мод. дюрация в годах

    # --- вычисляемые поля (заполняются в transform/classify) ---
    section: Optional[str] = None   # 'ofz_fixed' / 'floaters' / 'corp_fixed'

    @property
    def issuesize_mln_rub(self) -> Optional[float]:
        """Объём выпуска в млн рублей = штук * номинал / 1_000_000."""
        if self.issuesize is None or self.facevalue is None:
            return None
        return self.issuesize * self.facevalue / 1_000_000

    @property
    def is_floater(self) -> bool:
        """
        Флоатер: купон = 0 или None.
        У флоатеров в MOEX COUPONPERCENT обычно 0.0, потому что
        текущая ставка купона не зафиксирована.
        """
        return not self.couponpercent  # None или 0.0
