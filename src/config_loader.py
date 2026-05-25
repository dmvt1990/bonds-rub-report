"""Чтение config/bonds.yaml."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import yaml

from .settings import CONFIG_DIR


@dataclass
class BondRef:
    """Одна ссылка на бумагу в конфиге: ISIN + опциональные атрибуты."""
    isin: str
    coupon_formula: Optional[str] = None   # для флоатеров: "КС ЦБ РФ +1,5%"
    offer_date: Optional[str] = None        # DD.MM.YYYY, если задано — переопределяет CSV


@dataclass
class SectionConfig:
    id: str
    title: str
    bonds: List[BondRef] = field(default_factory=list)

    @property
    def isins(self) -> List[str]:
        """Обратная совместимость: раньше поле называлось .isins."""
        return [b.isin for b in self.bonds]

    def coupon_formulas(self) -> Dict[str, str]:
        """Быстрый доступ к формулам купонов по ISIN."""
        return {
            b.isin: b.coupon_formula
            for b in self.bonds
            if b.coupon_formula
        }

    def offer_dates(self) -> Dict[str, str]:
        """
        Даты оферт по ISIN (если заданы в yaml).
        Формат — строка DD.MM.YYYY, как в образце PDF.
        Когда задано — переопределяет значение из CSV (MOEX ISS не всегда
        отдаёт OFFERDATE для всех бумаг).
        """
        return {
            b.isin: b.offer_date
            for b in self.bonds
            if b.offer_date
        }


def load_sections(path: Path = None) -> List[SectionConfig]:
    """
    Возвращает список секций из bonds.yaml в порядке, заданном в файле.
    """
    if path is None:
        path = CONFIG_DIR / "bonds.yaml"

    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f)

    sections = []
    for sec_id, sec in (data.get("sections") or {}).items():
        bonds = []
        for entry in sec.get("bonds", []):
            if isinstance(entry, dict):
                bonds.append(BondRef(
                    isin=str(entry["isin"]).strip(),
                    coupon_formula=entry.get("coupon_formula"),
                    offer_date=entry.get("offer_date"),
                ))
            else:
                # запасной сценарий если кто-то писал просто список строк
                bonds.append(BondRef(isin=str(entry).strip()))

        sections.append(SectionConfig(
            id=sec_id,
            title=sec["title"],
            bonds=bonds,
        ))
    return sections
