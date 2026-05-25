"""
Утилиты работы со шрифтами в pptx.

Основная проблема: python-pptx при установке `run.font.name = 'Arial'` пишет
в XML только `<a:latin typeface="Arial"/>`. Для кириллицы и ряда других
символов PowerPoint ищет парный шрифт в слотах `<a:ea>` (East Asian) и
`<a:cs>` (Complex Script). Если там пусто — подставляется системный дефолт.
На Windows без установленных русских шрифтов это может дать иероглифы.

Чтобы гарантировать корректный рендер, прописываем один шрифт во все три
слота через low-level XML.
"""
from lxml import etree
from pptx.oxml.ns import qn


def set_run_font(run, font_name: str) -> None:
    """
    Прописывает шрифт во все три слота run'а: latin, ea, cs.
    Используй вместо `run.font.name = font_name` везде, где в тексте
    может быть кириллица.
    """
    rPr = run._r.get_or_add_rPr()
    # Удаляем существующие тэги — если python-pptx их уже поставил
    for tag in ("a:latin", "a:ea", "a:cs"):
        for existing in rPr.findall(qn(tag)):
            rPr.remove(existing)
    # Добавляем заново — все три слота с одним и тем же шрифтом
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font_name)
