"""Разметка абзацев документа: титульный лист, оглавление, приложения, служебные абзацы."""

import re
from typing import List

from docx import Document
from docx.oxml.ns import qn

from docx_rules.helpers import is_heading, is_toc_line, is_word_heading, style_chain


def classify_paragraphs(doc: Document) -> List:
    """Эвристическая разметка: отсутствие оглавления не скрывает весь текст."""
    # Локальный импорт: formulas.py сам импортирует classification.py —
    # на уровне модулей получился бы цикл.
    from docx_rules.formulas import FormulasRule
    classified = list(doc.paragraphs)
    # Титульный лист заканчивается на первом «словесном» заголовке (Содержание,
    # Введение, Глава...). Числовой шаблон is_heading сюда не подходит — иначе
    # код специальности «09.02.07. …» на титуле обрывает титульный лист.
    first = next((i for i, p in enumerate(classified)
                  if is_word_heading(p.text.strip()) and not is_toc_line(p.text)), 0)
    in_toc = in_appendix = False
    for i, p in enumerate(classified):
        text = p.text.strip()
        p.is_toc_heading = text.upper() in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ')
        if p.is_toc_heading:
            in_toc = True
        elif in_toc and is_heading(text) and not is_toc_line(text):
            in_toc = False
        p.is_title_page = i < first
        p.is_toc = in_toc or is_toc_line(text)
        if not p.is_toc and re.match(r'^ПРИЛОЖЕНИЕ\b', text, re.I):
            in_appendix = True
        p.is_appendix = in_appendix
        p.is_special = (is_heading(text) or in_appendix
                        or bool(re.match(r'^(?:Рис(?:унок)?\.?\s*\d+|Таблица\s+\d+|Продолжение таблицы|[–—-]\s+|\d+[.)]\s+)', text, re.I))
                        or bool(p._p.xpath('./w:pPr/w:numPr'))
                        or any(st.name.startswith('List') for st in style_chain(p.style))
                        or FormulasRule()._is_display_formula(p))
    for i, p in enumerate(classified):
        if p._p.getnext() is not None and p._p.getnext().tag == qn('w:tbl'):
            if i and re.match(r'^Таблица\s+\d+', classified[i-1].text, re.I):
                p.is_special = True
    return classified

