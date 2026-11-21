"""Разметка абзацев документа: титульный лист, оглавление, приложения, служебные абзацы."""

import re
from typing import List

from docx import Document
from docx.oxml.ns import qn

from docx_rules.helpers import is_heading, is_toc_line, is_word_heading, style_chain


# Заголовок приложения: «ПРИЛОЖЕНИЕ», «Приложение А», «Приложение А. Название»,
# «Приложение 1», «Приложение — Название». Фразы внутри текста («Приложение
# предназначено для …») раздел приложений не начинают.
APPENDIX_START_REGEX = re.compile(
    r'^ПРИЛОЖЕНИЕ(?:\s*$|\s+(?:[А-ЯЁA-Z]\b\.?|\d+\b)|\s*[-–—:])', re.I)


def classify_paragraphs(doc: Document) -> List:
    """Эвристическая разметка: отсутствие оглавления не скрывает весь текст."""
    classified = list(doc.paragraphs)
    # Титульный лист заканчивается на первом «словесном» заголовке (Содержание,
    # Введение, Глава...). Числовой шаблон is_heading сюда не подходит — иначе
    # код специальности «09.02.07. …» на титуле обрывает титульный лист.
    word_boundary = next((i for i, p in enumerate(classified)
                          if is_word_heading(p.text.strip()) and not is_toc_line(p.text)), None)
    first = word_boundary if word_boundary is not None else 0
    # Номер раздела сам по себе ненадёжен: на титуле бывает код специальности.
    # Допускаем его как границу только после характерных реквизитов ТЗ.
    approved = sheets = year = False
    numbered_boundary = None
    for i, p in enumerate(classified):
        text = p.text.strip()
        if ((approved and (sheets or year) or sheets and year)
                and (_is_first_level_heading(text) or
                     re.fullmatch(r'ОБЩИЕ\s+СВЕДЕНИЯ[.:]?', text, re.I))
                and not is_toc_line(text)):
            if word_boundary is None or i < first:
                first = numbered_boundary = i
            break
        approved |= bool(re.match(r'^УТВЕРЖДЕНО\b', text, re.I))
        sheets |= bool(re.match(r'^ЛИСТОВ\s+\d+\b', text, re.I))
        year |= bool(re.fullmatch(r'\d{4}', text))
    # Настоящее приложение до следующего словесного заголовка не превращает
    # весь предшествующий текст в титул. Название «Приложение „...“» не метка.
    for i, p in enumerate(classified[:first]):
        if re.fullmatch(r'ПРИЛОЖЕНИЕ(?:\s+[А-ЯA-ZЁ0-9]+)?', p.text.strip(), re.I):
            prefix = [q.text.strip() for q in classified[:i]]
            has_title_evidence = (numbered_boundary is not None or
                                  any(re.match(r'^(?:УТВЕРЖДЕНО|ТЕХНИЧЕСКОЕ\s+ЗАДАНИЕ)\b', t, re.I)
                                      for t in prefix) or
                                  (any(re.match(r'^ЛИСТОВ\s+\d+\b', t, re.I) for t in prefix)
                                   and any(re.fullmatch(r'\d{4}', t) for t in prefix)))
            first = i if has_title_evidence else 0
            break
    in_toc = in_appendix = False
    toc_entries = set()
    for i, p in enumerate(classified):
        text = p.text.strip()
        p.is_toc_heading = text.upper() in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ')
        if p.is_toc_heading:
            in_toc = True
            toc_entries.clear()
        elif in_toc:
            toc_style = any(re.match(r'^(?:TOC|Оглавление)\s*\d', st.name or '', re.I)
                            for st in style_chain(p.style))
            entry = _toc_entry_key(text)
            # Номера уровней существенны: 1 и 1.1 с одинаковым названием —
            # разные пункты. Ненумерованный заголовок может повторять запись.
            repeated = bool(entry[1] and (entry in toc_entries or
                            (entry[0] is None and any(e[1] == entry[1] for e in toc_entries))))
            boundary = _has_heading_style(p) or repeated or is_word_heading(text)
            if text and not is_toc_line(text) and not toc_style and boundary:
                in_toc = False
            elif text:
                toc_entries.add(entry)
        p.is_title_page = i < first
        p.is_toc = in_toc or is_toc_line(text)
        # Раздел приложений начинается только после титульного листа: на титулах
        # технических заданий (ГОСТ 19.104) слово «Приложение …» — часть названия
        # документа, а не заголовок раздела приложений.
        title_name = (first == 0 and re.match(r'^ПРИЛОЖЕНИЕ\s*[«"“]', text, re.I)
                      and any(re.match(r'^УТВЕРЖДЕНО\b', q.text.strip(), re.I)
                              for q in classified[:i]))
        if (not p.is_toc and i >= first and not title_name
                and APPENDIX_START_REGEX.match(text)):
            in_appendix = True
        p.is_appendix = in_appendix
        p.is_special = (is_heading(text) or in_appendix
                        or bool(re.match(r'^(?:Рис(?:унок)?\.?\s*\d+|Таблица\s+\d+|Продолжение таблицы|[–—-]\s+|\d+[.)]\s+)', text, re.I))
                        or bool(p._p.xpath('./w:pPr/w:numPr'))
                        or any(st.name.startswith('List') for st in style_chain(p.style))
                        or _is_display_formula(p))
    for i, p in enumerate(classified):
        if p._p.getnext() is not None and p._p.getnext().tag == qn('w:tbl'):
            if i and re.match(r'^Таблица\s+\d+', classified[i-1].text, re.I):
                p.is_special = True
    return classified



def _toc_entry_key(text):
    text = re.sub(r'(?:\t|\.{2,}|…)\s*\d+\s*$', '', text).strip()
    match = re.match(r'^(\d+(?:\.\d+)*)(?:[.)])?\s+', text)
    number = match.group(1) if match else None
    if match:
        text = text[match.end():]
    return number, re.sub(r'\s+', ' ', text).strip(' .:').casefold().replace('ё', 'е')


def _has_heading_style(paragraph):
    return (any(re.match(r'^(?:Heading|Заголовок)\s*\d', st.name or '', re.I)
                for st in style_chain(paragraph.style))
            or bool(paragraph._p.xpath('./w:pPr/w:outlineLvl')))


def _is_first_level_heading(text):
    # Не принимаем десятичные номера и коды вида 09.02.07.
    return bool(re.match(r'^[1-9]\d*\.?\s+\S', text))


def _is_display_formula(paragraph):
    """Ленивый импорт исключает цикл classification -> formulas -> classification."""
    from docx_rules.formulas import FormulasRule
    return FormulasRule()._is_display_formula(paragraph)
