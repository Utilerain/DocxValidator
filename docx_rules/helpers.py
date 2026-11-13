"""Общие вспомогательные функции: шаблоны заголовков, разрешение форматирования, поля Word."""

import re
from typing import List

from docx import Document
from docx.oxml.ns import qn


def is_toc_line(text: str) -> bool:
    """
    Проверяет, является ли строка элементом оглавления
    (номер страницы в конце отделяется табуляцией или точками-заполнителями).
    """
    return bool(re.search(r'(\t|\.{2,}|…)\s*\d+\s*$', text))

def is_heading(text):
    return bool(re.match(r'^(?:ГЛАВА\s+\d+|\d+(?:\.\d+)+\.?\s+|'
                         r'(?:СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|БИБЛИОГРАФИЯ)$|'
                         r'СПИСОК\s+(?:ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ)|ПРИЛОЖЕНИЕ\b)', text, re.I))


def is_word_heading(text):
    """
    Заголовок, основанный на словах. Числовые шаблоны («2.1.», код специальности
    «09.02.07. …» на титульном листе) исключены: они не должны обрывать титульный лист.
    """
    return bool(re.match(r'^(?:ГЛАВА\s+\d+|'
                         r'(?:СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|БИБЛИОГРАФИЯ)$|'
                         r'СПИСОК\s+(?:ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ)|ПРИЛОЖЕНИЕ\b)', text, re.I))

def style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def effective(paragraph, attr, run=None):
    """Прямой формат, цепочки символьных/абзацных стилей и docDefaults."""
    sources = [run.font] if run is not None else [paragraph.paragraph_format]
    if run is not None:
        sources.extend(s.font for s in style_chain(run.style))
    sources.extend((s.font if run is not None else s.paragraph_format)
                   for s in style_chain(paragraph.style))
    for source in sources:
        value = getattr(source, attr)
        if value is not None:
            return value
    from docx.text.font import Font
    from docx.text.parfmt import ParagraphFormat
    defaults = paragraph.part.package.main_document_part.styles.element.find(qn('w:docDefaults'))
    if defaults is not None:
        node = defaults.find(qn('w:rPrDefault' if run is not None else 'w:pPrDefault'))
        if node is not None:
            from copy import deepcopy
            from docx.oxml import OxmlElement
            wrapper = OxmlElement('w:r' if run is not None else 'w:p')
            props = node.find(qn('w:rPr' if run is not None else 'w:pPr'))
            if props is not None:
                wrapper.append(deepcopy(props))
            return getattr(Font(wrapper) if run is not None else ParagraphFormat(wrapper), attr)
    return None

def field_instructions(container) -> List[str]:
    """Инструкции полей: w:fldSimple@w:instr и сложные поля (fldChar begin/end)."""
    instructions = [n.get(qn('w:instr'), '') for n in container._element.iter(qn('w:fldSimple'))]
    stack = []
    for node in container._element.iter():
        if node.tag == qn('w:fldChar'):
            kind = node.get(qn('w:fldCharType'))
            if kind == 'begin':
                stack.append('')
            elif kind == 'separate' and stack:
                instructions.append(stack[-1])
            elif kind == 'end' and stack:
                instructions.append(stack.pop())
        elif node.tag == qn('w:instrText') and stack:
            stack[-1] += node.text or ''
    return instructions


def has_field(container, name: str) -> bool:
    """True, если контейнер содержит поле с именем name (PAGE, TOC, ...)."""
    return any(re.match(rf'^\s*{name}\b', instruction, re.I) for instruction in field_instructions(container))


def has_page_field(container) -> bool:
    return has_field(container, 'PAGE')


def full_text(paragraph) -> str:
    """Текст абзаца: w:t и m:t, включая содержимое полей и гиперссылок; w:tab -> '\\t'."""
    parts = []
    for node in paragraph._p.iter():
        if node.tag in (qn('w:t'), qn('m:t')):
            parts.append(node.text or '')
        elif node.tag == qn('w:tab'):
            parts.append('\t')
        elif node.tag in (qn('w:br'), qn('w:cr')):
            parts.append('\n')
    return ''.join(parts)


def paragraph_section_indices(doc: Document) -> List[int]:
    """Индекс секции для каждого абзаца doc.paragraphs (секции разделяются sectPr внутри абзаца)."""
    mapping = []
    current = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            mapping.append(current)
            pPr = child.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                current += 1
    return mapping
