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

    «Литера» и «Аннотация» — реквизиты технического задания по ГОСТ 19.104/19.201:
    они тоже завершают титульный лист (в таких титулах нет «Содержания»).
    Слово «Приложение …» сюда не входит: на титулах ТЗ это часть названия
    документа («Приложение "Организация свадебных торжеств"»), и оно не должно
    обрывать титульный лист.
    """
    return bool(re.match(r'^(?:ГЛАВА\s+\d+|'
                         r'(?:СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|БИБЛИОГРАФИЯ|'
                         r'ЛИТЕРА|АННОТАЦИЯ)$|'
                         r'СПИСОК\s+(?:ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ))', text, re.I))

def style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


# Шрифты из стандартного шаблона Word (например, «Заголовок 2») задаются
# ссылкой на тему (asciiTheme/hAnsiTheme), а не явным именем: python-docx
# возвращает для них None. Эти соответствия разрешают ссылку в имя шрифта.
_THEME_KINDS = {
    'majorHAnsi': 'major', 'majorAscii': 'major', 'majorBidi': 'major',
    'minorHAnsi': 'minor', 'minorAscii': 'minor', 'minorBidi': 'minor',
}


def _font_theme_kind(font):
    """'major'/'minor', если источник задаёт шрифт ссылкой на тему, иначе None."""
    element = getattr(font, '_element', None)
    if element is None:
        return None
    props = element.find(qn('w:rPr'))
    if props is None:
        return None
    rfonts = props.find(qn('w:rFonts'))
    if rfonts is None:
        return None
    for attribute in ('w:asciiTheme', 'w:hAnsiTheme', 'w:cstheme'):
        kind = _THEME_KINDS.get(rfonts.get(qn(attribute)))
        if kind:
            return kind
    return None


def theme_fonts(paragraph):
    """{'major': имя, 'minor': имя} из темы документа (word/theme/theme1.xml)."""
    package = paragraph.part.package
    cached = getattr(package, '_docxvalidator_theme_fonts', None)
    if cached is not None:
        return cached
    fonts = {'major': None, 'minor': None}
    try:
        for rel in package.main_document_part.rels.values():
            if rel.reltype.endswith('/theme'):
                xml = rel.target_part.blob.decode('utf-8', 'ignore')
                for kind, tag in (('major', 'a:majorFont'), ('minor', 'a:minorFont')):
                    match = re.search(rf'<{tag}>.*?<a:latin typeface="([^"]*)"', xml, re.S)
                    if match:
                        fonts[kind] = match.group(1)
                break
    except Exception:
        pass
    package._docxvalidator_theme_fonts = fonts
    return fonts


def effective(paragraph, attr, run=None):
    """Прямой формат, цепочки символьных/абзацных стилей, шрифты темы и docDefaults."""
    sources = [run.font] if run is not None else [paragraph.paragraph_format]
    if run is not None:
        sources.extend(s.font for s in style_chain(run.style))
    sources.extend((s.font if run is not None else s.paragraph_format)
                   for s in style_chain(paragraph.style))
    resolve_theme = attr == 'name' and run is not None
    for source in sources:
        value = getattr(source, attr)
        if value is not None:
            return value
        if resolve_theme:
            kind = _font_theme_kind(source)
            if kind:
                name = theme_fonts(paragraph).get(kind)
                if name:
                    return name
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
            value = getattr(Font(wrapper) if run is not None else ParagraphFormat(wrapper), attr)
            if value is not None:
                return value
            if resolve_theme:
                kind = _font_theme_kind(Font(wrapper))
                if kind:
                    return theme_fonts(paragraph).get(kind)
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
