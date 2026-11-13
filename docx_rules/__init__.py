"""
Правила проверки DOCX-документа на соответствие требованиям оформления.

Каждое правило живёт в отдельном модуле; общая база (Severity, Violation, BaseRule)
— в base.py, вспомогательные функции — в helpers.py, разметка абзацев — в
classification.py. До рефакторинга всё это был один файл docx_rules.py.
"""

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.classification import classify_paragraphs
from docx_rules.helpers import (effective, field_instructions, full_text, has_field,
                                has_page_field, is_heading, is_toc_line, is_word_heading,
                                paragraph_section_indices, style_chain)
from docx_rules.text_grid import TextGridRule
from docx_rules.page_margins import PageMarginsRule
from docx_rules.main_text import MainTextFormatRule
from docx_rules.page_numbering import PageNumberingRule
from docx_rules.title_page import TitlePageRule
from docx_rules.table_of_contents import TableOfContentsRule
from docx_rules.section_format import SectionFormatRule
from docx_rules.section_page_fill import SectionPageFillRule
from docx_rules.references_list import ReferencesListRule
from docx_rules.introduction_content import IntroductionContentRule
from docx_rules.paragraph_conclusion import ParagraphConclusionRule
from docx_rules.source_references import SourceReferencesRule
from docx_rules.appendix_format import AppendixFormatRule
from docx_rules.figures import FiguresRule
from docx_rules.tables_format import TablesFormatRule
from docx_rules.formulas import FormulasRule
from docx_rules.lists_format import ListsFormatRule
from docx_rules.dashes import DashesRule
from docx_rules.white_fill import WhiteFillRule

__all__ = [
    'BaseRule', 'Severity', 'Violation', 'classify_paragraphs',
    'effective', 'field_instructions', 'full_text', 'has_field', 'has_page_field',
    'is_heading', 'is_toc_line', 'is_word_heading', 'paragraph_section_indices',
    'style_chain',
    'TextGridRule', 'PageMarginsRule', 'MainTextFormatRule', 'PageNumberingRule',
    'TitlePageRule', 'TableOfContentsRule', 'SectionFormatRule', 'SectionPageFillRule',
    'ReferencesListRule', 'IntroductionContentRule', 'ParagraphConclusionRule',
    'SourceReferencesRule', 'AppendixFormatRule', 'FiguresRule', 'TablesFormatRule',
    'FormulasRule', 'ListsFormatRule', 'DashesRule', 'WhiteFillRule',
]
