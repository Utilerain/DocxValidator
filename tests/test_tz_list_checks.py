# -*- coding: utf-8 -*-
"""Списки в профиле tz-34-602: автосписки Word с тире на красной строке 1,25 см.

Образец ПР2 оформляет перечни автосписками с маркером «–» (U+2013) на красной
строке 1,25 см; пункты разделяются точкой с запятой, последний — точкой.
Маркеры «‒» (U+2012) и «-» из окна «Символ» (\uf02d, Symbol) не допускаются.
"""
import unittest

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

from docx_rules import ListsFormatRule
from profiles import get_rules


def add_numbering(doc, num_id, abstract_id, lvl_text, fmt='bullet', left=None, hanging=None):
    """Добавляет abstractNum + num в numbering-часть документа."""
    numbering = doc.part.numbering_part.element
    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(abstract_id))
    lvl = OxmlElement('w:lvl')
    lvl.set(qn('w:ilvl'), '0')
    fmt_el = OxmlElement('w:numFmt')
    fmt_el.set(qn('w:val'), fmt)
    text_el = OxmlElement('w:lvlText')
    text_el.set(qn('w:val'), lvl_text)
    lvl.append(fmt_el)
    lvl.append(text_el)
    if left is not None or hanging is not None:
        ppr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        if left is not None:
            ind.set(qn('w:left'), str(left))
        if hanging is not None:
            ind.set(qn('w:hanging'), str(hanging))
        ppr.append(ind)
        lvl.append(ppr)
    abstract.append(lvl)
    numbering.append(abstract)
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    ref = OxmlElement('w:abstractNumId')
    ref.set(qn('w:val'), str(abstract_id))
    num.append(ref)
    numbering.append(num)


def set_num_pr(paragraph, num_id, ilvl=0):
    ppr = paragraph._p.get_or_add_pPr()
    numpr = OxmlElement('w:numPr')
    ilvl_el = OxmlElement('w:ilvl')
    ilvl_el.set(qn('w:val'), str(ilvl))
    numid_el = OxmlElement('w:numId')
    numid_el.set(qn('w:val'), str(num_id))
    numpr.append(ilvl_el)
    numpr.append(numid_el)
    ppr.append(numpr)


def bullet_document(marker='–', first_line=None, items=('первый пункт;', 'второй пункт.'),
                    lead='Система предназначена для:'):
    doc = Document()
    doc.add_paragraph('СОДЕРЖАНИЕ')
    doc.add_paragraph('ВВЕДЕНИЕ')  # граница содержания
    doc.add_paragraph(lead)
    add_numbering(doc, 100, 100, marker)
    for item in items:
        paragraph = doc.add_paragraph(item)
        set_num_pr(paragraph, 100)
        if first_line is not None:
            paragraph.paragraph_format.first_line_indent = Cm(first_line)
            paragraph.paragraph_format.left_indent = Cm(0)
    return doc


class TzListRuleTests(unittest.TestCase):
    def rule(self):
        return ListsFormatRule(auto_dash_indent_cm=1.25, manual_indent_cm=1.25,
                               check_manual_items=False)

    def messages(self, doc):
        return [v.message for v in self.rule().check(doc)]

    def test_profile_contains_list_rule_with_sample_norms(self):
        rules = get_rules('tz-34-602')
        lists = [r for r in rules if isinstance(r, ListsFormatRule)]
        self.assertEqual(1, len(lists))
        self.assertEqual(1.25, lists[0].auto_dash_indent_cm)
        self.assertFalse(lists[0].check_manual_items)

    def test_sample_style_list_passes(self):
        doc = bullet_document(first_line=1.25)
        self.assertEqual([], self.messages(doc))

    def test_wrong_bullet_character_is_reported(self):
        for marker in ('‒', '\uf02d', '•'):
            with self.subTest(marker=marker):
                messages = ' '.join(self.messages(bullet_document(marker=marker, first_line=1.25)))
                self.assertIn('Маркер списка должен быть тире', messages)

    def test_wrong_dash_position_is_reported(self):
        messages = ' '.join(self.messages(bullet_document(first_line=0.63)))
        self.assertIn('красная строка 1.25 см', messages)

    def test_semicolon_and_final_point_are_required(self):
        messages = ' '.join(self.messages(bullet_document(
            items=('первый пункт.', 'второй пункт,'), first_line=1.25)))
        self.assertIn('точкой с запятой', messages)
        self.assertIn('В последнем пункте списка ставится точка', messages)

    def test_colon_before_list_is_required(self):
        messages = ' '.join(self.messages(bullet_document(lead='Система решает задачи', first_line=1.25)))
        self.assertIn('двоеточие', messages)

    def test_numbered_heading_is_not_checked_as_list(self):
        # Заголовок раздела с автоматической нумерацией (numId) — не пункт списка
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')
        heading = doc.add_paragraph('Общие сведения', style='Heading 2')
        add_numbering(doc, 100, 100, '%1.', fmt='decimal')
        set_num_pr(heading, 100)
        self.assertEqual([], self.messages(doc))

    def test_manual_typed_items_are_not_checked_in_tz(self):
        # «1. Общие сведения» в ТЗ — заголовок раздела, а не нумерованный список
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')
        item = doc.add_paragraph('1. Общие сведения')
        item.paragraph_format.first_line_indent = Cm(1.25)
        self.assertEqual([], self.messages(doc))


if __name__ == '__main__':
    unittest.main()
