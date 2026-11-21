# -*- coding: utf-8 -*-
"""Разрешение шрифтов темы.

Текст со стилями стандартного шаблона Word («Заголовок 2» и т.п.) задаёт шрифт
ссылкой на тему (asciiTheme) вместо явного имени: без разрешения темы валидатор
молча пропускал Calibri и показывал только неверный кегль.
"""
import unittest

from docx import Document

from docx_rules.helpers import effective
from docx_rules.tz34_text_format import Tz34TextFormatRule


class ThemeFontTests(unittest.TestCase):
    def test_theme_font_of_style_is_resolved(self):
        doc = Document()
        heading = doc.add_paragraph('Раздел', style='Heading 2')
        self.assertEqual('Calibri', effective(heading, 'name', heading.runs[0]))

    def test_theme_font_of_doc_defaults_is_resolved(self):
        doc = Document()
        paragraph = doc.add_paragraph('Обычный текст')
        self.assertEqual('Cambria', effective(paragraph, 'name', paragraph.runs[0]))

    def test_direct_formatting_wins_over_theme(self):
        from docx.shared import Pt
        doc = Document()
        heading = doc.add_paragraph('Раздел', style='Heading 2')
        heading.runs[0].font.name = 'Times New Roman'
        heading.runs[0].font.size = Pt(14)
        self.assertEqual('Times New Roman', effective(heading, 'name', heading.runs[0]))

    def test_text_format_rule_reports_theme_font(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')  # граница содержания
        body = doc.add_paragraph('Набрано шрифтом из темы шаблона.')
        found = Tz34TextFormatRule().check(doc)
        joined = ' '.join(v.message for v in found)
        self.assertIn("Шрифт 'Cambria'", joined)

    def test_styled_heading_is_reported_with_size(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')
        doc.add_paragraph('Раздел документа', style='Heading 2')
        found = Tz34TextFormatRule().check(doc)
        joined = ' '.join(v.message for v in found)
        self.assertIn("Шрифт 'Calibri'", joined)
        self.assertIn('а должен быть', joined)


if __name__ == '__main__':
    unittest.main()
