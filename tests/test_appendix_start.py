# -*- coding: utf-8 -*-
"""Начало раздела приложений: заголовок, а не фраза в тексте.

«Приложение предназначено для использования…» — обычное предложение в теле ТЗ;
раньше оно включало режим приложения до конца документа, из-за чего списки
переставали проверяться, а AppendixFormatRule заваливал документ замечаниями.
"""
import unittest

from docx import Document

from docx_rules import ListsFormatRule
from docx_rules.classification import APPENDIX_START_REGEX, classify_paragraphs


def body_with_sentence(sentence):
    doc = Document()
    doc.add_paragraph('СОДЕРЖАНИЕ')
    doc.add_paragraph('ВВЕДЕНИЕ')
    doc.add_paragraph(sentence)
    doc.add_paragraph('Текст раздела.')
    return doc


class AppendixStartTests(unittest.TestCase):
    def test_body_sentence_does_not_start_appendix(self):
        classified = classify_paragraphs(body_with_sentence(
            'Приложение предназначено для использования сотрудниками бюро.'))
        self.assertFalse(classified[2].is_appendix)
        self.assertFalse(classified[3].is_appendix)

    def test_heading_starts_appendix(self):
        for header in ('Приложение А', 'ПРИЛОЖЕНИЕ', 'Приложение 1',
                       'Приложение А. Требования к составу', 'Приложение — Требования'):
            with self.subTest(header=header):
                self.assertTrue(APPENDIX_START_REGEX.match(header))
        for sentence in ('Приложение предназначено для использования',
                         'Приложение реализует функции учета',
                         'Приложение «Ритуальный помощник» готово к работе'):
            with self.subTest(sentence=sentence):
                self.assertIsNone(APPENDIX_START_REGEX.match(sentence))

    def test_lists_are_checked_after_body_sentence(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')
        doc.add_paragraph('Приложение предназначено для использования сотрудниками бюро.')
        manual = doc.add_paragraph('– первый пункт;')
        manual.paragraph_format.first_line_indent = __import__('docx').shared.Cm(0)
        manual2 = doc.add_paragraph('– второй пункт.')
        manual2.paragraph_format.first_line_indent = __import__('docx').shared.Cm(0)
        found = ListsFormatRule().check(doc)
        self.assertTrue(any('Отступ текста списка' in v.message for v in found),
                        'Список после фразы «Приложение …» должен проверяться')


if __name__ == '__main__':
    unittest.main()
