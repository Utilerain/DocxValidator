# -*- coding: utf-8 -*-
"""Строка нумерации граф («1 | 2 | …») у таблиц, выходящих за пределы страницы.

По образцу ТЗ (ПР2) таблицы, выходящие за пределы страницы, имеют строку с
номерами граф после шапки и повторяют её в части «Продолжение таблицы N».
Выход за страницу достоверен при наличии «Продолжения»; у одиночной длинной
таблицы — только эвристическая оценка высоты (WARN).
"""
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from docx_rules import TablesFormatRule


def make_doc():
    doc = Document()
    doc.add_paragraph('СОДЕРЖАНИЕ')
    doc.add_paragraph('ВВЕДЕНИЕ')
    ref = doc.add_paragraph('Работы выполняются поэтапно (см.табл. 1).')
    ref.paragraph_format.line_spacing = 1.5
    return doc


def add_labeled_table(doc, rows_values, after_ref=True):
    """Таблица 1 (по правому краю надпись + название по центру) из списков строк."""
    label = doc.add_paragraph('Таблица 1')
    label.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title = doc.add_paragraph('Наименование таблицы')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Cm(1.25)
    table = doc.add_table(rows=len(rows_values), cols=len(rows_values[0]))
    for r_idx, row_values in enumerate(rows_values):
        for c_idx, value in enumerate(row_values):
            table.cell(r_idx, c_idx).text = value
    return table


def add_continuation(doc, rows_values):
    label = doc.add_paragraph('Продолжение таблицы 1')
    label.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table = doc.add_table(rows=len(rows_values), cols=len(rows_values[0]))
    for r_idx, row_values in enumerate(rows_values):
        for c_idx, value in enumerate(row_values):
            table.cell(r_idx, c_idx).text = value
    return table


class TableNumbersRowTests(unittest.TestCase):
    def rule(self):
        return TablesFormatRule(require_header_repeat=False)

    def messages(self, doc):
        return [v.message for v in self.rule().check(doc)]

    def locations(self, doc):
        return [v.location for v in self.rule().check(doc)]

    def test_table_with_continuation_and_numbers_rows_passes(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы'], ['1', '2'],
                                ['А', 'Б'], ['В', 'Г']])
        add_continuation(doc, [['1', '2'], ['Д', 'Е']])
        found = self.messages(doc)
        self.assertFalse(any('нумерацией граф' in m for m in found), found)

    def test_continuation_without_numbers_row_is_error(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы'], ['1', '2'], ['А', 'Б']])
        add_continuation(doc, [['Д', 'Е']])
        joined = ' '.join(self.messages(doc))
        self.assertIn("В части 'Продолжение таблицы N' нужна строка с нумерацией граф", joined)

    def test_main_part_without_numbers_row_is_error_when_continued(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы'], ['А', 'Б']])
        add_continuation(doc, [['1', '2'], ['Д', 'Е']])
        joined = ' '.join(self.messages(doc))
        self.assertIn('после шапки нужна строка с нумерацией граф', joined)

    def test_non_sequential_numbers_are_reported(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы'], ['1', '2'], ['А', 'Б']])
        add_continuation(doc, [['1', '4', '4']])
        joined = ' '.join(self.messages(doc))
        self.assertIn('Нумерация граф должна быть последовательной', joined)

    def test_long_single_table_without_numbers_is_warning(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы']] + [[f'Этап {i}', 'Разработка и отладка модулей системы'] for i in range(1, 31)])
        joined = ' '.join(self.messages(doc))
        self.assertIn('вероятно, выходит за пределы одной страницы', joined)

    def test_candidate_table_gets_conditional_reminder(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы']] + [[f'Этап {i}', 'Работы этапа'] for i in range(1, 6)])
        joined = ' '.join(self.messages(doc))
        self.assertIn('Если таблица выходит за пределы одной страницы', joined)

    def test_short_single_table_without_numbers_passes(self):
        doc = make_doc()
        add_labeled_table(doc, [['Стадия', 'Работы'], ['А', 'Б'], ['В', 'Г']])
        found = self.messages(doc)
        self.assertFalse(any('нумерацией граф' in m for m in found), found)


if __name__ == '__main__':
    unittest.main()
