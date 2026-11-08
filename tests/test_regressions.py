import base64
import contextlib
import hashlib
import importlib
import io
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from docx_rules import (TextGridRule, MainTextFormatRule, PageNumberingRule,
                        TableOfContentsRule, SectionFormatRule, AppendixFormatRule,
                        FiguresRule, TablesFormatRule, FormulasRule, ListsFormatRule,
                        SourceReferencesRule, Severity, classify_paragraphs, effective,
                        has_page_field, full_text)


def grid(section, kind=None, pitch=None, space=None):
    for old in section._sectPr.findall(qn('w:docGrid')):
        section._sectPr.remove(old)
    node = OxmlElement('w:docGrid')
    for attr, value in [('type', kind), ('linePitch', pitch), ('charSpace', space)]:
        if value is not None:
            node.set(qn('w:' + attr), value)
    section._sectPr.append(node)
    return node


def body_style(doc):
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    style.font.bold = style.font.italic = False
    fmt = style.paragraph_format
    fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt.first_line_indent = Cm(1.25)
    fmt.line_spacing = 1.5
    fmt.space_before = fmt.space_after = Pt(0)
    fmt.widow_control = True
    return style


def simple_page(paragraph, instruction='PAGE'):
    node = OxmlElement('w:fldSimple')
    node.set(qn('w:instr'), instruction)
    run = OxmlElement('w:r')
    props = OxmlElement('w:rPr')
    size = OxmlElement('w:sz')
    size.set(qn('w:val'), '20')
    fonts = OxmlElement('w:rFonts')
    fonts.set(qn('w:ascii'), 'Times New Roman')
    fonts.set(qn('w:hAnsi'), 'Times New Roman')
    props.extend([fonts, size])
    run.append(props)
    node.append(run)
    paragraph._p.append(node)
    return node


class GridTests(unittest.TestCase):
    def test_missing_and_inactive(self):
        for kind in [None, 'none', 'default', 'unknown']:
            with self.subTest(kind=kind):
                doc = Document()
                grid(doc.sections[0], kind, '360')
                errors = TextGridRule().check(doc)
                self.assertTrue(errors)
                self.assertTrue(all(v.severity == Severity.ERROR for v in errors))
        doc.sections[0]._sectPr.remove(doc.sections[0]._sectPr.find(qn('w:docGrid')))
        self.assertTrue(TextGridRule().check(doc))

    def test_active_modes_signed_charspace(self):
        for kind in ['lines', 'linesAndChars', 'snapToChars']:
            for space in [None, '0', '-123', '+24']:
                with self.subTest(kind=kind, space=space):
                    doc = Document()
                    grid(doc.sections[0], kind, '360' if kind != 'snapToChars' else None, space)
                    self.assertEqual([], TextGridRule().check(doc))

    def test_invalid_numbers(self):
        for pitch in [None, '0', '-1', '1.5', 'abc']:
            for kind in ['lines', 'linesAndChars']:
                doc = Document()
                grid(doc.sections[0], kind, pitch)
                self.assertTrue(TextGridRule().check(doc), (kind, pitch))
        for space in ['', '1.2', '--1', 'abc']:
            doc = Document()
            grid(doc.sections[0], 'snapToChars', space=space)
            self.assertTrue(TextGridRule().check(doc))

    def test_duplicate_grid(self):
        doc = Document()
        grid(doc.sections[0], 'lines', '360')
        extra = OxmlElement('w:docGrid')
        extra.set(qn('w:type'), 'snapToChars')
        doc.sections[0]._sectPr.append(extra)
        self.assertTrue(TextGridRule().check(doc))

    def test_multiple_sections_roundtrip(self):
        doc = Document()
        grid(doc.sections[0], 'lines', '360', '-20')
        doc.add_section()
        grid(doc.sections[1], 'none')
        doc.add_section()
        grid(doc.sections[2], 'linesAndChars', '400', '0')
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'grid.docx'
            doc.save(path)
            errors = TextGridRule().check(Document(path))
            self.assertEqual(['Секция 2'], [v.location for v in errors])


class FormattingTests(unittest.TestCase):
    def test_no_toc_body_is_checked(self):
        doc = Document()
        doc.add_paragraph('Обычный текст без содержания')
        self.assertFalse(classify_paragraphs(doc)[0].is_title_page)
        self.assertTrue(MainTextFormatRule().check(doc))

    def test_inherited_body_and_direct_override(self):
        doc = Document()
        parent = body_style(doc)
        child = doc.styles.add_style('Child', WD_STYLE_TYPE.PARAGRAPH)
        child.base_style = parent
        grandchild = doc.styles.add_style('Grandchild', WD_STYLE_TYPE.PARAGRAPH)
        grandchild.base_style = child
        paragraph = doc.add_paragraph('Основной текст', grandchild)
        self.assertEqual([], MainTextFormatRule().check(doc))
        parent.paragraph_format.widow_control = False
        parent.paragraph_format.space_after = Pt(8)
        parent.font.bold = True
        errors = MainTextFormatRule().check(doc)
        for token in ['widow_control', 'после', 'bold']:
            self.assertTrue(any(token in v.message for v in errors), token)
        paragraph.paragraph_format.widow_control = True
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.runs[0].bold = False
        self.assertEqual([], MainTextFormatRule().check(doc))

    def test_character_style(self):
        doc = Document()
        body_style(doc)
        p = doc.add_paragraph('Текст')
        style = doc.styles.add_style('EmphasisChild', WD_STYLE_TYPE.CHARACTER)
        style.font.italic = True
        p.runs[0].style = style
        self.assertTrue(effective(p, 'italic', p.runs[0]))
        self.assertTrue(any('italic' in v.message for v in MainTextFormatRule().check(doc)))

    def test_special_paragraphs_not_main_text(self):
        doc = Document()
        for text in ['ВВЕДЕНИЕ', '1.1 Параграф', 'Рисунок 1 – подпись',
                     'Таблица 1', 'Название таблицы']:
            doc.add_paragraph(text)
        doc.add_table(rows=1, cols=1)
        doc.add_paragraph('– пункт списка')
        p = doc.add_paragraph('(1)')
        p._p.append(OxmlElement('m:oMathPara'))
        doc.add_paragraph('ПРИЛОЖЕНИЕ А')
        doc.add_paragraph('Текст приложения')
        self.assertEqual([], MainTextFormatRule().check(doc))

    def test_heading_indent_and_appendix_no_conflict(self):
        doc = Document()
        body_style(doc)
        p = doc.add_paragraph('ВВЕДЕНИЕ')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        doc.add_paragraph()
        self.assertEqual([], SectionFormatRule().check(doc))
        p.paragraph_format.first_line_indent = Cm(0)
        self.assertTrue(any('1.25' in v.message for v in SectionFormatRule().check(doc)))
        doc = Document()
        p = doc.add_paragraph('ПРИЛОЖЕНИЕ А')
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.runs[0].bold = True
        p = doc.add_paragraph('Текст приложения')
        p.runs[0].font.size = Pt(14)
        p.paragraph_format.line_spacing = 1
        self.assertEqual([], SectionFormatRule().check(doc))
        self.assertEqual([], AppendixFormatRule().check(doc))


class NumberingTests(unittest.TestCase):
    def test_header_docdefaults_without_direct_format(self):
        doc = Document()
        p = doc.sections[0].header.paragraphs[0]
        run = p.add_run('1')
        self.assertIsNotNone(effective(p, 'size', run))
        self.assertIsNone(effective(p, 'underline', run))

    def test_missing_page(self):
        doc = Document()
        doc.add_paragraph('Обычный абзац основного текста.')
        self.assertTrue(any('отсутствует поле PAGE' in v.message
                            for v in PageNumberingRule().check(doc)))

    def test_simple_page_and_false_positives(self):
        for instruction, expected in [('PAGE', True), (' PAGE \\* MERGEFORMAT', True),
                                      ('NUMPAGES', False), ('PAGEREF bookmark', False)]:
            doc = Document()
            simple_page(doc.sections[0].header.paragraphs[0], instruction)
            self.assertEqual(expected, has_page_field(doc.sections[0].header))

    def test_complex_split_instruction(self):
        doc = Document()
        p = doc.sections[0].header.paragraphs[0]
        for tag, value in [('fldChar', 'begin'), ('instrText', ' PA'),
                           ('instrText', 'GE '), ('fldChar', 'separate'), ('fldChar', 'end')]:
            node = OxmlElement('w:' + tag)
            if tag == 'fldChar':
                node.set(qn('w:fldCharType'), value)
            else:
                node.text = value
            p.add_run()._r.append(node)
        self.assertTrue(has_page_field(p))

    def test_simple_field_run_format_checked(self):
        doc = Document()
        sec = doc.sections[0]
        sec.different_first_page_header_footer = True
        p = sec.header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        field = simple_page(p)
        self.assertEqual([], PageNumberingRule().check(doc))
        field.find('.//' + qn('w:sz')).set(qn('w:val'), '28')
        self.assertTrue(any('Размер шрифта' in v.message for v in PageNumberingRule().check(doc)))


class IntegrationTests(unittest.TestCase):
    def test_import_and_defaults(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            main = importlib.import_module('main')
            importlib.reload(main)
        self.assertEqual('', output.getvalue())
        rules = main.default_rules()
        # TextGridRule не входит в требования МДК 04.01 и в стандартный набор
        self.assertEqual(0, sum(isinstance(r, TextGridRule) for r in rules))
        self.assertEqual(1, sum(isinstance(r, TableOfContentsRule) for r in rules))

    def test_toc_stops_at_chapter_and_spacing(self):
        doc = Document()
        p = doc.add_paragraph('СОДЕРЖАНИЕ')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.runs[0].bold = True
        entry = doc.add_paragraph('1.1 Название\t3')
        entry.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        entry.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT)
        doc.add_paragraph('ГЛАВА 1 ТЕОРИЯ')
        self.assertEqual([], TableOfContentsRule().check(doc))
        for text in ['1.1  Название\t3', '1.1Название\t3', '1.1\tНазвание\t3']:
            entry.text = text
            self.assertTrue(any('строго один пробел' in v.message for v in TableOfContentsRule().check(doc)))

    def test_sample_smoke_no_modification(self):
        from main import DocxValidator, default_rules
        path = Path(__file__).resolve().parents[1] / 'docs' / 'test.docx'
        before = hashlib.sha256(path.read_bytes()).digest()
        report = DocxValidator(default_rules()).validate(str(path))
        self.assertTrue(report)  # тестовый документ не соответствует требованиям
        self.assertEqual(before, hashlib.sha256(path.read_bytes()).digest())


W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
M_NS = ('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')
PNG_1PX = base64.b64decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==')


def messages(violations):
    return [v.message for v in violations]


def right_tab(paragraph, position_cm=16.5):
    paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(position_cm), WD_TAB_ALIGNMENT.RIGHT)


class PageNumberingSectionsTests(unittest.TestCase):
    def build(self):
        doc = Document()
        doc.add_paragraph('ТЕМА ДИПЛОМНОГО ПРОЕКТА')
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('Введение\t3')
        doc.sections[0].different_first_page_header_footer = True
        doc.add_section()
        doc.add_paragraph('ВВЕДЕНИЕ')
        doc.add_paragraph('Основной текст введения.')
        header = doc.sections[1].header.paragraphs[0]
        header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        simple_page(header, 'PAGE')
        doc.add_section()
        appendix_header = doc.add_paragraph('ПРИЛОЖЕНИЕ А')
        appendix_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        appendix_header.runs[0].bold = True
        appendix_body = doc.add_paragraph('Текст приложения')
        appendix_body.paragraph_format.line_spacing = 1.0
        doc.sections[2].header.is_linked_to_previous = False
        doc.sections[2].footer.is_linked_to_previous = False
        return doc

    def test_positive_per_section(self):
        doc = self.build()
        self.assertEqual([], PageNumberingRule().check(doc))
        self.assertEqual([], AppendixFormatRule().check(doc))

    def test_main_section_without_page(self):
        doc = self.build()
        header = doc.sections[1].header.paragraphs[0]
        for node in header._p.xpath('.//w:fldSimple'):
            header._p.remove(node)
        self.assertTrue(any('отсутствует поле PAGE' in m for m in messages(PageNumberingRule().check(doc))))

    def test_page_in_footer(self):
        doc = self.build()
        footer = doc.sections[1].footer
        footer.is_linked_to_previous = False
        simple_page(footer.paragraphs[0], 'PAGE')
        self.assertTrue(any('не в нижнем' in m for m in messages(PageNumberingRule().check(doc))))

    def test_even_pages(self):
        doc = self.build()
        doc.settings.odd_and_even_pages_header_footer = True
        self.assertTrue(any('чётных страниц' in m for m in messages(PageNumberingRule().check(doc))))
        even = doc.sections[1].even_page_header
        even.is_linked_to_previous = False
        paragraph = even.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        simple_page(paragraph, 'PAGE')
        self.assertFalse(any('чётных страниц' in m for m in messages(PageNumberingRule().check(doc))))

    def test_extra_empty_header_lines(self):
        doc = self.build()
        doc.sections[1].header.add_paragraph('')
        self.assertTrue(any('лишние пустые строки' in m for m in messages(PageNumberingRule().check(doc))))

    def test_number_font_size(self):
        doc = self.build()
        node = doc.sections[1].header.paragraphs[0]._p.xpath('.//w:sz')[0]
        node.set(qn('w:val'), '28')
        self.assertTrue(any('Размер шрифта нумерации' in m for m in messages(PageNumberingRule().check(doc))))


class TableOfContentsTabsTests(unittest.TestCase):
    def build(self, tab_cm=16.5):
        doc = Document()
        header = doc.add_paragraph('СОДЕРЖАНИЕ')
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header.paragraph_format.first_line_indent = Cm(1.25)
        header.runs[0].bold = True
        for text in ['1.1 Название\t3', 'Глава 1 Теория\t7']:
            entry = doc.add_paragraph(text)
            entry.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if tab_cm is not None:
                right_tab(entry, tab_cm)
        doc.add_paragraph('ГЛАВА 1 ТЕОРИЯ')
        doc.add_paragraph('1.1 Параграф тела')
        return doc

    def test_positive_manual_toc(self):
        self.assertEqual([], TableOfContentsRule().check(self.build()))

    def test_entry_without_tab(self):
        doc = self.build(tab_cm=None)
        self.assertTrue(any('табулятор' in m for m in messages(TableOfContentsRule().check(doc))))

    def test_different_tab_positions(self):
        doc = Document()
        header = doc.add_paragraph('СОДЕРЖАНИЕ')
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header.runs[0].bold = True
        for text, pos in [('1.1 Название\t3', 16.5), ('Глава 1 Теория\t7', 15.0)]:
            entry = doc.add_paragraph(text)
            entry.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            right_tab(entry, pos)
        doc.add_paragraph('ГЛАВА 1 ТЕОРИЯ')
        doc.add_paragraph('1.1 Параграф тела')
        self.assertTrue(any('одной позиции' in m for m in messages(TableOfContentsRule().check(doc))))

    def test_toc_field_skips_manual_checks(self):
        doc = Document()
        header = doc.add_paragraph('СОДЕРЖАНИЕ')
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header.paragraph_format.first_line_indent = Cm(1.25)
        header.runs[0].bold = True
        paragraph = doc.add_paragraph()
        fld = OxmlElement('w:fldSimple')
        fld.set(qn('w:instr'), 'TOC \\o "1-3"')
        run = OxmlElement('w:r')
        node = OxmlElement('w:t')
        node.text = '1.1 Название....3'
        run.append(node)
        fld.append(run)
        paragraph._p.append(fld)
        doc.add_paragraph('ГЛАВА 1 ТЕОРИЯ')
        doc.add_paragraph('1.1 Параграф тела')
        self.assertEqual([], TableOfContentsRule().check(doc))

    def test_heading_missing_from_toc(self):
        doc = self.build()
        doc.add_paragraph('ГЛАВА 2 ПРАКТИКА')
        doc.add_paragraph('2.1 Отдельный параграф')
        found = messages(TableOfContentsRule().check(doc))
        self.assertTrue(any("Глава 2' нет в содержании" in m for m in found))
        self.assertTrue(any('2.1 нет в содержании' in m for m in found))


class AppendixMultipleTests(unittest.TestCase):
    def build(self):
        doc = Document()
        for letter in ['А', 'Б']:
            header = doc.add_paragraph(f'ПРИЛОЖЕНИЕ {letter}')
            header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            header.runs[0].bold = True
            body = doc.add_paragraph('Текст приложения')
            body.paragraph_format.line_spacing = 1.0
        return doc

    def test_positive_multiple_appendices(self):
        self.assertEqual([], AppendixFormatRule().check(self.build()))

    def test_non_sequential_letters(self):
        doc = Document()
        for letter in ['А', 'В']:
            header = doc.add_paragraph(f'ПРИЛОЖЕНИЕ {letter}')
            header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            header.runs[0].bold = True
            doc.add_paragraph('Текст')
        self.assertTrue(any('последовательно' in m for m in messages(AppendixFormatRule().check(doc))))

    def test_header_not_right(self):
        doc = self.build()
        doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.assertTrue(any('по правому краю' in m for m in messages(AppendixFormatRule().check(doc))))

    def test_page_in_appendix_footer(self):
        doc = self.build()
        doc.add_section()
        appendix_header = doc.add_paragraph('ПРИЛОЖЕНИЕ В')
        appendix_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        appendix_header.runs[0].bold = True
        doc.add_paragraph('Текст')
        section = doc.sections[-1]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        simple_page(section.footer.paragraphs[0], 'PAGE')
        self.assertTrue(any('Нижний колонтитул' in v.location and 'не должны нумероваться' in v.message
                            for v in AppendixFormatRule().check(doc)))


class FiguresExtendedTests(unittest.TestCase):
    def add_figure(self, doc, caption, indent=Cm(1.25)):
        figure = doc.add_paragraph()
        figure.alignment = WD_ALIGN_PARAGRAPH.CENTER
        figure.paragraph_format.first_line_indent = indent
        figure.add_run().add_picture(io.BytesIO(PNG_1PX))
        if caption is not None:
            cap = doc.add_paragraph(caption)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.first_line_indent = indent
        return figure

    def test_positive_figure(self):
        doc = Document()
        self.add_figure(doc, 'Рис. 1. Схема')
        doc.add_paragraph('Текст со ссылкой (см.рис. 1).')
        self.assertEqual([], FiguresRule().check(doc))

    def test_plain_reference_counts(self):
        doc = Document()
        self.add_figure(doc, 'Рис. 1. Схема')
        doc.add_paragraph('Текст со ссылкой (рис. 1).')
        self.assertEqual([], FiguresRule().check(doc))

    def test_numbering_gap_is_error(self):
        doc = Document()
        self.add_figure(doc, 'Рис. 3. Схема')
        doc.add_paragraph('Текст со ссылкой (см.рис. 3).')
        found = FiguresRule().check(doc)
        self.assertTrue(any(v.severity == Severity.ERROR and 'сквозной' in v.message for v in found))

    def test_missing_reference_warns(self):
        doc = Document()
        self.add_figure(doc, 'Рис. 1. Схема')
        doc.add_paragraph('Текст без ссылки.')
        self.assertTrue(any('нет ссылки' in m for m in messages(FiguresRule().check(doc))))

    def test_wrong_red_line(self):
        doc = Document()
        self.add_figure(doc, 'Рис. 1. Схема', indent=None)
        doc.add_paragraph('Текст со ссылкой (см.рис. 1).')
        self.assertTrue(any('красная строка' in m for m in messages(FiguresRule().check(doc))))


class TablesExtendedTests(unittest.TestCase):
    def prepare(self, doc):
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)
        style.paragraph_format.line_spacing = 1.0
        label = doc.add_paragraph('Таблица 1')
        label.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        title = doc.add_paragraph('Название таблицы')
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.paragraph_format.first_line_indent = Cm(1.25)
        return doc.add_table(rows=2, cols=2)

    def fill(self, table, header_center=True, valign=True, numeric_center=True, text_left=True):
        for idx, text in enumerate(['Показатель', 'Значение']):
            cell = table.rows[0].cells[idx]
            cell.paragraphs[0].text = text
            if header_center:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if valign:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        values = [('12.5', numeric_center), ('Данные', None)]
        for idx, (text, center) in enumerate(values):
            cell = table.rows[1].cells[idx]
            cell.paragraphs[0].text = text
            if center or (not text_left and not text.replace('.', '').isdigit()):
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    def test_positive_table(self):
        doc = Document()
        table = self.prepare(doc)
        self.fill(table)
        doc.add_paragraph('Ссылка на таблицу (см.табл. 1).')
        self.assertEqual([], TablesFormatRule().check(doc))

    def test_header_not_centered_horizontally(self):
        doc = Document()
        table = self.prepare(doc)
        self.fill(table, header_center=False)
        self.assertTrue(any('по горизонтали (jc)' in m for m in messages(TablesFormatRule().check(doc))))

    def test_header_not_centered_vertically(self):
        doc = Document()
        table = self.prepare(doc)
        self.fill(table, valign=False)
        self.assertTrue(any('по вертикали (vAlign)' in m for m in messages(TablesFormatRule().check(doc))))

    def test_numeric_cell_left(self):
        doc = Document()
        table = self.prepare(doc)
        self.fill(table, numeric_center=False)
        self.assertTrue(any('Числовые ячейки' in m for m in messages(TablesFormatRule().check(doc))))

    def test_text_cell_centered(self):
        doc = Document()
        table = self.prepare(doc)
        self.fill(table, text_left=False)
        self.assertTrue(any('Текстовые ячейки' in m for m in messages(TablesFormatRule().check(doc))))


class FormulasExtendedTests(unittest.TestCase):
    def add_formula(self, doc, jc='center', number=True, math_text='E=mc2'):
        paragraph = doc.add_paragraph()
        pr = f'<m:oMathParaPr><m:jc m:val="{jc}"/></m:oMathParaPr>' if jc else ''
        paragraph._p.append(parse_xml(
            f'<m:oMathPara {M_NS}>{pr}<m:oMath><m:r><m:t>{math_text}</m:t></m:r></m:oMath></m:oMathPara>'))
        if number:
            paragraph.add_run('(1)')
        return paragraph

    def test_positive_formula(self):
        doc = Document()
        doc.add_paragraph('Расчет ведется по формуле:')
        self.add_formula(doc)
        doc.add_paragraph('где E — энергия.')
        self.assertEqual([], FormulasRule().check(doc))

    def test_default_omathpara_is_centered(self):
        doc = Document()
        doc.add_paragraph('Расчет ведется по формуле:')
        self.add_formula(doc, jc=None)
        doc.add_paragraph('где E — энергия.')
        self.assertEqual([], FormulasRule().check(doc))

    def test_left_math_jc_is_error(self):
        doc = Document()
        doc.add_paragraph('Расчет ведется по формуле:')
        self.add_formula(doc, jc='left')
        doc.add_paragraph('где E — энергия.')
        self.assertTrue(any('по центру' in m for m in messages(FormulasRule().check(doc))))

    def test_missing_number(self):
        doc = Document()
        doc.add_paragraph('Расчет ведется по формуле:')
        self.add_formula(doc, number=False)
        doc.add_paragraph('где E — энергия.')
        self.assertTrue(any('нумерация' in m for m in messages(FormulasRule().check(doc))))

    def test_number_inside_math_counts(self):
        doc = Document()
        doc.add_paragraph('Расчет ведется по формуле:')
        paragraph = doc.add_paragraph()
        paragraph._p.append(parse_xml(
            f'<m:oMathPara {M_NS}><m:oMath><m:r><m:t>x+y</m:t></m:r>'
            f'<m:r><m:t>(2)</m:t></m:r></m:oMath></m:oMathPara>'))
        doc.add_paragraph('где x — аргумент.')
        self.assertFalse(any('нумерация' in m for m in messages(FormulasRule().check(doc))))

    def test_description_without_colon(self):
        doc = Document()
        doc.add_paragraph('Расчет ведется по формуле.')
        self.add_formula(doc)
        doc.add_paragraph('где E — энергия.')
        self.assertTrue(any('двоеточием' in m for m in messages(FormulasRule().check(doc))))


class AutoNumberedListsTests(unittest.TestCase):
    def add_list_definition(self, doc, abstract_id, num_id, num_fmt, lvl_text,
                            ind_left=None, ind_first=None, ind_hanging=None):
        numbering = doc.part.numbering_part.element
        attrs = ''
        if ind_left is not None:
            attrs += f' w:left="{ind_left}"'
        if ind_first is not None:
            attrs += f' w:firstLine="{ind_first}"'
        if ind_hanging is not None:
            attrs += f' w:hanging="{ind_hanging}"'
        ind = f'<w:pPr><w:ind{attrs}/></w:pPr>' if attrs else ''
        numbering.insert(0, parse_xml(
            f'<w:abstractNum {W_NS} w:abstractNumId="{abstract_id}">'
            f'<w:multiLevelType w:val="singleLevel"/>'
            f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{num_fmt}"/>'
            f'<w:lvlText w:val="{lvl_text}"/><w:lvlJc w:val="left"/>{ind}</w:lvl>'
            f'</w:abstractNum>'))
        numbering.append(parse_xml(
            f'<w:num {W_NS} w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num>'))

    def auto_paragraph(self, doc, text, num_id, first=Cm(1.75), left=None):
        paragraph = doc.add_paragraph(text)
        paragraph.paragraph_format.first_line_indent = first
        if left is not None:
            paragraph.paragraph_format.left_indent = left
        num_pr = parse_xml(f'<w:numPr {W_NS}><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>')
        paragraph._p.pPr.insert(0, num_pr)
        return paragraph

    def test_positive_auto_lists(self):
        doc = Document()
        self.add_list_definition(doc, 101, 201, 'bullet', '–')
        self.add_list_definition(doc, 102, 202, 'decimal', '%1)')
        self.auto_paragraph(doc, 'маркированный пункт;', 201)
        self.auto_paragraph(doc, 'маркированный пункт.', 201)
        self.auto_paragraph(doc, 'первый пункт;', 202)
        self.auto_paragraph(doc, 'последний пункт.', 202)
        self.assertEqual([], ListsFormatRule().check(doc))

    def test_wrong_bullet_char(self):
        doc = Document()
        self.add_list_definition(doc, 100, 200, 'bullet', '•')
        self.auto_paragraph(doc, 'пункт списка;', 200)
        found = ListsFormatRule().check(doc)
        self.assertTrue(any(v.severity == Severity.ERROR and 'Маркер списка' in v.message for v in found))

    def test_decimal_paren_punctuation(self):
        doc = Document()
        self.add_list_definition(doc, 102, 202, 'decimal', '%1)')
        self.auto_paragraph(doc, 'первый пункт.', 202)
        self.auto_paragraph(doc, 'последний пункт;', 202)
        found = messages(ListsFormatRule().check(doc))
        self.assertTrue(any('точкой с запятой' in m for m in found))
        self.assertTrue(any('точкой (последний)' in m for m in found))

    def test_decimal_dot_punctuation(self):
        doc = Document()
        self.add_list_definition(doc, 103, 203, 'decimal', '%1.')
        self.auto_paragraph(doc, 'первый пункт;', 203)
        self.assertTrue(any('должен заканчиваться точкой' in m for m in messages(ListsFormatRule().check(doc))))

    def test_auto_indent_wrong(self):
        doc = Document()
        self.add_list_definition(doc, 104, 204, 'decimal', '%1.')
        self.auto_paragraph(doc, 'пункт списка.', 204, left=Cm(1.0))
        self.assertTrue(any('Отступ авто-списка' in m for m in messages(ListsFormatRule().check(doc))))

    def test_indent_from_numbering_level(self):
        doc = Document()
        self.add_list_definition(doc, 105, 205, 'decimal', '%1.', ind_first='992')
        self.auto_paragraph(doc, 'пункт списка.', 205, first=None)
        self.assertFalse(any('Отступ авто-списка' in m for m in messages(ListsFormatRule().check(doc))))
        self.add_list_definition(doc, 106, 206, 'decimal', '%1.', ind_left='720', ind_hanging='720')
        self.auto_paragraph(doc, 'пункт списка.', 206, first=None)
        self.assertTrue(any('Отступ авто-списка' in m for m in messages(ListsFormatRule().check(doc))))

    def test_numfmt_none_skipped(self):
        doc = Document()
        self.add_list_definition(doc, 107, 207, 'none', '%1.')
        self.auto_paragraph(doc, 'пункт списка;', 207)
        self.assertEqual([], ListsFormatRule().check(doc))


class SourceReferencesBibliographyTests(unittest.TestCase):
    def build(self):
        doc = Document()
        doc.add_paragraph('ВВЕДЕНИЕ')
        doc.add_paragraph('Текст со ссылкой [9].')
        doc.add_paragraph('Текст со ссылкой [2].')
        doc.add_paragraph('Текст со ссылкой [3, с.5].')
        doc.add_paragraph('СПИСОК ИСТОЧНИКОВ')
        doc.add_paragraph('1. Иванов И.И. Книга. – М.: Издательство, 2020. – 100 с.')
        doc.add_paragraph('2. Петров П.П. Другая книга. – СПб.: Нева, 2021. – 50 с.')
        doc.add_paragraph('3. Сидоров С.С. Сайт. URL: http://example.com (дата обращения: 01.02.2020).')
        return doc

    def test_missing_source_is_error(self):
        found = SourceReferencesRule().check(self.build())
        self.assertTrue(any(v.severity == Severity.ERROR and 'отсутствует в списке' in v.message for v in found))

    def test_page_in_electronic_warns(self):
        found = SourceReferencesRule().check(self.build())
        self.assertTrue(any('электронный ресурс' in m for m in messages(found)))

    def test_valid_links_have_no_other_violations(self):
        found = [v for v in SourceReferencesRule().check(self.build())
                 if 'отсутствует' not in v.message and 'электронный ресурс' not in v.message]
        self.assertEqual([], found)

    def test_no_bibliography_skips_crosscheck(self):
        doc = Document()
        doc.add_paragraph('ВВЕДЕНИЕ')
        doc.add_paragraph('Текст со ссылкой [9].')
        self.assertEqual([], SourceReferencesRule().check(doc))


class CliErrorHandlingTests(unittest.TestCase):
    ROOT = Path(__file__).resolve().parents[1]

    def run_cli(self, target):
        return subprocess.run([sys.executable, str(self.ROOT / 'main.py'), str(target)],
                              capture_output=True, text=True, timeout=180)

    def test_missing_file_exit_2(self):
        result = self.run_cli(self.ROOT / 'docs' / 'no_such_file.docx')
        self.assertEqual(2, result.returncode)
        self.assertNotIn('Traceback', result.stderr)
        self.assertNotIn('Traceback', result.stdout)

    def test_corrupt_file_exit_2(self):
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as handle:
            handle.write(b'not a docx at all')
            path = Path(handle.name)
        try:
            result = self.run_cli(path)
            self.assertEqual(2, result.returncode)
            self.assertNotIn('Traceback', result.stderr)
            self.assertNotIn('Traceback', result.stdout)
        finally:
            path.unlink()


if __name__ == '__main__':
    unittest.main()
