"""Проверка оформления таблиц."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective
from docx_rules.classification import classify_paragraphs


class TablesFormatRule(BaseRule):
    LABEL_REGEX = re.compile(r'^Таблица\s+(\d+)\s*$', re.IGNORECASE)
    CONTINUATION_REGEX = re.compile(r'^Продолжение таблицы\s+(\d+)\s*$', re.IGNORECASE)
    REF_REGEX = re.compile(r'\(\s*см\.\s*табл\.?\s*(\d+)\s*\)', re.IGNORECASE)

    def __init__(self, target_font_name: str = "Times New Roman", target_font_size: float = 14.0):
        self.font_name = target_font_name
        self.font_size = target_font_size

    @staticmethod
    def _body_sequence(paragraphs: List, tables: List):
        """Последовательность ('p', абзац, индекс) / ('tbl', таблица, индекс) в порядке следования."""
        items = []
        p_idx = 0
        t_idx = 0
        para_iter = iter(paragraphs)
        tbl_iter = iter(tables)
        if paragraphs:
            body = paragraphs[0]._element.getparent()
            for child in body.iterchildren():
                if child.tag == qn('w:p'):
                    para = next(para_iter, None)
                    if para is not None:
                        items.append(('p', para, p_idx))
                        p_idx += 1
                elif child.tag == qn('w:tbl'):
                    table = next(tbl_iter, None)
                    if table is not None:
                        items.append(('tbl', table, t_idx))
                        t_idx += 1
        return items

    def _get_effective_font_name(self, run, paragraph) -> Optional[str]:
        return effective(paragraph, 'name', run)

    def _get_effective_font_size(self, run, paragraph) -> Optional[float]:
        size = effective(paragraph, 'size', run)
        return size.pt if size is not None else None

    def _is_centered(self, paragraph) -> bool:
        align = effective(paragraph, 'alignment')
        return align == WD_ALIGN_PARAGRAPH.CENTER

    def _has_red_line(self, paragraph) -> bool:
        indent = paragraph.paragraph_format.first_line_indent
        if indent is None and paragraph.style and paragraph.style.paragraph_format:
            indent = paragraph.style.paragraph_format.first_line_indent
        return indent is not None and indent.cm > 0.05

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "TablesFormatRule"

        paragraphs = classify_paragraphs(doc)
        items = self._body_sequence(paragraphs, doc.tables)

        labeled_numbers = []
        referenced_numbers = []

        for pos, (kind, obj, obj_idx) in enumerate(items):
            if kind != 'tbl':
                continue

            table = obj
            table_loc = f"Таблица №{obj_idx + 1} по порядку в документе"

            # Контекст: до 3 непустых абзацев перед таблицей
            back = []
            j = pos - 1
            while j >= 0 and len(back) < 3:
                if items[j][0] == 'p' and items[j][1].text.strip():
                    back.append(items[j])
                j -= 1

            label_para = None
            title_para = None
            is_continuation = False

            if back:
                m = self.CONTINUATION_REGEX.match(back[0][1].text.strip())
                if m:
                    is_continuation = True
                    if back[0][1].alignment != WD_ALIGN_PARAGRAPH.RIGHT:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Продолжение таблицы {m.group(1)}",
                                message="Надпись 'Продолжение таблицы N' должна быть выровнена по правому краю."
                            )
                        )
                else:
                    m = self.LABEL_REGEX.match(back[0][1].text.strip())
                    if m:
                        label_para = back[0][1]
                    elif len(back) > 1:
                        m = self.LABEL_REGEX.match(back[1][1].text.strip())
                        if m:
                            label_para = back[1][1]
                            title_para = back[0][1]

            if is_continuation:
                continue  # Продолжение таблицы: надпись и название не требуются

            if label_para is None:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=table_loc,
                        message="Перед таблицей не найдена надпись 'Таблица N' (выравнивание по правому краю)."
                    )
                )
            else:
                num = int(self.LABEL_REGEX.match(label_para.text.strip()).group(1))
                labeled_numbers.append(num)

                if label_para.alignment != WD_ALIGN_PARAGRAPH.RIGHT:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Надпись 'Таблица {num}'",
                            message="Надпись 'Таблица N' должна быть выровнена по правому краю."
                        )
                    )

                if title_para is None:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Таблица {num}",
                            message="После надписи 'Таблица N' отсутствует название таблицы."
                        )
                    )
                else:
                    title_text = title_para.text.strip()[:30]
                    if not self._is_centered(title_para):
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Название таблицы {num} ('{title_text}')",
                                message="Название таблицы должно быть выровнено по центру."
                            )
                        )
                    if not self._has_red_line(title_para):
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=f"Название таблицы {num} ('{title_text}')",
                                message="У названия таблицы должна быть красная строка."
                            )
                        )

            # Шрифт и интервал в ячейках таблицы
            bad_font = []
            bad_spacing = []
            for row in table.rows:
                for cell in row.cells:
                    for cell_para in cell.paragraphs:
                        if not cell_para.text.strip():
                            continue
                        line_sp = cell_para.paragraph_format.line_spacing
                        if line_sp is None and cell_para.style and cell_para.style.paragraph_format:
                            line_sp = cell_para.style.paragraph_format.line_spacing
                        if line_sp is not None and abs(float(line_sp) - 1.0) > 0.05:
                            bad_spacing.append(cell_para.text.strip()[:20])
                        for run in cell_para.runs:
                            if not run.text.strip():
                                continue
                            font_name = self._get_effective_font_name(run, cell_para)
                            if font_name and font_name != self.font_name:
                                bad_font.append(f"шрифт '{font_name}'")
                            font_size = self._get_effective_font_size(run, cell_para)
                            if font_size and abs(font_size - self.font_size) > 0.1:
                                bad_font.append(f"размер {font_size} pt")

            if bad_font:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=table_loc,
                        message=f"В таблице должен быть шрифт {self.font_name} {int(self.font_size)} pt "
                                f"(обнаружено: {', '.join(sorted(set(bad_font))[:3])})."
                    )
                )

            if bad_spacing:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=table_loc,
                        message=f"Междустрочный интервал в таблице должен быть 1 "
                                f"(обнаружен отличный в {len(bad_spacing)} фрагментах)."
                    )
                )

            # Ячейки данных: числовые — по центру, текстовые — по левому краю (WARN)
            numeric_regex = re.compile(r'^[+-]?\d+(?:[.,]\d+)?$')
            bad_numeric = []
            bad_text = []
            for row in table.rows[1:]:
                for cell in row.cells:
                    for cell_para in cell.paragraphs:
                        text = cell_para.text.strip()
                        if not text:
                            continue
                        align = effective(cell_para, 'alignment')
                        if align is None:
                            align = WD_ALIGN_PARAGRAPH.LEFT  # отсутствие jc означает левый край
                        compact = text.replace(' ', '').replace('\xa0', '')
                        if numeric_regex.fullmatch(compact):
                            if align != WD_ALIGN_PARAGRAPH.CENTER:
                                bad_numeric.append(text[:20])
                        else:
                            if align != WD_ALIGN_PARAGRAPH.LEFT:
                                bad_text.append(text[:20])
            if bad_numeric:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=table_loc,
                        message=f"Числовые ячейки таблицы должны быть выровнены по центру "
                                f"(обнаружено проблем: {len(bad_numeric)}, например: {', '.join(bad_numeric[:3])})."
                    )
                )
            if bad_text:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=table_loc,
                        message=f"Текстовые ячейки таблицы должны быть выровнены по левому краю "
                                f"(обнаружено проблем: {len(bad_text)}, например: {', '.join(bad_text[:3])})."
                    )
                )

            # Шапка (первая строка): по центру по горизонтали (jc) и вертикали (vAlign)
            if table.rows:
                bad_horizontal = []
                bad_vertical = []
                for cell in table.rows[0].cells:
                    cell_para = cell.paragraphs[0] if cell.paragraphs else None
                    if cell_para is None:
                        continue
                    align = effective(cell_para, 'alignment')
                    if align != WD_ALIGN_PARAGRAPH.CENTER:
                        bad_horizontal.append(cell.text.strip()[:20])
                    if cell.vertical_alignment != WD_ALIGN_VERTICAL.CENTER:
                        bad_vertical.append(cell.text.strip()[:20])
                if bad_horizontal:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=table_loc,
                            message=f"Названия столбцов должны быть выровнены по центру по горизонтали (jc); "
                                    f"проблемы в: {', '.join(bad_horizontal[:3])}."
                        )
                    )
                if bad_vertical:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=table_loc,
                            message=f"Названия столбцов должны быть выровнены по центру по вертикали (vAlign); "
                                    f"проблемы в: {', '.join(bad_vertical[:3])}."
                        )
                    )

        # Ссылки на таблицы: (см.табл. N) в конце абзаца перед точкой
        for p_idx, paragraph in enumerate(paragraphs, start=1):
            if getattr(paragraph, 'is_title_page', False) or getattr(paragraph, 'is_toc', False):
                continue
            text = paragraph.text
            if not text.strip():
                continue

            for match in self.REF_REGEX.finditer(text):
                num = int(match.group(1))
                referenced_numbers.append(num)
                tail = text[match.end():].strip()
                if tail not in ("", "."):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=f"Абзац {p_idx} ('{text.strip()[:30]}...')",
                            message=f"Ссылка на таблицу '(см.табл. {num})' должна стоять в конце абзаца перед точкой."
                        )
                    )

        for num in sorted(set(referenced_numbers)):
            if num not in labeled_numbers:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location="Документ",
                        message=f"Ссылка на несуществующую таблицу {num} (надпись 'Таблица {num}' не найдена)."
                    )
                )

        for num in set(labeled_numbers):
            if num not in referenced_numbers:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location="Документ",
                        message=f"На таблицу {num} нет ссылки в тексте (должна быть: '(см.табл. {num})')."
                    )
                )

        expected = list(range(1, len(labeled_numbers) + 1))
        if labeled_numbers and sorted(labeled_numbers) != expected:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.WARNING,
                    location="Документ",
                    message=f"Нумерация таблиц должна быть сквозной начиная с 1 (найдены: {sorted(labeled_numbers)})."
                )
            )

        return violations
