"""Проверка оформления таблиц."""

import math
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective, full_text, has_field
from docx_rules.classification import classify_paragraphs


class TablesFormatRule(BaseRule):
    LABEL_REGEX = re.compile(r'^Таблица\s+(\d+)\s*$', re.IGNORECASE)
    CONTINUATION_REGEX = re.compile(r'^Продолжение таблицы\s+(\d+)\s*$', re.IGNORECASE)
    REF_REGEX = re.compile(r'\(\s*см\.\s*табл\.?\s*(\d+)\s*\)', re.IGNORECASE)
    NATURAL_REF_REGEX = re.compile(r'(?<![\w(])в\s+таблице\s+(\d+)\b', re.IGNORECASE)
    REQUIRED_BORDERS = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')

    # Оценка высоты таблицы: строка 14 пт при одинарном интервале и средняя
    # ширина символа Times New Roman 14 пт. Грубые значения — нужен только
    # признак «таблица заведомо не влезает в страницу» или, наоборот,
    # «таблица заведомо умещается».
    LINE_HEIGHT_CM = 0.6
    CHAR_WIDTH_CM = 0.25
    CELL_PADDING_CM = 0.2
    HEIGHT_SHARE = 0.85
    FIT_SHARE = 0.55

    def __init__(self, target_font_name: str = "Times New Roman", target_font_size: float = 14.0,
                 require_header_repeat: bool = True):
        self.font_name = target_font_name
        self.font_size = target_font_size
        # Требовать w:tblHeader (повторение шапки на новых страницах) можно не
        # всегда: в ТЗ по образцу таблицы короткие и повторение шапки не практикуется.
        self.require_header_repeat = require_header_repeat

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

    @classmethod
    def _missing_borders(cls, table) -> List[str]:
        """Границы с учётом прямых свойств и цепочки табличных стилей."""
        border_values = {}
        sources = [table._tbl.tblPr]
        style = table.style
        seen = set()
        while style is not None and style.style_id not in seen:
            seen.add(style.style_id)
            tbl_pr = style.element.find(qn('w:tblPr'))
            if tbl_pr is not None:
                sources.append(tbl_pr)
            style = style.base_style
        has_border_declaration = False
        for tbl_pr in sources:
            borders = tbl_pr.find(qn('w:tblBorders'))
            if borders is None:
                continue
            has_border_declaration = True
            for name in cls.REQUIRED_BORDERS:
                if name in border_values:
                    continue
                node = borders.find(qn(f'w:{name}'))
                if node is not None:
                    border_values[name] = node.get(qn('w:val'), 'single')
        # При полном отсутствии tblBorders итог может задаваться темой Word; не выдаём
        # ложную ошибку. Явная, но неполная/отключённая конфигурация проверяется строго.
        if not has_border_declaration:
            return []
        return [name for name in cls.REQUIRED_BORDERS
                if border_values.get(name) in (None, 'nil', 'none')]

    @staticmethod
    def _first_row_repeats(table) -> bool:
        if not table.rows:
            return False
        tr_pr = table.rows[0]._tr.find(qn('w:trPr'))
        if tr_pr is None:
            return False
        header = tr_pr.find(qn('w:tblHeader'))
        return header is not None and header.get(qn('w:val'), 'true') not in ('0', 'false', 'off')

    @staticmethod
    def _numbers_row_values(table, start: int) -> Optional[List[int]]:
        """Номера граф из строки вида «1 | 2 | …» (строка start) или None.

        Строка считается нумерацией граф, если все её непустые ячейки —
        целые числа и первая из них равна 1.
        """
        rows = table.rows
        if len(rows) <= start:
            return None
        values = []
        for cell in rows[start].cells:
            text = cell.text.strip().replace('\xa0', ' ')
            if not text:
                continue
            if not re.fullmatch(r'\d+', text):
                return None
            values.append(int(text))
        if len(values) >= 2 and values[0] == 1:
            return values
        return None

    def _usable_area_cm(self, doc: Document) -> Optional[tuple]:
        section = doc.sections[0] if doc.sections else None
        if section is None or section.page_height is None:
            return None
        return (section.page_height.cm - section.top_margin.cm - section.bottom_margin.cm,
                section.page_width.cm - section.left_margin.cm - section.right_margin.cm)

    def _estimate_height_cm(self, doc: Document, table) -> float:
        """Грубая оценка отрисованной высоты таблицы, см (без разбивки на страницы)."""
        widths = []
        grid = table._tbl.find(qn('w:tblGrid'))
        if grid is not None:
            for column in grid.findall(qn('w:gridCol')):
                try:
                    widths.append(int(column.get(qn('w:w'))) / 567.0)  # twips -> см
                except (TypeError, ValueError):
                    widths.append(0.0)
        area = self._usable_area_cm(doc)
        fallback_width = area[1] if area else 17.0
        total = 0.0
        for row in table.rows:
            row_height = 0.0
            for idx, cell in enumerate(row.cells):
                width = widths[idx] if idx < len(widths) and widths[idx] > 0.5 else fallback_width
                text_width = max(width - 0.4, 1.0)
                cell_height = 0.0
                for cell_para in cell.paragraphs:
                    text = cell_para.text.strip()
                    if not text:
                        continue
                    spacing = effective(cell_para, 'line_spacing')
                    factor = spacing if isinstance(spacing, float) else 1.0
                    lines = max(1, math.ceil(len(text) * self.CHAR_WIDTH_CM / text_width))
                    cell_height += lines * self.LINE_HEIGHT_CM * factor
                row_height = max(row_height, cell_height)
            total += row_height + self.CELL_PADDING_CM
        return total

    def _estimated_beyond_page(self, doc: Document, table) -> bool:
        """True, если оценённая высота таблицы заведомо больше страницы."""
        area = self._usable_area_cm(doc)
        if not area:
            return False
        return self._estimate_height_cm(doc, table) > area[0] * self.HEIGHT_SHARE

    def _estimated_fits_page(self, doc: Document, table) -> bool:
        """True, если таблица заведомо умещается на одной странице целиком."""
        area = self._usable_area_cm(doc)
        if not area:
            return False
        return self._estimate_height_cm(doc, table) <= area[0] * self.FIT_SHARE

    def _check_numbers_sequence(self, values, location, violations):
        if values != list(range(1, len(values) + 1)):
            violations.append(Violation(
                "TablesFormatRule", Severity.WARNING, location,
                "Нумерация граф должна быть последовательной от 1 "
                f"(найдено: {', '.join(map(str, values))})."))

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "TablesFormatRule"

        paragraphs = classify_paragraphs(doc)
        items = self._body_sequence(paragraphs, doc.tables)

        # Таблицы на титульном листе (лист согласования «Подп. и дата» по ГОСТ 19.104)
        # — часть титула, а не табличные данные: их оформление не проверяем и в
        # сквозную нумерацию таблиц они не входят.
        title_table_indexes = set()
        first_body_p = next((i for i, p in enumerate(paragraphs) if not p.is_title_page), None)
        if first_body_p is not None:
            for kind, obj, obj_idx in items:
                if kind == 'p' and obj_idx == first_body_p:
                    break
                if kind == 'tbl':
                    title_table_indexes.add(obj_idx)

        labeled_numbers = []
        referenced_numbers = []
        main_tables = {}
        continuation_tables = []

        for pos, (kind, obj, obj_idx) in enumerate(items):
            if kind != 'tbl':
                continue

            if obj_idx in title_table_indexes:
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
                    continuation_num = int(m.group(1))
                    is_continuation = True
                    if continuation_num not in labeled_numbers:
                        violations.append(Violation(
                            rule_name, Severity.WARNING,
                            f"Продолжение таблицы {continuation_num}",
                            (f"Неоднозначная надпись 'Продолжение таблицы {continuation_num}': "
                             "ранее таблица с таким номером не распознана. Проверьте номер и связь с исходной таблицей.")))
                    if back[0][1].alignment != WD_ALIGN_PARAGRAPH.RIGHT:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Продолжение таблицы {continuation_num}",
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

            missing_borders = self._missing_borders(table)
            if missing_borders:
                violations.append(Violation(
                    rule_name, Severity.ERROR, table_loc,
                    "У таблицы должны быть видимые внешние и внутренние границы "
                    f"(не найдены: {', '.join(missing_borders)})."))
            # У короткой таблицы нет перехода на следующую страницу; повторяемую шапку
            # требуем только там, где она потенциально нужна и включена в профиле.
            if self.require_header_repeat and len(table.rows) > 2 and not self._first_row_repeats(table):
                violations.append(Violation(
                    rule_name, Severity.WARNING, table_loc,
                    "Первая строка таблицы должна быть помечена w:tblHeader для повторения шапки на следующих страницах."))

            if is_continuation:
                continuation_tables.append((continuation_num, table))
                continue  # Название для продолжения не требуется; общие свойства уже проверены

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
                main_tables[num] = table

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

        # Строка нумерации граф («1 | 2 | …») нужна таблицам, выходящим за
        # пределы одной страницы (см. образец ТЗ): после шапки основной части
        # и в части «Продолжение таблицы N». Выход за страницу достоверен при
        # наличии «Продолжения», у одиночной длинной таблицы — оценкой высоты.
        # Обратная норма тоже проверяется: у таблицы, умещающейся на одной
        # странице, цифр под шапкой быть не должно (позицию на странице видно
        # только по грубой оценке высоты, поэтому — предупреждение).
        continuation_nums = {num for num, _ in continuation_tables}
        for num, table in main_tables.items():
            values = self._numbers_row_values(table, start=1)
            if values is None:
                if num in continuation_nums:
                    violations.append(Violation(
                        "TablesFormatRule", Severity.ERROR, f"Таблица {num}",
                        "Таблица выходит за пределы одной страницы (есть 'Продолжение "
                        "таблицы N'): после шапки нужна строка с нумерацией граф (1, 2, …)."))
                elif self._estimated_beyond_page(doc, table):
                    violations.append(Violation(
                        "TablesFormatRule", Severity.WARNING, f"Таблица {num}",
                        "Таблица, вероятно, выходит за пределы одной страницы: после шапки "
                        "нужна строка с нумерацией граф (1, 2, …), как в образце."))
                elif len(table.rows) >= 5:
                    violations.append(Violation(
                        "TablesFormatRule", Severity.WARNING, f"Таблица {num}",
                        "Если таблица выходит за пределы одной страницы, после шапки нужна "
                        "строка с нумерацией граф (1, 2, …), как в образце."))
            else:
                if num not in continuation_nums and self._estimated_fits_page(doc, table):
                    violations.append(Violation(
                        "TablesFormatRule", Severity.WARNING, f"Таблица {num}",
                        "Цифры под шапкой не нужны: таблица умещается на одной странице. "
                        "Нумерация граф (1, 2, …) ставится, только если таблица выходит "
                        "за пределы страницы."))
                self._check_numbers_sequence(values, f"Таблица {num}", violations)
        for num, table in continuation_tables:
            values = (self._numbers_row_values(table, start=0)
                      or self._numbers_row_values(table, start=1))
            if values is None:
                violations.append(Violation(
                    "TablesFormatRule", Severity.ERROR, f"Продолжение таблицы {num}",
                    "В части 'Продолжение таблицы N' нужна строка с нумерацией граф (1, 2, …)."))
            else:
                self._check_numbers_sequence(values, f"Продолжение таблицы {num}", violations)

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
            # Естественная ссылка может находиться внутри предложения: «данные приведены в таблице 2».
            referenced_numbers.extend(
                int(match.group(1)) for match in self.NATURAL_REF_REGEX.finditer(text))

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
