"""Проверка оформления заголовков глав и параграфов."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective, is_toc_line
from docx_rules.classification import classify_paragraphs


class SectionFormatRule(BaseRule):
    MAJOR_HEADER_REGEX = r'^(ГЛАВА\s+\d+|СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|БИБЛИОГРАФИЯ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)'
    SUBHEADING_REGEX = r'^\d+(\.\d+)+\.?\s+'

    def _is_empty(self, paragraph) -> bool:
        return not paragraph.text.strip()

    def _is_bold(self, paragraph) -> bool:
        runs_with_text = [r for r in paragraph.runs if r.text.strip()]
        if runs_with_text:
            return all(effective(paragraph, 'bold', r) for r in runs_with_text)
        if paragraph.style and paragraph.style.font:
            return paragraph.style.font.bold or False
        return False

    def _is_centered(self, paragraph) -> bool:
        align = effective(paragraph, 'alignment')
        return align == WD_ALIGN_PARAGRAPH.CENTER

    def _has_red_line(self, paragraph) -> bool:
        fmt = paragraph.paragraph_format
        indent = fmt.first_line_indent
        if indent is None and paragraph.style and paragraph.style.paragraph_format:
            indent = paragraph.style.paragraph_format.first_line_indent
        indent = effective(paragraph, 'first_line_indent')
        return indent is not None and abs(indent.cm - 1.25) <= 0.05

    def _has_page_break_before(self, paragraphs: List, idx: int) -> bool:
        if idx == 0:
            return True

        p = paragraphs[idx]

        if effective(p, 'page_break_before'):
            return True

        prev_p = paragraphs[idx - 1]
        xml_str = prev_p._element.xml
        if 'w:br' in xml_str and 'w:type="page"' in xml_str:
            return True

        if 'w:sectPr' in xml_str:
            return True

        return False

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "SectionFormatRule"

        paragraphs = classify_paragraphs(doc)

        for idx, p in enumerate(paragraphs):
            text = p.text.strip()
            if not text:
                continue

            # Игнорируем титульный лист и строки оглавления
            if p.is_title_page or (p.is_toc and not p.is_toc_heading) or p.is_appendix or is_toc_line(text):
                continue

            text_upper = text.upper()
            is_major = bool(re.match(self.MAJOR_HEADER_REGEX, text_upper, re.IGNORECASE))
            is_sub = bool(re.match(self.SUBHEADING_REGEX, text, re.IGNORECASE))

            if not (is_major or is_sub):
                continue

            header_type = "Глава/Крупный раздел" if is_major else "Параграф"
            loc = f"{header_type} (абзац {idx+1}: '{text[:30]}...')"

            # 1. Проверка полужирного начертания
            if not self._is_bold(p):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"{header_type} должен быть выделен полужирным шрифтом."
                    )
                )

            # 2. Проверка выравнивания по центру
            if not self._is_centered(p):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Заголовок ({header_type.lower()}) должен быть выровнен по центру."
                    )
                )

            if not self._has_red_line(p):
                violations.append(Violation(rule_name, Severity.ERROR, loc,
                                            'Красная строка заголовка должна быть 1.25 см.'))

            # 4. Проверка правила «с новой страницы»
            starts_new_page = self._has_page_break_before(paragraphs, idx)
            if is_major and not starts_new_page:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Крупные разделы (Глава, Содержание, Введение, Заключение, Список источников) должны начинаться с новой страницы."
                    )
                )
            elif is_sub and starts_new_page:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Параграф должен продолжаться с места окончания предыдущего (не должен принудительно начинаться с новой страницы)."
                    )
                )

            # 5. Проверка Enter'ов перед заголовком
            if is_sub:
                prev_non_empty_idx = idx - 1
                while prev_non_empty_idx >= 0 and self._is_empty(paragraphs[prev_non_empty_idx]):
                    prev_non_empty_idx -= 1

                prev_p_text = paragraphs[prev_non_empty_idx].text.strip().upper() if prev_non_empty_idx >= 0 else ""
                after_chapter = bool(re.match(self.MAJOR_HEADER_REGEX, prev_p_text, re.IGNORECASE))

                empty_count_before = 0
                check_i = idx - 1
                while check_i >= 0 and self._is_empty(paragraphs[check_i]):
                    empty_count_before += 1
                    check_i -= 1

                if after_chapter:
                    if empty_count_before != 1:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=loc,
                                message=f"После названия главы перед параграфом должен быть ровно 1 Enter (обнаружено пустых строк: {empty_count_before})."
                            )
                        )
                else:
                    if empty_count_before != 2 and not starts_new_page:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=loc,
                                message=f"Перед началом параграфа должно ставиться 2 Enter'а (обнаружено пустых строк: {empty_count_before})."
                            )
                        )

            # 6. Проверка Enter'а после заголовка
            empty_count_after = 0
            check_after = idx + 1
            while check_after < len(paragraphs) and self._is_empty(paragraphs[check_after]):
                empty_count_after += 1
                check_after += 1

            if empty_count_after != 1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"После названия {header_type.lower()} должен оставаться ровно 1 Enter (обнаружено пустых строк: {empty_count_after})."
                    )
                )

        return violations
