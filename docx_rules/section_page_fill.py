"""Проверка наполнения разделов текстом."""

import re
import math

from docx import Document
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import is_toc_line
from docx_rules.classification import classify_paragraphs


class SectionPageFillRule(BaseRule):
    MAJOR_HEADER_REGEX = r'^(ГЛАВА\s+\d+|СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)|ПРИЛОЖЕНИЕ)'

    def __init__(
        self,
        min_fill_ratio: float = 0.70,
        chars_per_line: int = 70,
        lines_per_full_page: int = 28
    ):
        self.min_fill_ratio = min_fill_ratio
        self.chars_per_line = chars_per_line
        self.lines_per_full_page = lines_per_full_page
        self.min_lines_required = math.ceil(lines_per_full_page * min_fill_ratio)

    def _has_page_break(self, paragraph) -> bool:
        if paragraph.paragraph_format.page_break_before:
            return True
        xml_str = paragraph._element.xml
        return 'w:br' in xml_str and 'w:type="page"' in xml_str

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "SectionPageFillRule"

        paragraphs = classify_paragraphs(doc)

        section_blocks = []
        current_block = []

        for idx, p in enumerate(paragraphs):
            text = p.text.strip()
            text_upper = text.upper()

            # Игнорируем титульный лист и оглавление
            if getattr(p, 'is_title_page', False) or getattr(p, 'is_toc', False) or is_toc_line(text):
                continue

            is_major = bool(re.match(self.MAJOR_HEADER_REGEX, text_upper, re.IGNORECASE))
            has_break = self._has_page_break(p)

            if (is_major or has_break) and current_block:
                section_blocks.append((idx, current_block))
                current_block = []

            if text or 'w:drawing' in p._element.xml or 'w:pict' in p._element.xml:
                current_block.append(p)

        for end_idx, block_paragraphs in section_blocks:
            if not block_paragraphs:
                continue

            total_estimated_lines = 0

            for p in block_paragraphs:
                p_text = p.text.strip()

                if p_text:
                    lines_in_p = max(1, math.ceil(len(p_text) / self.chars_per_line))
                    total_estimated_lines += lines_in_p

                xml_str = p._element.xml
                if 'w:drawing' in xml_str or 'w:pict' in xml_str:
                    total_estimated_lines += 10

            fill_ratio = min(1.0, total_estimated_lines / self.lines_per_full_page)

            if fill_ratio < self.min_fill_ratio:
                last_text = block_paragraphs[-1].text.strip()[:30]
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Конец раздела перед абзацем {end_idx} ('{last_text}...')",
                        message=(
                            f"При окончании раздела страница заполнена ориентировочно на {int(fill_ratio * 100)}% "
                            f"(~{total_estimated_lines} из {self.lines_per_full_page} строк). "
                            f"Допустимо заполнение листа не менее, чем на {int(self.min_fill_ratio * 100)}%."
                        )
                    )
                )

        return violations
