"""Проверка рисунков и подписей к ним."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective
from docx_rules.classification import classify_paragraphs


class FiguresRule(BaseRule):
    CAPTION_REGEX = re.compile(r'^Рис(унок)?\.?\s*(\d+)', re.IGNORECASE)
    REF_REGEX = re.compile(r'\(\s*см\.\s*рис\.?\s*(\d+)\s*\)', re.IGNORECASE)
    PLAIN_REF_REGEX = re.compile(r'\(\s*рис\.?\s*(\d+)', re.IGNORECASE)

    def _figure_kind(self, paragraph) -> Optional[str]:
        """Тип абзаца с рисунком: 'inline', 'anchored' или None."""
        xml = paragraph._element.xml
        if '<w:drawing' not in xml and '<w:pict' not in xml:
            return None
        has_visual = ('<pic:pic' in xml or 'c:chart' in xml
                      or '<v:imagedata' in xml or '<w:object' in xml)
        if not has_visual:
            return None  # Надписи и фигуры без изображения рисунком не считаются
        if '<wp:anchor' in xml:
            return 'anchored'
        return 'inline'

    def _is_centered(self, paragraph) -> bool:
        align = effective(paragraph, 'alignment')
        return align == WD_ALIGN_PARAGRAPH.CENTER

    def _has_red_line(self, paragraph) -> bool:
        indent = effective(paragraph, 'first_line_indent')
        return indent is not None and abs(indent.cm - 1.25) <= 0.05

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "FiguresRule"

        paragraphs = classify_paragraphs(doc)

        caption_numbers = []
        referenced_numbers = []

        for idx, paragraph in enumerate(paragraphs):
            if getattr(paragraph, 'is_title_page', False) or getattr(paragraph, 'is_toc', False):
                continue

            kind = self._figure_kind(paragraph)
            if kind is None:
                continue

            loc = f"Рисунок в абзаце {idx + 1}"

            # 1. Рисунок выравнивается по центру (только встроенные рисунки)
            if kind == 'inline' and not self._is_centered(paragraph):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Рисунок должен быть выровнен по центру."
                    )
                )

            if kind == 'inline' and not self._has_red_line(paragraph):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="У рисунка должна быть красная строка."
                    )
                )

            # 2. Подпись под рисунком: 'Рис. N. Название'
            caption_para = None
            caption_text = None
            own_lines = paragraph.text.split('\n')
            if len(own_lines) > 1 and own_lines[-1].strip():
                caption_text = own_lines[-1].strip()
            else:
                for nxt in range(idx + 1, len(paragraphs)):
                    if paragraphs[nxt].text.strip():
                        caption_para = paragraphs[nxt]
                        caption_text = caption_para.text.strip()
                        break

            if caption_text and self.CAPTION_REGEX.match(caption_text):
                num = int(self.CAPTION_REGEX.match(caption_text).group(2))
                caption_numbers.append(num)

                cap_para = caption_para if caption_para is not None else paragraph
                if not self._is_centered(cap_para):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Подпись рисунка (абзац {idx + 1}: '{caption_text[:30]}')",
                            message="Подпись к рисунку должна быть выровнена по центру."
                        )
                    )
                if not self._has_red_line(cap_para):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=f"Подпись рисунка (абзац {idx + 1}: '{caption_text[:30]}')",
                            message="У подписи к рисунку должна быть красная строка."
                        )
                    )
            else:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="После рисунка отсутствует подпись формата 'Рис. N. Название'."
                    )
                )

        # 3. Ссылки на рисунки: (см.рис. N) в конце абзаца перед точкой
        for idx, paragraph in enumerate(paragraphs):
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
                            location=f"Абзац {idx + 1} ('{text.strip()[:30]}...')",
                            message=f"Ссылка на рисунок '(см.рис. {num})' должна стоять в конце абзаца перед точкой."
                        )
                    )

            # Ссылка без 'см.': '(рис. 5)' — допустимая форма ссылки, также учитывается
            cleaned = self.REF_REGEX.sub('', text)
            for match in self.PLAIN_REF_REGEX.finditer(cleaned):
                referenced_numbers.append(int(match.group(1)))

        # 4. Соответствие ссылок и подписей
        for num in sorted(set(referenced_numbers)):
            if num not in caption_numbers:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location="Документ",
                        message=f"Ссылка на несуществующий рисунок {num} (подпись 'Рис. {num}.' не найдена)."
                    )
                )

        for num in set(caption_numbers):
            if num not in referenced_numbers:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location="Документ",
                        message=f"На рисунок {num} нет ссылки в тексте (должна быть: '(см.рис. {num})')."
                    )
                )

        # 5. Сквозная нумерация рисунков с 1 без пропусков
        expected = list(range(1, len(caption_numbers) + 1))
        if sorted(caption_numbers) != expected and caption_numbers:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Документ",
                    message=f"Нумерация рисунков должна быть сквозной начиная с 1 без пропусков (найдены: {sorted(caption_numbers)})."
                )
            )

        return violations
