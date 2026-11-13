"""Проверка форматирования основного текста."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective
from docx_rules.classification import classify_paragraphs


class MainTextFormatRule(BaseRule):
    def __init__(
        self,
        target_font_name: str = "Times New Roman",
        target_font_size: float = 14.0,
        target_line_spacing: float = 1.5,
        target_first_line_indent_cm: float = 1.25,
        target_space_before_pt: float = 0.0,
        target_space_after_pt: float = 0.0,
        allow_formatting_styles: bool = False,  # Разрешать ли bold/italic
        indent_tolerance_cm: float = 0.05       # Допустимая погрешность отступа
    ):
        self.font_name = target_font_name
        self.font_size = target_font_size
        self.line_spacing = target_line_spacing
        self.first_line_indent_cm = target_first_line_indent_cm
        self.space_before_pt = target_space_before_pt
        self.space_after_pt = target_space_after_pt
        self.allow_formatting_styles = allow_formatting_styles
        self.indent_tolerance_cm = indent_tolerance_cm

    def _get_effective_font_name(self, run, paragraph) -> Optional[str]:
        return effective(paragraph, 'name', run)

    def _get_effective_font_size(self, run, paragraph) -> Optional[float]:
        size = effective(paragraph, 'size', run)
        return size.pt if size is not None else None

    def _get_effective_alignment(self, paragraph):
        return effective(paragraph, 'alignment')

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "MainTextFormatRule"

        paragraphs = classify_paragraphs(doc)

        for p_idx, paragraph in enumerate(paragraphs, start=1):
            if paragraph.is_title_page or paragraph.is_toc or paragraph.is_special:
                continue  # Пропускаем титульный лист и оглавление

            text = paragraph.text.strip()
            if not text:
                continue

            loc = f"Абзац {p_idx} ('{text[:30]}...')"
            fmt = paragraph.paragraph_format

            # 1. Выравнивание (по ширине / JUSTIFY)
            alignment = self._get_effective_alignment(paragraph)
            if alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Выравнивание текста должно быть по ширине (JUSTIFY)."
                    )
                )

            # 2. Красная строка (1,25 см)
            indent = effective(paragraph, 'first_line_indent')

            if indent is None:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Не задана красная строка (ожидается {self.first_line_indent_cm} см)."
                    )
                )
            else:
                indent_cm = indent.cm
                if abs(indent_cm - self.first_line_indent_cm) > self.indent_tolerance_cm:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message=f"Отступ красной строки равен {round(indent_cm, 2)} см, а должен быть {self.first_line_indent_cm} см."
                        )
                    )

            # 3. Междустрочный интервал (1.5)
            line_sp = effective(paragraph, 'line_spacing')

            if line_sp is None or abs(float(line_sp) - self.line_spacing) > 0.05:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Междустрочный интервал {line_sp}, ожидается {self.line_spacing}."
                    )
                )

            # 4. Отступы до и после абзаца (0 pt)
            space_before = effective(paragraph, 'space_before')
            space_before = space_before.pt if space_before is not None else 0.0
            space_after = effective(paragraph, 'space_after')
            space_after = space_after.pt if space_after is not None else 0.0

            if abs(space_before - self.space_before_pt) > 0.1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Отступ перед абзацем равен {space_before} pt (ожидается {self.space_before_pt} pt)."
                    )
                )

            if abs(space_after - self.space_after_pt) > 0.1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Отступ после абзаца равен {space_after} pt (ожидается {self.space_after_pt} pt)."
                    )
                )

            # 5. Запрет висячих строк
            if effective(paragraph, 'widow_control') is False:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Отключен запрет висячих строк (widow_control)."
                    )
                )

            # 6. Проверка свойств текста (Runs): шрифт, кегль, начертания
            for r_idx, run in enumerate(paragraph.runs, start=1):
                run_text = run.text.strip()
                if not run_text:
                    continue

                run_loc = f"{loc}, фрагмент {r_idx} ('{run_text[:15]}')"

                font_name = self._get_effective_font_name(run, paragraph)
                if font_name != self.font_name:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=run_loc,
                            message=f"Шрифт '{font_name}', а должен быть '{self.font_name}'."
                        )
                    )

                font_size = self._get_effective_font_size(run, paragraph)
                if font_size is None or abs(font_size - self.font_size) > 0.1:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=run_loc,
                            message=f"Размер шрифта {font_size} pt, а должен быть {self.font_size} pt."
                        )
                    )

                if not self.allow_formatting_styles:
                    if effective(paragraph, 'bold', run):
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=run_loc,
                                message="В тексте использовано полужирное начертание (bold)."
                            )
                        )
                    if effective(paragraph, 'italic', run):
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=run_loc,
                                message="В тексте использован курсив (italic)."
                            )
                        )
                    if effective(paragraph, 'underline', run):
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=run_loc,
                                message="В тексте использовано подчеркивание (underline)."
                            )
                        )

        return violations
