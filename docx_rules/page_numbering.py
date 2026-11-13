"""Проверка нумерации страниц по секциям."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective, has_page_field, paragraph_section_indices
from docx_rules.classification import classify_paragraphs


class PageNumberingRule(BaseRule):
    def __init__(
        self,
        target_font_name: str = "Times New Roman",
        target_font_size: float = 10.0,
        target_alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ):
        self.font_name = target_font_name
        self.font_size = target_font_size
        self.alignment = target_alignment

    def _has_page_number_field(self, header_or_paragraph) -> bool:
        return has_page_field(header_or_paragraph)

    def _get_effective_font_name(self, run, paragraph) -> Optional[str]:
        return effective(paragraph, 'name', run)

    def _get_effective_font_size(self, run, paragraph) -> Optional[float]:
        size = effective(paragraph, 'size', run)
        return size.pt if size is not None else None

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "PageNumberingRule"

        if not doc.sections:
            return violations

        paragraphs = classify_paragraphs(doc)
        section_indices = paragraph_section_indices(doc)

        # Контент по секциям: основной текст / приложение (титул и содержание не в счёт)
        section_content: dict = {}
        for paragraph, sec_idx in zip(paragraphs, section_indices):
            if paragraph.is_title_page or paragraph.is_toc:
                continue
            has_visual = bool(paragraph._p.xpath('.//w:drawing|.//w:pict'))
            if not (paragraph.text.strip() or has_visual):
                continue
            info = section_content.setdefault(sec_idx, {'main': False, 'appendix': False})
            if paragraph.is_appendix:
                info['appendix'] = True
            else:
                info['main'] = True

        for sec_idx, section in enumerate(doc.sections, start=1):
            info = section_content.get(sec_idx - 1, {})
            has_main = info.get('main', False)

            # Секции приложений проверяет AppendixFormatRule (PAGE там быть не должен)
            if info.get('appendix', False):
                continue

            header = section.header

            # 1. В секции с основным текстом PAGE обязан быть в верхнем колонтитуле
            if has_main and not self._has_page_number_field(header):
                violations.append(Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location=f"Секция {sec_idx} (Верхний колонтитул)",
                    message="В секции с основным текстом отсутствует поле PAGE в верхнем колонтитуле."
                ))

            # 2. Чётные страницы: при включённой опции PAGE должен быть и в чётном колонтитуле
            if has_main and doc.settings.odd_and_even_pages_header_footer \
                    and not self._has_page_number_field(section.even_page_header):
                violations.append(Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location=f"Секция {sec_idx} (Верхний колонтитул чётных страниц)",
                    message="Включена нумерация чётных страниц, но поле PAGE отсутствует в верхнем колонтитуле чётных страниц."
                ))

            # 3. Нижние колонтитулы не должны содержать нумерацию
            footers = [section.footer]
            if section.different_first_page_header_footer:
                footers.append(section.first_page_footer)
            if doc.settings.odd_and_even_pages_header_footer:
                footers.append(section.even_page_footer)
            for footer in footers:
                if footer.is_linked_to_previous:
                    continue  # проверяется в секции, где задан колонтитул
                if self._has_page_number_field(footer):
                    violations.append(Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Секция {sec_idx} (Нижний колонтитул)",
                        message="Нумерация страниц должна располагаться в верхнем колонтитуле, а не в нижнем."
                    ))

            # 4. Лишние пустые строки в колонтитуле
            if (self._has_page_number_field(header) or has_main) and len(header.paragraphs) > 1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Секция {sec_idx} (Верхний колонтитул)",
                        message=f"В верхнем колонтитуле обнаружены лишние пустые строки (всего строк: {len(header.paragraphs)}, должна быть 1)."
                    )
                )

            # 5. Формат нумерации (шрифт, кегль, выравнивание)
            if not self._has_page_number_field(header):
                continue

            for p_idx, paragraph in enumerate(header.paragraphs, start=1):
                if "\n" in paragraph.text:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}",
                            message="В строке нумерации присутствуют лишние переносы строк."
                        )
                    )

                alignment = paragraph.alignment
                if alignment is None and paragraph.style and paragraph.style.paragraph_format:
                    alignment = paragraph.style.paragraph_format.alignment

                if alignment != self.alignment:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}",
                            message="Нумерация страниц должна быть выровнена по правому краю (правый верхний угол)."
                        )
                    )

                for r_idx, node in enumerate(paragraph._p.xpath('.//w:r'), start=1):
                    from docx.text.run import Run
                    run = Run(node, paragraph)
                    font_name = self._get_effective_font_name(run, paragraph)
                    if font_name and font_name != self.font_name:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}, фрагмент {r_idx}",
                                message=f"Шрифт нумерации '{font_name}', а должен быть '{self.font_name}'."
                            )
                        )

                    font_size = self._get_effective_font_size(run, paragraph)
                    if font_size and abs(font_size - self.font_size) > 0.1:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}, фрагмент {r_idx}",
                                message=f"Размер шрифта нумерации {font_size} pt, а должен быть {self.font_size} pt."
                            )
                        )

        # Титульный лист (первая секция) не нумеруется
        first_section = doc.sections[0]
        has_title = any(p.is_title_page and p.text.strip() for p in paragraphs)
        if has_title:
            if first_section.different_first_page_header_footer:
                if self._has_page_number_field(first_section.first_page_header):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location="Секция 1 (Колонтитул первой страницы)",
                            message="На титульном листе присутствует нумерация страниц."
                        )
                    )
            else:
                if self._has_page_number_field(first_section.header):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location="Секция 1 (Верхний колонтитул)",
                            message="Не включена опция 'Особый колонтитул для первой страницы'. На титульном листе не должно быть нумерации."
                        )
                    )

        return violations
