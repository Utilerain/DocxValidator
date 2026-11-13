"""Проверка оформления приложений."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective, has_page_field, paragraph_section_indices
from docx_rules.classification import classify_paragraphs


class AppendixFormatRule(BaseRule):
    HEADER_REGEX = r'^ПРИЛОЖЕНИЕ'
    LETTER_REGEX = re.compile(r'^ПРИЛОЖЕНИЕ\s+([А-Я])\b')
    # Русский алфавит без Ё — порядок букв приложений по ГОСТ 7.32
    ALPHABET = 'АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ'

    def _has_page_number_field(self, header) -> bool:
        return has_page_field(header)

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "AppendixFormatRule"

        paragraphs = classify_paragraphs(doc)
        section_indices = paragraph_section_indices(doc)

        appendix_indices = [idx for idx, paragraph in enumerate(paragraphs)
                            if not paragraph.is_title_page and not paragraph.is_toc
                            and re.match(self.HEADER_REGEX, paragraph.text.strip().upper())]
        if not appendix_indices:
            return violations  # Приложений в документе нет

        # 1. Заголовки ВСЕХ приложений: правое выравнивание, полужирный, последовательные буквы
        prev_letter = None
        for idx in appendix_indices:
            header_para = paragraphs[idx]
            header_text = header_para.text.strip()
            loc = f"Приложение (абзац {idx + 1}: '{header_text[:30]}')"

            align = effective(header_para, 'alignment')
            if align != WD_ALIGN_PARAGRAPH.RIGHT:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Заголовок приложения должен быть выровнен по правому краю."
                    )
                )

            is_bold = any(effective(header_para, 'bold', run) for run in header_para.runs if run.text.strip())
            if not is_bold and header_para.style and header_para.style.font:
                is_bold = bool(header_para.style.font.bold)
            if not is_bold:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Заголовок приложения должен быть выделен полужирным шрифтом."
                    )
                )

            letter_match = self.LETTER_REGEX.match(header_text.upper())
            if letter_match and letter_match.group(1) in self.ALPHABET:
                letter = letter_match.group(1)
                if prev_letter is not None:
                    expected = self.ALPHABET[self.ALPHABET.index(prev_letter) + 1] \
                        if self.ALPHABET.index(prev_letter) + 1 < len(self.ALPHABET) else None
                    if letter != expected:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=loc,
                                message=f"Буквы приложений должны идти последовательно: "
                                        f"после 'ПРИЛОЖЕНИЕ {prev_letter}' ожидается "
                                        f"'ПРИЛОЖЕНИЕ {expected}', найдено 'ПРИЛОЖЕНИЕ {letter}'."
                            )
                        )
                prev_letter = letter

        # 2. Межстрочный интервал 1 у текста приложений (заголовки проверяются отдельно)
        for offset, paragraph in enumerate(paragraphs[appendix_indices[0] + 1:], start=appendix_indices[0] + 1):
            text = paragraph.text.strip()
            if not text or re.match(self.HEADER_REGEX, text.upper()):
                continue

            p_loc = f"Приложение (абзац {offset + 1}: '{text[:30]}...')"

            line_sp = effective(paragraph, 'line_spacing')
            if line_sp is not None and abs(float(line_sp) - 1.0) > 0.05:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=p_loc,
                        message=f"Междустрочный интервал в приложении {line_sp}, а должен быть 1."
                    )
                )

        # 3. Страницы приложений не нумеруются (ни в верхнем, ни в нижнем колонтитуле)
        appendix_section = section_indices[appendix_indices[0]]
        for sec_idx, section in enumerate(doc.sections):
            if sec_idx < appendix_section:
                continue
            for holder, label in [(section.header, 'Верхний колонтитул'),
                                  (section.footer, 'Нижний колонтитул')]:
                if self._has_page_number_field(holder):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Секция {sec_idx + 1} ({label})",
                            message="В приложении страницы не должны нумероваться."
                        )
                    )

        return violations
