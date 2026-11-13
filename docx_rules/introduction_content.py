"""Проверка обязательных элементов введения."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import is_toc_line
from docx_rules.classification import classify_paragraphs


class IntroductionContentRule(BaseRule):
    MAJOR_HEADER_REGEX = r'^(ГЛАВА\s+\d+|СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)'

    # Обязательные элементы введения (слайд 203)
    REQUIRED_ELEMENTS = [
        ("Теоретическое описание", r'ТЕОРЕТИЧЕСК'),
        ("Актуальность", r'АКТУАЛЬНОСТ'),
        ("Тема", r'\bТЕМ[АЕУОЙ]\b'),
        ("Объект", r'\bОБЪЕКТ'),
        ("Предмет", r'\bПРЕДМЕТ'),
        ("Цель", r'\bЦЕЛЬ'),
        ("Задачи", r'\bЗАДАЧ'),
        ("Методы", r'\bМЕТОД'),
        ("Теоретическая и практическая значимость", r'ЗНАЧИМОСТ'),
        ("Структура работы", r'\bСТРУКТУР'),
    ]

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "IntroductionContentRule"

        classify_paragraphs(doc)

        intro_start = None
        for idx, p in enumerate(classify_paragraphs(doc)):
            text = p.text.strip()
            if getattr(p, 'is_title_page', False) or getattr(p, 'is_toc', False) or is_toc_line(text):
                continue
            if text.upper() == "ВВЕДЕНИЕ":
                intro_start = idx + 1
                break

        if intro_start is None:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Документ",
                    message="Раздел 'Введение' не найден в документе."
                )
            )
            return violations

        intro_parts = []
        for p in doc.paragraphs[intro_start:]:
            text = p.text.strip()
            text_upper = text.upper()

            # Конец введения: крупный заголовок либо любой центрированный абзац
            # (в некоторых работах разделы идут без слова 'ГЛАВА')
            if text and re.match(self.MAJOR_HEADER_REGEX, text_upper):
                break

            align = p.alignment
            if align is None and p.style and p.style.paragraph_format:
                align = p.style.paragraph_format.alignment
            if text and align == WD_ALIGN_PARAGRAPH.CENTER:
                break

            intro_parts.append(text)
        intro_text = " ".join(intro_parts).upper()

        for element_name, pattern in self.REQUIRED_ELEMENTS:
            if not re.search(pattern, intro_text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location="Введение",
                        message=f"Во введении не найден обязательный элемент: {element_name}."
                    )
                )

        return violations
