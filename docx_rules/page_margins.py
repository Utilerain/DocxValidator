"""Проверка полей страницы."""

import re

from docx import Document
from typing import List

from docx_rules.base import BaseRule, Severity, Violation


class PageMarginsRule(BaseRule):
    def __init__(self, top=2.0, bottom=2.0, left=3.0, right=1.5):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right

    def check(self, doc: Document) -> List[Violation]:
        violations = []
        for i, section in enumerate(doc.sections, 1):
            top_cm = round(section.top_margin.cm, 2) if section.top_margin else None
            bottom_cm = round(section.bottom_margin.cm, 2) if section.bottom_margin else None
            left_cm = round(section.left_margin.cm, 2) if section.left_margin else None
            right_cm = round(section.right_margin.cm, 2) if section.right_margin else None

            if top_cm != self.top or bottom_cm != self.bottom or left_cm != self.left or right_cm != self.right:
                violations.append(Violation(
                    rule_name="Поля страницы",
                    severity=Severity.ERROR,
                    location="Общее",
                    message=f"Ожидалось: Л: {self.left}cм, П: {self.right}cм, В: {self.top}cм, Н: {self.bottom}см.\n" +
                            f"Фактически: Л:{left_cm}см, П:{right_cm}см, В:{top_cm}см, Н:{bottom_cm}см."
                ))
        return violations
