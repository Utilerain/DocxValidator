"""Проверка оформления формул."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective, full_text
from docx_rules.classification import classify_paragraphs


class FormulasRule(BaseRule):
    NUMBERING_REGEX = re.compile(r'\((\d+)\)\s*$')
    M_JC_REGEX = re.compile(r'<m:jc\s+m:val="([^"]+)"')

    def _is_display_formula(self, paragraph) -> bool:
        """Выносная формула: oMathPara либо короткий абзац с oMath."""
        xml = paragraph._element.xml
        if '<m:oMath' not in xml:
            return False
        if '<m:oMathPara' in xml:
            return True
        text = self.NUMBERING_REGEX.sub('', paragraph.text).strip()
        return len(text) <= 25

    def _is_centered(self, paragraph) -> bool:
        """Формула центрирована: jc абзаца либо m:jc выносной формулы (по умолчанию centerGroup)."""
        if effective(paragraph, 'alignment') == WD_ALIGN_PARAGRAPH.CENTER:
            return True
        xml = paragraph._element.xml
        if '<m:oMathPara' in xml:
            match = self.M_JC_REGEX.search(xml)
            return match is None or match.group(1) in ('center', 'centerGroup')
        return False

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "FormulasRule"

        paragraphs = classify_paragraphs(doc)
        formula_indices = set()

        for idx, paragraph in enumerate(paragraphs):
            if not self._is_display_formula(paragraph):
                continue
            formula_indices.add(idx)

        formula_numbers = []

        for idx in sorted(formula_indices):
            paragraph = paragraphs[idx]
            text = paragraph.text.strip()
            loc = f"Формула (абзац {idx + 1}: '{text[:30]}')"

            # 1. Формула выравнивается по центру (jc абзаца или m:jc oMathPara)
            if not self._is_centered(paragraph):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Формула должна быть выровнена по центру."
                    )
                )

            # 2. Справа проставляется нумерация, например (1); номер может быть и в m:t
            match = self.NUMBERING_REGEX.search(full_text(paragraph).strip())
            if match:
                formula_numbers.append(int(match.group(1)))
            else:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Справа от формулы должна проставляться нумерация, например (1)."
                    )
                )

            # 3. Перед формулой описание, заканчивающееся двоеточием
            #    (в предыдущем абзаце либо в этом же абзаце перед формулой)
            prev_idx = idx - 1
            while prev_idx >= 0 and (not paragraphs[prev_idx].text.strip() or prev_idx in formula_indices):
                prev_idx -= 1
            own_text = self.NUMBERING_REGEX.sub('', full_text(paragraph)).strip()
            prev_ok = prev_idx >= 0 and paragraphs[prev_idx].text.strip().endswith(':')
            if not prev_ok and not own_text.endswith(':'):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Перед формулой должно быть описание, заканчивающееся двоеточием."
                    )
                )

            # 4. После формулы описываются примененные обозначения
            next_idx = idx + 1
            while next_idx < len(paragraphs) and (not paragraphs[next_idx].text.strip() or next_idx in formula_indices):
                next_idx += 1
            if next_idx < len(paragraphs):
                next_text = paragraphs[next_idx].text.strip().lower()
                if not next_text.startswith('где'):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message="После формулы должны быть описаны примененные обозначения (например: 'где: ...')."
                        )
                    )

        expected = list(range(1, len(formula_numbers) + 1))
        if formula_numbers and sorted(formula_numbers) != expected:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.WARNING,
                    location="Документ",
                    message=f"Нумерация формул должна быть сквозной начиная с 1 (найдены: {sorted(formula_numbers)})."
                )
            )

        return violations
