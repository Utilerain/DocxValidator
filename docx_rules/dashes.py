"""Проверка тире и дефисов в тексте."""

import re

from docx import Document
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.classification import classify_paragraphs


class DashesRule(BaseRule):
    # Строки листингов кода не проверяются (математический минус, дефисы в идентификаторах)
    CODE_LINE_REGEX = re.compile(
        r'^\s*('
        r'import\s|from\s+[A-Za-z_][\w.]*\s+import|'
        r'[A-Za-z_]\w*\s*(\[[^\]]*\])?\s*=[^=]|'
        r'print\s*\(|plt\.|def\s+[A-Za-z_]\w*\s*\(|return\b|'
        r'for\s+[A-Za-z_]\w*\s+in\s|while\s|elif\b|else:|#'
        r')'
    )
    EM_DASH_REGEX = re.compile(r'—')

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "DashesRule"

        paragraphs = classify_paragraphs(doc)

        in_code = False
        for idx, paragraph in enumerate(paragraphs, start=1):
            if getattr(paragraph, 'is_title_page', False) or getattr(paragraph, 'is_toc', False):
                continue

            text = paragraph.text
            stripped = text.strip()
            if not stripped:
                continue

            # Список литературы и приложения оформляются по ГОСТ — не проверяются
            if re.match(r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)', stripped.upper()):
                break

            if self.CODE_LINE_REGEX.match(text):
                in_code = True
                continue
            if in_code:
                if re.search(r'[а-яё]{3}', stripped.lower()) and stripped.endswith('.') and ' ' in stripped:
                    in_code = False
                else:
                    continue

            if 'oMath' in paragraph._element.xml:
                continue

            # 1. Длинное тире — должно быть среднее
            em_matches = list(self.EM_DASH_REGEX.finditer(text))
            if em_matches:
                pos = em_matches[0].start()
                context = text[max(0, pos - 25):pos + 25].strip()
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Абзац {idx} ('{stripped[:30]}...')",
                        message=f"В тексте применимо только тире средней длины (–), а не длинное (—): "
                                f"'{context}' ({len(em_matches)} шт.)."
                    )
                )

            # 2. Дефис вместо тире между словами
            hyphen_positions = []
            for m in re.finditer(r'(^|\n)?\s?-+\s', text):
                # Маркер списка в начале абзаца или строки — не тире в тексте
                if m.group(1) is not None:
                    continue
                hyphen_positions.append(m.start())

            if hyphen_positions:
                pos = hyphen_positions[0]
                context = text[max(0, pos - 25):pos + 25].strip()
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Абзац {idx} ('{stripped[:30]}...')",
                        message=f"Вместо дефиса между словами ставится тире средней длины (–): "
                                f"'{context}' ({len(hyphen_positions)} шт.)."
                    )
                )

        return violations
