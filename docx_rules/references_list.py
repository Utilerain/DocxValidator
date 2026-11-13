"""Проверка списка источников/литературы."""

import re

from docx import Document
from typing import List

from docx_rules.base import BaseRule, Severity, Violation


class ReferencesListRule(BaseRule):
    HEADER_REGEX = r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ))'
    NEXT_SECTION_REGEX = r'^(ПРИЛОЖЕНИЕ|ГЛАВА\s+\d+|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СОДЕРЖАНИЕ)'
    URL_REGEX = r'(https?://[^\s]+|www\.[^\s]+|\[?электронный\s+ресурс\]?)'
    ACCESS_DATE_REGEX = r'дата\s+обращения:?\s*\d{2}\.\d{2}\.\d{4}'
    PERIODICAL_KEYWORDS = ['ЖУРНАЛ', 'ВЕСТНИК', 'ИЗВЕСТИЯ', 'ПЕРИОДИЧЕСКОЕ ИЗДАНИЕ', 'ВЫП.']

    def _get_sort_key(self, text: str) -> str:
        cleaned = re.sub(r'^\d+[\.\)]\s*', '', text).strip()
        cleaned = re.sub(r'^\s*\[\d+\]\s*', '', cleaned)
        return cleaned.lower()

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "ReferencesListRule"

        ref_header_idx = None
        for idx, p in enumerate(doc.paragraphs):
            text = p.text.strip().upper()
            if re.match(self.HEADER_REGEX, text) and len(text) < 100:
                ref_header_idx = idx

        if ref_header_idx is None:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Документ",
                    message="Раздел 'Список источников и литературы' не найден в документе."
                )
            )
            return violations

        items = []
        for p_idx in range(ref_header_idx + 1, len(doc.paragraphs)):
            p = doc.paragraphs[p_idx]
            text = p.text.strip()
            if not text:
                continue

            text_upper = text.upper()
            if re.match(self.NEXT_SECTION_REGEX, text_upper) and p_idx > ref_header_idx + 1:
                break

            items.append((p_idx + 1, text))

        if not items:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location=f"Абзац {ref_header_idx + 1}",
                    message="Раздел 'Список источников и литературы' пуст."
                )
            )
            return violations

        printed_items = []
        electronic_items = []
        seen_electronic = False

        for p_num, text in items:
            loc = f"Список источников, абзац {p_num} ('{text[:30]}...')"

            if not re.match(r'^\d+[\.\)]\s*', text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Источник должен быть частью нумерованного списка (начинаться с '1. ')."
                    )
                )

            clean_text = re.sub(r'^\d+[\.\)]\s*', '', text).strip()
            is_electronic = bool(re.search(self.URL_REGEX, text, re.IGNORECASE))

            if is_electronic:
                seen_electronic = True
                electronic_items.append((p_num, text))

                if not re.search(self.ACCESS_DATE_REGEX, text, re.IGNORECASE):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Для электронного ресурса обязательно должна быть указана дата обращения (например: 'дата обращения: 06.06.2025')."
                        )
                    )
            else:
                if seen_electronic:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Печатное издание находится после электронного ресурса. Электронные ресурсы прописываются строго ПОСЛЕ печатных изданий."
                        )
                    )
                printed_items.append((p_num, text))

            text_upper = text.upper()
            if any(kw in text_upper for kw in self.PERIODICAL_KEYWORDS) and '//' in text:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Обнаружены признаки периодического издания (статья/журнал). Использование периодических изданий не допускается."
                    )
                )

            initials_first_match = re.match(r'^([А-ЯA-Z]\.\s*[А-ЯA-Z]\.|[А-ЯA-Z]\.)\s+([А-ЯA-Z][а-яa-z]+)', clean_text)
            if initials_first_match:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Нарушен порядок записи автора ('{initials_first_match.group(0)}'). Инициалы должны идти ПОСЛЕ фамилии: Фамилия И.О. (например, 'Достоевский Ф.М.')."
                    )
                )

        for i in range(1, len(printed_items)):
            prev_pnum, prev_text = printed_items[i - 1]
            curr_pnum, curr_text = printed_items[i]

            if self._get_sort_key(curr_text) < self._get_sort_key(prev_text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Список источников, элемент {curr_pnum}",
                        message=f"Нарушен алфавитный порядок печатных изданий! Элемент '{curr_text[:25]}...' должен располагаться раньше '{prev_text[:25]}...'."
                    )
                )

        for i in range(1, len(electronic_items)):
            prev_pnum, prev_text = electronic_items[i - 1]
            curr_pnum, curr_text = electronic_items[i]

            if self._get_sort_key(curr_text) < self._get_sort_key(prev_text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Список источников, элемент {curr_pnum}",
                        message=f"Нарушен алфавитный порядок электронных ресурсов! Элемент '{curr_text[:25]}...' должен располагаться раньше '{prev_text[:25]}...'."
                    )
                )

        return violations
