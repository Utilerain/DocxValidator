"""Проверка титульной страницы."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation


class TitlePageRule(BaseRule):
    def _get_title_page_paragraphs(self, doc: Document) -> List:
        title_paragraphs = []
        for p in doc.paragraphs:
            text_upper = p.text.strip().upper()

            has_page_break = any(
                'w:br' in run._element.xml and 'w:type="page"' in run._element.xml
                for run in p.runs
            )

            if text_upper in ["СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ОГЛАВЛЕНИЕ"] or has_page_break:
                break

            title_paragraphs.append(p)

            if re.search(r'\b20\d{2}\b', text_upper) and len(text_upper) < 30:
                break

        return title_paragraphs

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "TitlePageRule"

        title_paragraphs = self._get_title_page_paragraphs(doc)
        total_p = len(title_paragraphs)

        if total_p < 5:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Титульный лист",
                    message="Титульный лист не найден или содержит слишком мало текста."
                )
            )
            return violations

        topic_idx: Optional[int] = None
        topic_paragraph = None
        has_institution_header = False
        has_work_type = False
        has_year = False

        for idx, p in enumerate(title_paragraphs):
            text = p.text.strip()
            text_upper = text.upper()

            if not text:
                continue

            align = p.alignment
            if align is None and p.style and p.style.paragraph_format:
                align = p.style.paragraph_format.alignment

            if any(word in text_upper for word in ["МИНИСТЕРСТВО", "ОБРАЗОВАТЕЛЬНОЕ", "КОЛЛЕДЖ", "УНИВЕРСИТЕТ", "ИНСТИТУТ"]):
                has_institution_header = True
                if align != WD_ALIGN_PARAGRAPH.CENTER:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Титульный лист, абзац {idx+1} ('{text[:30]}...')",
                            message="Шапка учебного заведения должна быть выровнена по центру."
                        )
                    )

            if "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА" in text_upper or "ДИПЛОМНЫЙ ПРОЕКТ" in text_upper:
                has_work_type = True

            is_bold = any(run.bold for run in p.runs if run.text.strip())
            is_topic_candidate = (
                ("РАЗРАБОТКА" in text_upper or "ТЕМА" in text_upper or "ПРОЕКТИРОВАНИЕ" in text_upper or is_bold)
                and not has_work_type
                and idx > 2
            )

            if is_topic_candidate and topic_paragraph is None and "МИНИСТЕРСТВО" not in text_upper and "КОЛЛЕДЖ" not in text_upper:
                topic_idx = idx
                topic_paragraph = p

            if re.search(r'\b20\d{2}\b', text_upper):
                has_year = True
                if align != WD_ALIGN_PARAGRAPH.CENTER:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Титульный лист, абзац {idx+1} ('{text}')",
                            message="Город и год внизу титульного листа должны быть выровнены по центру."
                        )
                    )

            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.cm > 0.1:
                if align == WD_ALIGN_PARAGRAPH.CENTER or "МИНИСТЕРСТВО" in text_upper or "РАЗРАБОТКА" in text_upper:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=f"Титульный лист, абзац {idx+1} ('{text[:30]}...')",
                            message="У центрального блока титульного листа установлена красная строка. Значение должно быть 0."
                        )
                    )

        if topic_paragraph is not None and topic_idx is not None:
            topic_align = topic_paragraph.alignment
            if topic_align is None and topic_paragraph.style and topic_paragraph.style.paragraph_format:
                topic_align = topic_paragraph.style.paragraph_format.alignment

            if topic_align != WD_ALIGN_PARAGRAPH.CENTER:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Титульный лист, абзац {topic_idx+1} ('{topic_paragraph.text[:30]}...')",
                        message="Тема дипломного проекта должна быть выровнена по центру."
                    )
                )

            topic_has_bold = any(run.bold for run in topic_paragraph.runs if run.text.strip())
            if not topic_has_bold:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Титульный лист, абзац {topic_idx+1}",
                        message="Тема дипломного проекта должна быть выделена полужирным шрифтом."
                    )
                )

            relative_position = topic_idx / total_p
            if relative_position < 0.20 or relative_position > 0.75:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Титульный лист, абзац {topic_idx+1}",
                        message=f"Тема дипломного проекта расположена слишком {'высоко' if relative_position < 0.20 else 'низко'}. Она должна быть примерно в центре титульного листа."
                    )
                )
        else:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Титульный лист",
                    message="Тема дипломного проекта не найдена на титульном листе."
                )
            )

        if not has_institution_header:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.WARNING,
                    location="Титульный лист",
                    message="Не найдена шапка учебного заведения / министерства."
                )
            )

        if not has_year:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.WARNING,
                    location="Титульный лист",
                    message="Внизу титульного листа не найден город и год."
                )
            )

        return violations
