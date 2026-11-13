"""Проверка вывода в конце параграфов и глав."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective
from docx_rules.classification import classify_paragraphs


class ParagraphConclusionRule(BaseRule):
    MAJOR_HEADER_REGEX = r'^(ГЛАВА\s+\d+|СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)'
    SUBHEADING_REGEX = r'^\d+(\.\d+)+\.?\s+'
    NUMBERED_SECTION_REGEX = r'^\d+\.?\s+'
    STOP_SECTION_REGEX = r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)'
    SKIP_SECTION_REGEX = r'^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СОДЕРЖАНИЕ)'

    # Начало вывода по итогу параграфа / главы
    CONCLUSION_START_REGEX = (
        r'^(ТАКИМ ОБРАЗОМ|ИТАК\b|ПОДВОДЯ ИТОГ|ОБОБЩАЯ|РЕЗЮМИРУЯ|СЛЕДОВАТЕЛЬНО|'
        r'ПРОАНАЛИЗИРОВАВ|ВЫВОД[Ы]?\b|ВСЁ ВЫШЕСКАЗАННОЕ|ВЫШЕИЗЛОЖЕННОЕ)'
    )

    @staticmethod
    def iter_body_sequence(paragraphs: List, tables: List):
        """
        Последовательность элементов тела документа в порядке следования:
        ('p', абзац, индекс_абзаца) и ('tbl', таблица, None).
        Нужна, чтобы видеть таблицы между абзацами.
        """
        items = []
        p_idx = 0
        para_iter = iter(paragraphs)
        tbl_iter = iter(tables)
        body = None
        if paragraphs:
            body = paragraphs[0]._element.getparent()
        if body is None:
            return items
        for child in body.iterchildren():
            if child.tag == qn('w:p'):
                para = next(para_iter, None)
                if para is not None:
                    items.append(('p', para, p_idx))
                    p_idx += 1
            elif child.tag == qn('w:tbl'):
                table = next(tbl_iter, None)
                if table is not None:
                    items.append(('tbl', table, None))
        return items

    def _is_centered(self, paragraph) -> bool:
        align = effective(paragraph, 'alignment')
        return align == WD_ALIGN_PARAGRAPH.CENTER

    def _header_type(self, paragraph) -> Optional[str]:
        text = paragraph.text.strip()
        if not text:
            return None
        text_upper = text.upper()
        if re.match(self.MAJOR_HEADER_REGEX, text_upper):
            return 'major'
        if re.match(self.SUBHEADING_REGEX, text):
            return 'sub'
        # Раздел без слова 'Глава' (например '1. Понятие ...'), оформленный по центру
        if re.match(self.NUMBERED_SECTION_REGEX, text) and self._is_centered(paragraph) and len(text) < 120:
            return 'major'
        return None

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "ParagraphConclusionRule"

        paragraphs = classify_paragraphs(doc)
        items = self.iter_body_sequence(paragraphs, doc.tables)

        starts = []
        for i, (kind, obj, p_idx) in enumerate(items):
            if kind != 'p':
                continue
            if getattr(obj, 'is_title_page', False) or getattr(obj, 'is_toc', False):
                continue
            header_type = self._header_type(obj)
            if header_type:
                starts.append((i, header_type, obj, p_idx))

        for pos, (start_idx, header_type, header_para, header_idx) in enumerate(starts):
            text_upper = header_para.text.strip().upper()

            if re.match(self.STOP_SECTION_REGEX, text_upper):
                break  # После списка источников и приложения выводы не проверяются
            if header_type != 'major' and header_type != 'sub':
                continue
            if re.match(self.SKIP_SECTION_REGEX, text_upper):
                continue  # Введение, заключение и содержание не являются параграфами/главами

            end_idx = starts[pos + 1][0] if pos + 1 < len(starts) else len(items)
            block = items[start_idx + 1:end_idx]

            last_para = None
            last_para_idx = None
            last_item_kind = None
            for kind, obj, p_idx in block:
                if kind == 'p':
                    if obj.text.strip():
                        last_para = obj
                        last_para_idx = p_idx
                        last_item_kind = 'p'
                else:
                    last_item_kind = 'tbl'

            if last_para is None:
                continue  # Блок без содержания (например, заголовок главы перед параграфом)

            header_text = header_para.text.strip()[:40]
            loc = f"Конец '{header_text}' (абзац {last_para_idx + 1})"

            # 1. Таблица в конце раздела
            if last_item_kind == 'tbl':
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="В конце параграфа не допустима таблица (она должна быть внутри текста раздела)."
                    )
                )

            # 2. Вывод по итогу раздела
            if not re.match(self.CONCLUSION_START_REGEX, last_para.text.strip().upper()):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Параграф, глава всегда должны заканчиваться выводом (например: 'Таким образом, ...')."
                    )
                )

            # 3. Недопустимые элементы в конце раздела
            last_xml = last_para._element.xml
            if '<w:drawing' in last_xml or '<w:pict' in last_xml:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="В конце параграфа не допустимы рисунки."
                    )
                )

            if 'oMath' in last_xml:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="В конце параграфа не допустимы формулы."
                    )
                )

            if re.search(r'\[\d+', last_para.text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="В конце параграфа не допустимы ссылки на источники."
                    )
                )

        return violations
