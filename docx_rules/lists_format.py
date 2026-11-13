"""Проверка оформления списков."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import effective
from docx_rules.classification import classify_paragraphs


class ListsFormatRule(BaseRule):
    DASH_ITEM_REGEX = re.compile(r'^(–|-|—)\s*(.+)$')
    BRACKET_ITEM_REGEX = re.compile(r'^(\d+)\)\s*(.+)$')
    DOT_ITEM_REGEX = re.compile(r'^(\d+)\.\s+(.+)$')

    INDENT_TOLERANCE_CM = 0.15
    # Допустимые маркеры авто-списков (только тире)
    NUMBERING_DASHES = ('–', '-', '—')

    @staticmethod
    def _item_kind(text: str):
        """Тип пункта списка: 'dot', 'bracket', 'dash' или None."""
        if ListsFormatRule.DOT_ITEM_REGEX.match(text):
            return 'dot'
        if ListsFormatRule.BRACKET_ITEM_REGEX.match(text):
            return 'bracket'
        if ListsFormatRule.DASH_ITEM_REGEX.match(text):
            return 'dash'
        return None

    def _get_effective_left_indent(self, paragraph):
        indent = paragraph.paragraph_format.left_indent
        if indent is None and paragraph.style and paragraph.style.paragraph_format:
            indent = paragraph.style.paragraph_format.left_indent
        return indent

    def _get_effective_first_indent(self, paragraph):
        indent = paragraph.paragraph_format.first_line_indent
        if indent is None and paragraph.style and paragraph.style.paragraph_format:
            indent = paragraph.style.paragraph_format.first_line_indent
        return indent

    @staticmethod
    def _num_pr(paragraph):
        """(numId, ilvl) прямого w:numPr абзаца или None."""
        num_pr = paragraph._p.xpath('./w:pPr/w:numPr')
        if not num_pr:
            return None
        num_id = num_pr[0].xpath('./w:numId/@w:val')
        if not num_id:
            return None
        ilvl = num_pr[0].xpath('./w:ilvl/@w:val')
        return int(num_id[0]), int(ilvl[0]) if ilvl else 0

    @staticmethod
    def _numbering_level(doc, num_id: int, ilvl: int):
        """w:lvl уровня списка из numbering.xml (numId -> abstractNumId -> lvl) или None."""
        try:
            numbering = doc.part.numbering_part.element
        except Exception:
            return None
        for num in numbering.iter(qn('w:num')):
            if num.get(qn('w:numId')) == str(num_id):
                abstract = num.find(qn('w:abstractNumId'))
                if abstract is None:
                    return None
                abstract_id = abstract.get(qn('w:val'))
                for candidate in numbering.iter(qn('w:abstractNum')):
                    if candidate.get(qn('w:abstractNumId')) == abstract_id:
                        for lvl in candidate.iter(qn('w:lvl')):
                            if lvl.get(qn('w:ilvl')) == str(ilvl):
                                return lvl
                return None
        return None

    def _check_auto_item(self, doc, paragraphs, idx, num_id, ilvl, is_last, violations):
        """Проверка пункта авто-списка Word (номер/маркер генерируются автоматически)."""
        rule_name = "ListsFormatRule"
        paragraph = paragraphs[idx]
        text = paragraph.text.strip()
        loc = f"Абзац {idx + 1} ('{text[:30]}...')"

        lvl = self._numbering_level(doc, num_id, ilvl)
        if lvl is None:
            return  # numbering.xml или уровень недоступны — проверять нечего

        fmt_node = lvl.find(qn('w:numFmt'))
        text_node = lvl.find(qn('w:lvlText'))
        num_fmt = fmt_node.get(qn('w:val')) if fmt_node is not None else None
        lvl_text = text_node.get(qn('w:val')) if text_node is not None else None

        if num_fmt in (None, 'none'):
            return  # numFmt other/none — пропуск

        if num_fmt == 'bullet':
            if lvl_text not in self.NUMBERING_DASHES:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Маркер списка должен быть тире (–), а не '{lvl_text}'."
                    )
                )
        elif num_fmt == 'decimal':
            last_char = text[-1:]
            if lvl_text == '%1)':
                expected = '.' if is_last else ';'
                if last_char != expected:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message=f"В списке с нумерацией 'N)' пункт должен заканчиваться "
                                    f"{'точкой (последний)' if is_last else 'точкой с запятой'}."
                        )
                    )
            elif lvl_text == '%1.':
                if last_char != '.':
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message="В списке с нумерацией 'N.' пункт должен заканчиваться точкой."
                        )
                    )

        # Отступ: left = 0/отсутствует и красная строка 1.75 см
        left = effective(paragraph, 'left_indent')
        first = effective(paragraph, 'first_line_indent')
        if left is None and first is None:
            level_ppr = lvl.find(qn('w:pPr'))
            ind = level_ppr.find(qn('w:ind')) if level_ppr is not None else None
            if ind is not None:
                from docx.shared import Twips
                left_val = ind.get(qn('w:left')) or ind.get(qn('w:start'))
                hang_val = ind.get(qn('w:hanging'))
                first_val = ind.get(qn('w:firstLine'))
                if left_val is not None:
                    left = Twips(int(left_val))
                if hang_val is not None:
                    first = Twips(-int(hang_val))
                elif first_val is not None:
                    first = Twips(int(first_val))
        left_cm = left.cm if left is not None else 0.0
        first_cm = first.cm if first is not None else 0.0
        if abs(left_cm) > self.INDENT_TOLERANCE_CM or abs(first_cm - 1.75) > self.INDENT_TOLERANCE_CM:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.WARNING,
                    location=loc,
                    message=f"Отступ авто-списка: левый 0 см и красная строка 1.75 см "
                            f"(фактически: левый {round(left_cm, 2)} см, красная строка {round(first_cm, 2)} см)."
                )
            )

    def _check_item(self, paragraphs, idx, kind, is_last, violations):
        rule_name = "ListsFormatRule"
        paragraph = paragraphs[idx]
        text = paragraph.text.strip()
        loc = f"Абзац {idx + 1} ('{text[:30]}...')"

        if kind == 'dot':
            match = self.DOT_ITEM_REGEX.match(text)
            number = int(match.group(1))
            body = match.group(2)
            expected_indent = 2.0 if number >= 10 else 1.75
        elif kind == 'bracket':
            match = self.BRACKET_ITEM_REGEX.match(text)
            body = match.group(2)
            expected_indent = 1.75
        else:
            body = self.DASH_ITEM_REGEX.match(text).group(2)
            expected_indent = 1.75

        # 1. Регистр первой буквы после маркера
        first_char = body[:1]
        if first_char.isalpha():
            if kind in ('dash', 'bracket') and first_char.isupper():
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"После {'тире' if kind == 'dash' else 'скобки'} список пишется с маленькой буквы."
                    )
                )
            if kind == 'dot' and first_char.islower():
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="После точки список пишется с большой буквы."
                    )
                )

        # 2. Знак в конце пункта
        last_char = text[-1:]
        if kind == 'dot':
            if last_char not in ('.',):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="В нумерованном списке с точкой в конце пункта всегда ставится точка."
                    )
                )
        elif is_last:
            if last_char not in ('.', ':'):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="В последнем пункте списка ставится точка."
                    )
                )
        else:
            if last_char not in (';',):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Пункты {'маркированного' if kind == 'dash' else 'нумерованного'} списка "
                                f"заканчиваются точкой с запятой (кроме последнего)."
                    )
                )

        # 3. Отступ текста списка
        left = self._get_effective_left_indent(paragraph)
        first = self._get_effective_first_indent(paragraph)
        left_cm = left.cm if left is not None else None
        first_cm = first.cm if first is not None else None

        indent_values = [v for v in (left_cm, first_cm) if v is not None]
        if left_cm is not None and first_cm is not None:
            indent_values.append(left_cm + first_cm)
        if not any(abs(v - expected_indent) <= self.INDENT_TOLERANCE_CM for v in indent_values):
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.WARNING,
                    location=loc,
                    message=f"Отступ текста списка должен составлять {expected_indent} см "
                            f"(фактически: левый {round(left_cm, 2) if left_cm is not None else '—'} см, "
                            f"красная строка {round(first_cm, 2) if first_cm is not None else '—'} см)."
                )
            )

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "ListsFormatRule"

        paragraphs = classify_paragraphs(doc)

        def is_centered(p):
            align = p.alignment
            if align is None and p.style and p.style.paragraph_format:
                align = p.style.paragraph_format.alignment
            return align == WD_ALIGN_PARAGRAPH.CENTER

        item_indices = set()
        auto_map = {}
        for idx, p in enumerate(paragraphs):
            if getattr(p, 'is_title_page', False) or getattr(p, 'is_toc', False):
                continue
            text = p.text.strip()
            if not text:
                continue
            # Список литературы и приложения не проверяются как списки
            if re.match(r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)', text.upper()):
                break
            if is_centered(p):
                continue
            num_pr = self._num_pr(p)
            if num_pr is not None:
                auto_map[idx] = num_pr
            elif self._item_kind(text):
                item_indices.add(idx)

        # Авто-списки Word (номер генерируется через numPr/numbering.xml)
        for idx in sorted(auto_map):
            paragraph = paragraphs[idx]
            num_id, ilvl = auto_map[idx]
            if ilvl != 0:
                continue  # вложенные уровни не проверяются

            next_item_idx = idx + 1
            while next_item_idx < len(paragraphs) and not paragraphs[next_item_idx].text.strip():
                next_item_idx += 1
            is_last = not (next_item_idx in auto_map and auto_map[next_item_idx][0] == num_id
                           and auto_map[next_item_idx][1] == ilvl)

            self._check_auto_item(doc, paragraphs, idx, num_id, ilvl, is_last, violations)

        for idx in sorted(item_indices):
            paragraph = paragraphs[idx]
            text = paragraph.text.strip()
            kind = self._item_kind(text)
            loc = f"Абзац {idx + 1} ('{text[:30]}...')"

            # Соседний предыдущий/следующий пункт того же списка (пустые строки не считаются)
            prev_item_idx = idx - 1
            while prev_item_idx >= 0 and not paragraphs[prev_item_idx].text.strip():
                prev_item_idx -= 1
            next_item_idx = idx + 1
            while next_item_idx < len(paragraphs) and not paragraphs[next_item_idx].text.strip():
                next_item_idx += 1

            has_prev_same = prev_item_idx in item_indices and self._item_kind(paragraphs[prev_item_idx].text.strip()) == kind
            has_next_same = next_item_idx in item_indices and self._item_kind(paragraphs[next_item_idx].text.strip()) == kind
            is_last = not has_next_same

            self._check_item(paragraphs, idx, kind, is_last, violations)

            # 4. Двоеточие в абзаце перед первым пунктом списка (из двух и более пунктов)
            if not has_prev_same and has_next_same and prev_item_idx >= 0:
                prev_text = paragraphs[prev_item_idx].text.strip()
                if not prev_text.endswith(':'):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message="В абзаце перед списком ставится двоеточие."
                        )
                    )

        return violations
