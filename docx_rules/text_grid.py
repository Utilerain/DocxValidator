"""Проверка активной сетки текста (w:docGrid) в секциях документа."""

import re

from docx import Document
from docx.oxml.ns import qn
from typing import List

from docx_rules.base import BaseRule, Severity, Violation


class TextGridRule(BaseRule):
    """Проверяет активную сетку, не навязывая точный шаг."""
    def check(self, doc: Document) -> List[Violation]:
        violations = []
        for index, section in enumerate(doc.sections, 1):
            grids = section._sectPr.findall(qn('w:docGrid'))
            errors = []
            if len(grids) > 1:
                errors.append('В секции задано несколько противоречивых w:docGrid.')
            if not grids:
                errors.append('Отсутствует сетка текста (w:docGrid).')
            for grid in grids:
                kind = grid.get(qn('w:type'), 'none')
                if kind not in ('lines', 'linesAndChars', 'snapToChars'):
                    errors.append(f'Сетка текста не активна или имеет неизвестный type: {kind}.')
                if kind in ('lines', 'linesAndChars'):
                    pitch = grid.get(qn('w:linePitch'), '')
                    if not re.fullmatch(r'[+]?[0-9]+', pitch) or int(pitch) <= 0:
                        errors.append('linePitch сетки строк должен быть положительным целым.')
                space = grid.get(qn('w:charSpace'))
                if space is not None and not re.fullmatch(r'[+-]?[0-9]+', space):
                    errors.append('charSpace должен быть целым числом (включая отрицательные и ноль).')
            violations.extend(Violation('TextGridRule', Severity.ERROR, f'Секция {index}', e) for e in errors)
        return violations
