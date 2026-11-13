"""Проверка белой заливки фона текста."""

import re

from docx import Document
from docx.oxml.ns import qn
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation


class WhiteFillRule(BaseRule):
    """
    Белая заливка фона текста: на белом листе её не видно, но при включении
    отображения сетки/границ текст оказывается на белой подложке.
    Проверяется заливка (w:shd fill="FFFFFF") и белое выделение (w:highlight white)
    на уровне абзацев, фрагментов и ячеек таблиц.
    """
    WHITE_FILLS = {"ffffff", "white"}

    def _is_white_shd(self, shd) -> bool:
        if shd is None:
            return False
        # w:val="nil" — заливки нет вовсе
        if (shd.get(qn('w:val')) or '').lower() == 'nil':
            return False
        fill = (shd.get(qn('w:fill')) or '').strip().lstrip('#').lower()
        return fill in self.WHITE_FILLS

    def _check_shading_in_paragraph(self, paragraph) -> Optional[str]:
        """Возвращает описание белой заливки в абзаце или None."""
        problems = []

        pPr = paragraph._p.pPr
        if pPr is not None and self._is_white_shd(pPr.find(qn('w:shd'))):
            problems.append("заливка фона всего абзаца")

        for run in paragraph.runs:
            if not run.text.strip():
                continue
            rPr = run._r.rPr
            if rPr is None:
                continue
            if self._is_white_shd(rPr.find(qn('w:shd'))):
                problems.append(f"заливка фона фрагмента '{run.text.strip()[:20]}'")
            highlight = rPr.find(qn('w:highlight'))
            if highlight is not None and (highlight.get(qn('w:val')) or '').lower() in self.WHITE_FILLS:
                problems.append(f"белое выделение фрагмента '{run.text.strip()[:20]}'")

        return "; ".join(problems) if problems else None

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "Белая заливка текста"

        # 1. Обычные абзацы документа
        for idx, paragraph in enumerate(doc.paragraphs, start=1):
            text = paragraph.text.strip()
            description = self._check_shading_in_paragraph(paragraph)
            if description:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Абзац {idx} ('{text[:30]}...')" if text else f"Абзац {idx}",
                        message=f"Обнаружена белая заливка фона ({description}): при отображении "
                                f"сетки текст оказывается на белой подложке. Заливку необходимо убрать."
                    )
                )

        # 2. Ячейки таблиц (абзацы внутри ячеек в doc.paragraphs не попадают)
        for tbl_idx, table in enumerate(doc.tables, start=1):
            affected_cells = 0
            for row in table.rows:
                for cell in row.cells:
                    tcPr = cell._tc.tcPr
                    has_cell_fill = tcPr is not None and self._is_white_shd(tcPr.find(qn('w:shd')))
                    has_text_fill = any(
                        self._check_shading_in_paragraph(p) for p in cell.paragraphs
                    )
                    if has_cell_fill or has_text_fill:
                        affected_cells += 1
            if affected_cells:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Таблица №{tbl_idx} по порядку в документе",
                        message=f"Обнаружена белая заливка фона в {affected_cells} яч.: при отображении "
                                f"сетки текст оказывается на белой подложке. Заливку необходимо убрать."
                    )
                )

        return violations
