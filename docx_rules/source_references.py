"""Проверка ссылок на источники в тексте."""

import re
import math

from docx import Document
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import is_toc_line
from docx_rules.classification import classify_paragraphs


class SourceReferencesRule(BaseRule):
    CHAPTER_REGEX = r'^ГЛАВА\s+\d+'
    STOP_SECTION_REGEX = r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|ПРИЛОЖЕНИЕ)'
    BIBLIOGRAPHY_HEADER_REGEX = r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ)|БИБЛИОГРАФИЯ)'
    LINK_REGEX = re.compile(r'\[([^\[\]]{1,30})\]')
    VALID_LINK_REGEX = r'^\d+(,\s*с\.\s*\d+)?$'
    # Признаки электронного ресурса в записи библиографии
    ELECTRONIC_REGEX = re.compile(r'(https?://|www\.|электронный\s+ресурс|режим\s+доступа|дата\s+обращения)',
                                  re.IGNORECASE)

    # Строки листингов кода не являются ссылками на источники
    CODE_LINE_REGEX = re.compile(
        r'^\s*('
        r'import\s|from\s+[A-Za-z_][\w.]*\s+import|'
        r'[A-Za-z_]\w*\s*(\[[^\]]*\])?\s*=[^=]|'
        r'print\s*\(|plt\.|def\s+[A-Za-z_]\w*\s*\(|return\b|'
        r'for\s+[A-Za-z_]\w*\s+in\s|while\s|elif\b|else:|#'
        r')'
    )

    CHARS_PER_LINE = 70
    LINES_PER_PAGE = 28

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "SourceReferencesRule"

        paragraphs = classify_paragraphs(doc)

        # Номера записей библиографии (для сверки со ссылками [N])
        bibliography = {}
        electronic = set()
        in_refs = False
        for paragraph in paragraphs:
            text = paragraph.text.strip()
            upper = text.upper()
            if not in_refs:
                if (re.match(self.BIBLIOGRAPHY_HEADER_REGEX, upper) and len(upper) < 100
                        and not is_toc_line(text)
                        and not getattr(paragraph, 'is_toc', False)
                        and not getattr(paragraph, 'is_title_page', False)):
                    in_refs = True
                continue
            if re.match(self.STOP_SECTION_REGEX, upper):
                break  # следующая секция после списка источников
            item_match = re.match(r'^(\d+)[\.\)]\s+', text)
            if item_match:
                num = int(item_match.group(1))
                bibliography[num] = text
                if self.ELECTRONIC_REGEX.search(text):
                    electronic.add(num)

        in_code = False
        for p_idx, paragraph in enumerate(paragraphs, start=1):
            if getattr(paragraph, 'is_title_page', False) or getattr(paragraph, 'is_toc', False):
                continue

            text = paragraph.text
            if not text.strip():
                continue

            if re.match(self.STOP_SECTION_REGEX, text.strip().upper()):
                break  # Список литературы и приложения не содержат ссылок-цитирований

            if self.CODE_LINE_REGEX.match(text):
                in_code = True
                continue

            if in_code:
                stripped = text.strip()
                # Конец листинга: абзац выглядит как обычный русский текст
                if re.search(r'[а-яё]{3}', stripped.lower()) and stripped.endswith('.') and ' ' in stripped:
                    in_code = False
                else:
                    continue

            for match in self.LINK_REGEX.finditer(text):
                content = match.group(1).strip()

                # Цитата состоит только из цифр, запятых, точек, пробелов и 'с.'
                if not re.search(r'\d', content) or not re.match(r'^[\d\s,с.]+$', content):
                    continue

                start, end = match.span()
                loc = f"Абзац {p_idx} ('{text.strip()[:30]}...')"

                if not re.match(self.VALID_LINK_REGEX, content):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message=f"Ссылка на источник '[{content}]' должна иметь формат [N] или [N, с.M] (например: [7, с.45])."
                        )
                    )
                    continue

                # Сверка с библиографией: источник должен присутствовать в списке
                num_match = re.match(self.VALID_LINK_REGEX, content)
                num = int(re.match(r'\d+', content).group(0))
                if bibliography and num not in bibliography:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message=f"Источник [{num}] отсутствует в списке источников (нумерованная запись не найдена)."
                        )
                    )
                elif num in electronic and num_match.group(1):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message=f"Источник [{num}] — электронный ресурс, указание страницы не имеет смысла: '[{content}]'."
                        )
                    )

                # Ссылка ставится в конце абзаца перед точкой
                tail = text[end:].strip()
                if tail not in ("", "."):
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message=f"Ссылка на источник '[{content}]' должна стоять в конце абзаца перед точкой."
                        )
                    )

                # Перед ссылкой должен быть пробел
                if start > 0 and text[start - 1].isalnum():
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message=f"Перед ссылкой '[{content}]' должен быть пробел."
                        )
                    )

        # В теоретическом разделе ссылки должны присутствовать на каждой странице
        chapter_paragraphs = []
        in_chapter = False
        for p_idx, paragraph in enumerate(paragraphs, start=1):
            text = paragraph.text.strip()
            if getattr(paragraph, 'is_title_page', False) or getattr(paragraph, 'is_toc', False):
                continue

            if re.match(self.CHAPTER_REGEX, text.upper()):
                if in_chapter:
                    break  # Конец первой главы (теоретического раздела)
                in_chapter = True
                estimated_lines = 0.0
                continue

            if in_chapter:
                if re.match(r'^(ГЛАВА\s+\d+|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)|ПРИЛОЖЕНИЕ)', text.upper()):
                    break
                chapter_paragraphs.append((p_idx, paragraph, estimated_lines))
                if text:
                    estimated_lines += math.ceil(len(text) / self.CHARS_PER_LINE)
                if '<w:drawing' in paragraph._element.xml or '<w:pict' in paragraph._element.xml:
                    estimated_lines += 10

        if chapter_paragraphs:
            pages_refs = {}
            for p_idx, paragraph, est_lines in chapter_paragraphs:
                page_no = int(est_lines // self.LINES_PER_PAGE)
                pages_refs.setdefault(page_no, 0)
                if re.search(r'\[\d+', paragraph.text):
                    pages_refs[page_no] += 1

            for page_no in sorted(pages_refs):
                if pages_refs[page_no] == 0:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=f"Теоретический раздел (Глава 1), ~страница {page_no + 1} от начала главы",
                            message="В теоретическом разделе ссылки на источники должны присутствовать на каждой странице."
                        )
                    )

        return violations
