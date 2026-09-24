from dataclasses import dataclass
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from enum import Enum
from typing import List, Optional
from tqdm import tqdm
import re
import math

class Severity(Enum):
    ERROR = "ERR"
    WARNING = "WARN"

@dataclass
class Violation:
    rule_name: str
    severity: Severity
    location: str
    message: str

class BaseRule:
    def check(self, doc: Document) -> List[Violation]:
        raise NotImplementedError

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

class MainTextFormatRule(BaseRule):
    def __init__(
        self,
        target_font_name: str = "Times New Roman",
        target_font_size: float = 14.0,
        target_line_spacing: float = 1.5,
        target_first_line_indent_cm: float = 1.25,
        target_space_before_pt: float = 0.0,
        target_space_after_pt: float = 0.0,
        allow_formatting_styles: bool = False,  # Разрешать ли bold/italic
        indent_tolerance_cm: float = 0.05       # Допустимая погрешность отступа
    ):
        self.font_name = target_font_name
        self.font_size = target_font_size
        self.line_spacing = target_line_spacing
        self.first_line_indent_cm = target_first_line_indent_cm
        self.space_before_pt = target_space_before_pt
        self.space_after_pt = target_space_after_pt
        self.allow_formatting_styles = allow_formatting_styles
        self.indent_tolerance_cm = indent_tolerance_cm

    def _get_effective_font_name(self, run, paragraph) -> Optional[str]:
        if run.font.name:
            return run.font.name
        if paragraph.style and paragraph.style.font and paragraph.style.font.name:
            return paragraph.style.font.name
        return None

    def _get_effective_font_size(self, run, paragraph) -> Optional[float]:
        if run.font.size:
            return run.font.size.pt
        if paragraph.style and paragraph.style.font and paragraph.style.font.size:
            return paragraph.style.font.size.pt
        return None

    def _get_effective_alignment(self, paragraph):
        if paragraph.alignment is not None:
            return paragraph.alignment
        if paragraph.style and paragraph.style.paragraph_format.alignment is not None:
            return paragraph.style.paragraph_format.alignment
        return None

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "MainTextFormatRule"

        for p_idx, paragraph in enumerate(doc.paragraphs, start=1):
            text = paragraph.text.strip()
            # Пропускаем пустые абзацы
            if not text:
                continue

            loc = f"Абзац {p_idx} ('{text[:30]}...')"
            fmt = paragraph.paragraph_format

            # 1. Выравнивание (по ширине / JUSTIFY)
            alignment = self._get_effective_alignment(paragraph)
            if alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Выравнивание текста должно быть по ширине (JUSTIFY)."
                    )
                )

            # 2. Красная строка (1,25 см)
            indent = fmt.first_line_indent
            if indent is None and paragraph.style and paragraph.style.paragraph_format.first_line_indent:
                indent = paragraph.style.paragraph_format.first_line_indent

            if indent is None:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Не задана красная строка (ожидается {self.first_line_indent_cm} см)."
                    )
                )
            else:
                indent_cm = indent.cm
                if abs(indent_cm - self.first_line_indent_cm) > self.indent_tolerance_cm:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message=f"Отступ красной строки равен {round(indent_cm, 2)} см, а должен быть {self.first_line_indent_cm} см."
                        )
                    )

            # 3. Междустрочный интервал (1.5)
            line_sp = fmt.line_spacing
            if line_sp is None and paragraph.style and paragraph.style.paragraph_format.line_spacing:
                line_sp = paragraph.style.paragraph_format.line_spacing

            if line_sp is not None and abs(float(line_sp) - self.line_spacing) > 0.05:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Междустрочный интервал {line_sp}, ожидается {self.line_spacing}."
                    )
                )

            # 4. Отступы до и после абзаца (0 pt)
            space_before = fmt.space_before.pt if fmt.space_before else 0.0
            space_after = fmt.space_after.pt if fmt.space_after else 0.0

            if abs(space_before - self.space_before_pt) > 0.1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Отступ перед абзацем равен {space_before} pt (ожидается {self.space_before_pt} pt)."
                    )
                )

            if abs(space_after - self.space_after_pt) > 0.1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Отступ после абзаца равен {space_after} pt (ожидается {self.space_after_pt} pt)."
                    )
                )

            # 5. Запрет висячих строк
            if fmt.widow_control is False:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Отключен запрет висячих строк (widow_control)."
                    )
                )

            # 6. Проверка свойств текста (Runs): шрифт, кегль, начертания
            for r_idx, run in enumerate(paragraph.runs, start=1):
                run_text = run.text.strip()
                if not run_text:
                    continue

                run_loc = f"{loc}, фрагмент {r_idx} ('{run_text[:15]}')"

                # Проверка шрифта
                font_name = self._get_effective_font_name(run, paragraph)
                if font_name and font_name != self.font_name:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=run_loc,
                            message=f"Шрифт '{font_name}', а должен быть '{self.font_name}'."
                        )
                    )

                # Проверка размера шрифта (14 pt)
                font_size = self._get_effective_font_size(run, paragraph)
                if font_size and abs(font_size - self.font_size) > 0.1:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=run_loc,
                            message=f"Размер шрифта {font_size} pt, а должен быть {self.font_size} pt."
                        )
                    )

                # Виды начертаний в тексте не применимы
                if not self.allow_formatting_styles:
                    if run.bold:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=run_loc,
                                message="В тексте использовано полужирное начертание (bold)."
                            )
                        )
                    if run.italic:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=run_loc,
                                message="В тексте использован курсив (italic)."
                            )
                        )
                    if run.underline:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=run_loc,
                                message="В тексте использовано подчеркивание (underline)."
                            )
                        )

        return violations

class PageNumberingRule(BaseRule):
    """
    Правило проверки нумерации страниц:
    1. Расположение: правый верхний угол
    2. Шрифт: Times New Roman, 10 pt
    3. Отсутствие лишних пустых строк в колонтитуле
    4. Отсутствие нумерации на титульном листе
    """

    def __init__(
        self,
        target_font_name: str = "Times New Roman",
        target_font_size: float = 10.0,
        target_alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ):
        self.font_name = target_font_name
        self.font_size = target_font_size
        self.alignment = target_alignment

    def _has_page_number_field(self, header_or_paragraph) -> bool:
        """
        Проверяет наличие XML-поля динамической нумерации страниц (PAGE) в OpenXML.
        """
        xml_str = header_or_paragraph._element.xml
        return ("instrText" in xml_str and "PAGE" in xml_str) or ("w:fldSimple" in xml_str and "PAGE" in xml_str)

    def _get_effective_font_name(self, run, paragraph) -> Optional[str]:
        if run.font.name:
            return run.font.name
        if paragraph.style and paragraph.style.font and paragraph.style.font.name:
            return paragraph.style.font.name
        return None

    def _get_effective_font_size(self, run, paragraph) -> Optional[float]:
        if run.font.size:
            return run.font.size.pt
        if paragraph.style and paragraph.style.font and paragraph.style.font.size:
            return paragraph.style.font.size.pt
        return None

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "PageNumberingRule"

        if not doc.sections:
            return violations

        # 1. Проверка первой страницы (титульный лист)
        first_section = doc.sections[0]

        if first_section.different_first_page_header_footer:
            # Если особый колонтитул включен, проверяем, что в нем НЕТ нумерации
            if self._has_page_number_field(first_section.first_page_header):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location="Секция 1 (Колонтитул первой страницы)",
                        message="На титульном листе присутствует нумерация страниц."
                    )
                )
        else:
            # Если опция не включена, проверяем основной колонтитул первой секции
            if self._has_page_number_field(first_section.header):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location="Секция 1 (Верхний колонтитул)",
                        message="Не включена опция 'Особый колонтитул для первой страницы' (different_first_page_header_footer). На титульном листе не должно быть нумерации."
                    )
                )

        # 2. Проверка колонтитулов во всех секциях
        for sec_idx, section in enumerate(doc.sections, start=1):
            header = section.header

            # Если в колонтитуле нет поля нумерации, пропускаем проверки оформления
            if not self._has_page_number_field(header):
                continue

            paragraphs = header.paragraphs

            # 3. Проверка на лишние строки в колонтитуле
            if len(paragraphs) > 1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Секция {sec_idx} (Верхний колонтитул)",
                        message=f"В верхнем колонтитуле с нумерацией обнаружены лишние пустые строки (всего строк: {len(paragraphs)}, должна быть 1)."
                    )
                )

            for p_idx, paragraph in enumerate(paragraphs, start=1):
                # Проверка наличия лишних символов перевода строки \n внутри текста
                if "\n" in paragraph.text:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}",
                            message="В строке нумерации присутствуют лишние переносы строк."
                        )
                    )

                # 4. Проверка выравнивания (правый верхний угол)
                alignment = paragraph.alignment
                if alignment is None and paragraph.style and paragraph.style.paragraph_format:
                    alignment = paragraph.style.paragraph_format.alignment

                if alignment != self.alignment:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}",
                            message="Нумерация страниц должна быть выровнена по правому краю (правый верхний угол)."
                        )
                    )

                # 5. Проверка шрифта и размера шрифта в фрагментах (Runs)
                for r_idx, run in enumerate(paragraph.runs, start=1):
                    font_name = self._get_effective_font_name(run, paragraph)
                    if font_name and font_name != self.font_name:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}, фрагмент {r_idx}",
                                message=f"Шрифт нумерации '{font_name}', а должен быть '{self.font_name}'."
                            )
                        )

                    font_size = self._get_effective_font_size(run, paragraph)
                    if font_size and abs(font_size - self.font_size) > 0.1:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.ERROR,
                                location=f"Секция {sec_idx}, колонтитул, абзац {p_idx}, фрагмент {r_idx}",
                                message=f"Размер шрифта нумерации {font_size} pt, а должен быть {self.font_size} pt."
                            )
                        )

        return violations

class TitlePageRule(BaseRule):
    """
    Правило проверки соответствия оформления титульного листа:
    1. Расположение темы дипломного проекта примерно в центре страницы.
    2. Выравнивание темы и шапки по центру.
    3. Полужирное начертание темы.
    4. Выравнивание наименования города и года внизу страницы по центру.
    5. Отсутствие красной строки у центральных блоков титульного листа.
    """

    def _get_title_page_paragraphs(self, doc: Document) -> List:
        """Извлекает абзацы только титульного листа."""
        title_paragraphs = []
        for p in doc.paragraphs:
            text_upper = p.text.strip().upper()

            # Проверка разрыва страницы внутри элементов (runs)
            has_page_break = any(
                'w:br' in run._element.xml and 'w:type="page"' in run._element.xml
                for run in p.runs
            )

            # Остановка при переходе к Содержанию/Введению или явном разрыве страницы
            if text_upper in ["СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ОГЛАВЛЕНИЕ"] or has_page_break:
                break

            title_paragraphs.append(p)

            # Остановка, если найден город и год внизу (например: ТУЙМАЗЫ 2026)
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

            # 1. Проверка шапки (Министерство / Колледж / Учреждение)
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

            # 2. Поиск типа работы
            if "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА" in text_upper or "ДИПЛОМНЫЙ ПРОЕКТ" in text_upper:
                has_work_type = True

            # 3. Поиск темы дипломного проекта
            is_bold = any(run.bold for run in p.runs if run.text.strip())
            is_topic_candidate = (
                ("РАЗРАБОТКА" in text_upper or "ТЕМА" in text_upper or "ПРОЕКТИРОВАНИЕ" in text_upper or is_bold)
                and not has_work_type
                and idx > 2
            )

            if is_topic_candidate and topic_paragraph is None and "МИНИСТЕРСТВО" not in text_upper and "КОЛЛЕДЖ" not in text_upper:
                topic_idx = idx
                topic_paragraph = p

            # 4. Проверка наименования города и года внизу страницы (например, ТУЙМАЗЫ 2026)
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

            # 5. Проверка отсутствия красной строки для заголовочных и центральных блоков
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

        # 6. Проверка правильности темы дипломного проекта
        if topic_paragraph is not None and topic_idx is not None:
            topic_align = topic_paragraph.alignment
            if topic_align is None and topic_paragraph.style and topic_paragraph.style.paragraph_format:
                topic_align = topic_paragraph.style.paragraph_format.alignment

            # Проверка выравнивания по центру
            if topic_align != WD_ALIGN_PARAGRAPH.CENTER:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=f"Титульный лист, абзац {topic_idx+1} ('{topic_paragraph.text[:30]}...')",
                        message="Тема дипломного проекта должна быть выровнена по центру."
                    )
                )

            # Проверка выделения полужирным шрифтом
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

            # Проверка расположения примерно в центре страницы (от 20% до 75% высоты титульника)
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

class TableOfContentsRule(BaseRule):
    """
    Правило проверки оглавления (содержания):
    1. Название ("Содержание"): выравнивание по центру, полужирный шрифт, наличие красной строки.
    2. Пункты оглавления: выравнивание по ширине, без красной строки.
    3. Запрет ручного проставления точек (набор нескольких точек подряд клавишей).
    4. Строго один пробел между номером (главы/параграфа) и названием.
    """

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "TableOfContentsRule"

        toc_header_index: Optional[int] = None

        # 1. Поиск заголовка "Содержание" или "Оглавление"
        for idx, p in enumerate(doc.paragraphs):
            clean_text = p.text.strip().upper()
            if clean_text in ["СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"]:
                toc_header_index = idx
                loc = f"Заголовок оглавления (абзац {idx+1})"

                # А) Выравнивание заголовка по центру
                align = p.alignment
                if align is None and p.style and p.style.paragraph_format:
                    align = p.style.paragraph_format.alignment
                if align != WD_ALIGN_PARAGRAPH.CENTER:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Заголовок 'Содержание' должен быть выровнен по центру."
                        )
                    )

                # Б) Выделение полужирным начертанием
                is_bold = any(run.bold for run in p.runs if run.text.strip())
                if not is_bold and p.style and p.style.font:
                    is_bold = p.style.font.bold or False
                if not is_bold:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Название 'Содержание' должно быть выделено полужирным шрифтом."
                        )
                    )

                # В) Наличие красной строки у заголовка
                fmt = p.paragraph_format
                first_indent = fmt.first_line_indent
                if first_indent is None and p.style and p.style.paragraph_format:
                    first_indent = p.style.paragraph_format.first_line_indent
                if first_indent is None or first_indent.cm <= 0.05:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message="У заголовка 'Содержание' должна быть установлена красная строка."
                        )
                    )
                break

        if toc_header_index is None:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Документ",
                    message="Раздел 'Содержание' (или 'Оглавление') не найден в документе."
                )
            )
            return violations

        # 2. Поиск и сбор всех строк оглавления
        toc_entries = []
        for p_idx in range(toc_header_index + 1, len(doc.paragraphs)):
            p = doc.paragraphs[p_idx]
            text = p.text.strip()
            if not text:
                continue

            # Признаки окончания раздела Оглавления:
            # Появление основного текста Введения/Главы без номерной табуляции и поля TOC
            is_toc_field = 'TOC' in p._element.xml or 'w:fldSimple' in p._element.xml or 'w:instrText' in p._element.xml
            has_page_num_at_end = re.search(r'\d+\s*$', text) is not None

            if not has_page_num_at_end and not is_toc_field and text.upper() in ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "ПРИЛОЖЕНИЕ"]:
                break

            # Разрыв страницы
            has_page_break = any(
                'w:br' in run._element.xml and 'w:type="page"' in run._element.xml
                for run in p.runs
            )

            toc_entries.append((p_idx + 1, p))

            if has_page_break:
                break

        # 3. Проверка пунктов оглавления
        for p_num, p in toc_entries:
            text = p.text.strip()
            loc = f"Оглавление, абзац {p_num} ('{text[:30]}...')"

            # А) Запрет ручного проставления точек
            # Если в тексте найдено 3 или более точек подряд (классическая ошибка ручного выравнивания)
            if re.search(r'\.{3,}', text) or re.search(r'(\.\s*){4,}', text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Проставление точек не допустимо! Используйте стандартную табуляцию с заполнителем или авто-оглавление Word."
                    )
                )

            # Б) Выравнивание по ширине (JUSTIFY)
            fmt = p.paragraph_format
            align = p.alignment
            if align is None and p.style and p.style.paragraph_format:
                align = p.style.paragraph_format.alignment

            if align != WD_ALIGN_PARAGRAPH.JUSTIFY:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Пункт оглавления должен быть выровнен по ширине (JUSTIFY)."
                    )
                )

            # В) Отсутствие красной строки
            first_indent = fmt.first_line_indent
            if first_indent is None and p.style and p.style.paragraph_format:
                first_indent = p.style.paragraph_format.first_line_indent

            if first_indent is not None and first_indent.cm > 0.05:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"У пункта оглавления не должно быть красной строки (сейчас {round(first_indent.cm, 2)} см)."
                    )
                )

            # Г) Проверка на строго один пробел после номера главы/параграфа
            # Паттерны: "Глава 1. ", "1.1. ", "2.3.1. "
            num_match = re.match(r'^(Глава\s+\d+\.|Глава\s+\d+|\d+(\.\d+)*\.)(\s+)', text, re.IGNORECASE)
            if num_match:
                spaces = num_match.group(3)
                if len(spaces) > 1:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Между номером главы/параграфа и названием должен быть строго один пробел."
                        )
                    )

        return violations

class SectionFormatRule(BaseRule):
    """
    Правило проверки оформления разделов (глав и параграфов):
    - Полужирное начертание заголовков.
    - Выравнивание заголовков по центру (с обязательной красной строкой).
    - Размещение крупных разделов с новой страницы.
    - Соблюдение отступов в виде пустых строк (Enter'ов) до и после заголовков.
    """

    MAJOR_HEADER_REGEX = r'^(ГЛАВА\s+\d+|СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)|ПРИЛОЖЕНИЕ)'
    SUBHEADING_REGEX = r'^\d+(\.\d+)+\.?\s+'

    def _is_empty(self, paragraph) -> bool:
        """Проверяет, является ли абзац пустым (нажат Enter)."""
        return not paragraph.text.strip()

    def _is_bold(self, paragraph) -> bool:
        """Проверяет, выделен ли текст абзаца полужирным шрифтом."""
        runs_with_text = [r for r in paragraph.runs if r.text.strip()]
        if runs_with_text:
            return all(r.bold for r in runs_with_text)
        if paragraph.style and paragraph.style.font:
            return paragraph.style.font.bold or False
        return False

    def _is_centered(self, paragraph) -> bool:
        """Проверяет выравнивание по центру."""
        align = paragraph.alignment
        if align is None and paragraph.style and paragraph.style.paragraph_format:
            align = paragraph.style.paragraph_format.alignment
        return align == WD_ALIGN_PARAGRAPH.CENTER

    def _has_red_line(self, paragraph) -> bool:
        """Проверяет наличие красной строки (первой линии отступа)."""
        fmt = paragraph.paragraph_format
        indent = fmt.first_line_indent
        if indent is None and paragraph.style and paragraph.style.paragraph_format:
            indent = paragraph.style.paragraph_format.first_line_indent
        return indent is not None and indent.cm > 0.05

    def _has_page_break_before(self, doc: Document, idx: int) -> bool:
        """Проверяет, начинается ли абзац с новой страницы."""
        if idx == 0:
            return True

        p = doc.paragraphs[idx]

        # 1. Параметр page_break_before у абзаца
        if p.paragraph_format.page_break_before:
            return True

        # 2. Наличие явного разрыва страницы в предшествующем абзаце
        prev_p = doc.paragraphs[idx - 1]
        xml_str = prev_p._element.xml
        if 'w:br' in xml_str and 'w:type="page"' in xml_str:
            return True

        # 3. Разрыв секции
        if 'w:sectPr' in xml_str:
            return True

        return False

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "SectionFormatRule"
        paragraphs = doc.paragraphs

        for idx, p in enumerate(paragraphs):
            text = p.text.strip()
            if not text:
                continue

            text_upper = text.upper()
            is_major = bool(re.match(self.MAJOR_HEADER_REGEX, text_upper, re.IGNORECASE))
            is_sub = bool(re.match(self.SUBHEADING_REGEX, text, re.IGNORECASE))

            if not (is_major or is_sub):
                continue

            header_type = "Глава/Крупный раздел" if is_major else "Параграф"
            loc = f"{header_type} (абзац {idx+1}: '{text[:30]}...')"

            # 1. Проверка полужирного начертания
            if not self._is_bold(p):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"{header_type} должен быть выделен полужирным шрифтом."
                    )
                )

            # 2. Проверка выравнивания по центру
            if not self._is_centered(p):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"Заголовок ({header_type.lower()}) должен быть выровнен по центру."
                    )
                )

            # 3. Проверка наличия красной строки (абзацного отступа)
            if not self._has_red_line(p):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"У заголовка ({header_type.lower()}) должна быть установлена красная строка."
                    )
                )

            # 4. Проверка правила «с новой страницы»
            starts_new_page = self._has_page_break_before(doc, idx)
            if is_major and not starts_new_page:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Крупные разделы (Глава, Содержание, Введение, Заключение, Список источников) должны начинаться с новой страницы."
                    )
                )
            elif is_sub and p.paragraph_format.page_break_before:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Параграф должен продолжаться с места окончания предыдущего (не должен принудительно начинаться с новой страницы)."
                    )
                )

            # 5. Проверка Enter'ов перед заголовком
            if is_sub:
                # Проверяем, идет ли параграф сразу после главы
                prev_non_empty_idx = idx - 1
                while prev_non_empty_idx >= 0 and self._is_empty(paragraphs[prev_non_empty_idx]):
                    prev_non_empty_idx -= 1

                prev_p_text = paragraphs[prev_non_empty_idx].text.strip().upper() if prev_non_empty_idx >= 0 else ""
                after_chapter = bool(re.match(self.MAJOR_HEADER_REGEX, prev_p_text, re.IGNORECASE))

                empty_count_before = 0
                check_i = idx - 1
                while check_i >= 0 and self._is_empty(paragraphs[check_i]):
                    empty_count_before += 1
                    check_i -= 1

                if after_chapter:
                    # После названия главы, перед параграфом остается 1 Enter
                    if empty_count_before != 1:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=loc,
                                message=f"После названия главы перед параграфом должен быть ровно 1 Enter (обнаружено пустых строк: {empty_count_before})."
                            )
                        )
                else:
                    # Перед началом очередного параграфа ставится 2 Enter'а
                    if empty_count_before != 2 and not starts_new_page:
                        violations.append(
                            Violation(
                                rule_name=rule_name,
                                severity=Severity.WARNING,
                                location=loc,
                                message=f"Перед началом параграфа должно ставиться 2 Enter'а (обнаружено пустых строк: {empty_count_before})."
                            )
                        )

            # 6. Проверка Enter'а после заголовка (1 Enter после параграфа или главы)
            empty_count_after = 0
            check_after = idx + 1
            while check_after < len(paragraphs) and self._is_empty(paragraphs[check_after]):
                empty_count_after += 1
                check_after += 1

            if empty_count_after != 1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"После названия {header_type.lower()} должен оставаться ровно 1 Enter (обнаружено пустых строк: {empty_count_after})."
                    )
                )

        return violations

class SectionPageFillRule(BaseRule):
    """
    Правило проверки наполненности последней страницы раздела (не менее 70%).
    """

    MAJOR_HEADER_REGEX = r'^(ГЛАВА\s+\d+|СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)|ПРИЛОЖЕНИЕ)'

    def __init__(
        self,
        min_fill_ratio: float = 0.70,     # Минимальный процент заполнения (70%)
        chars_per_line: int = 70,         # Среднее кол-во символов в 1 строке (A4, TNR 14pt)
        lines_per_full_page: int = 28     # Среднее кол-во строк на полной странице A4
    ):
        self.min_fill_ratio = min_fill_ratio
        self.chars_per_line = chars_per_line
        self.lines_per_full_page = lines_per_full_page
        self.min_lines_required = math.ceil(lines_per_full_page * min_fill_ratio)

    def _has_page_break(self, paragraph) -> bool:
        """Проверяет наличие явного разрыва страницы перед абзацем или внутри него."""
        if paragraph.paragraph_format.page_break_before:
            return True
        xml_str = paragraph._element.xml
        return 'w:br' in xml_str and 'w:type="page"' in xml_str

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "SectionPageFillRule"
        paragraphs = doc.paragraphs

        section_blocks = []
        current_block = []

        # 1. Группируем абзацы по блокам перед разрывами страниц / новыми главами
        for idx, p in enumerate(paragraphs):
            text = p.text.strip()
            text_upper = text.upper()

            is_major = bool(re.match(self.MAJOR_HEADER_REGEX, text_upper, re.IGNORECASE))
            has_break = self._has_page_break(p)

            # Если начинается новая глава или стоит разрыв страницы — завершаем предыдущий блок
            if (is_major or has_break) and current_block:
                section_blocks.append((idx, current_block))
                current_block = []

            if text or 'w:drawing' in p._element.xml or 'w:pict' in p._element.xml:
                current_block.append(p)

        # 2. Проверяем наполненность каждого завершающего блока раздела
        for end_idx, block_paragraphs in section_blocks:
            if not block_paragraphs:
                continue

            total_estimated_lines = 0

            for p in block_paragraphs:
                p_text = p.text.strip()

                # Считаем количество строк тексте
                if p_text:
                    lines_in_p = max(1, math.ceil(len(p_text) / self.chars_per_line))
                    total_estimated_lines += lines_in_p

                # Если в абзаце есть рисунок или диаграмма (рисунок занимает ~8-12 строк)
                xml_str = p._element.xml
                if 'w:drawing' in xml_str or 'w:pict' in xml_str:
                    total_estimated_lines += 10

            # Расчет процента заполнения
            fill_ratio = min(1.0, total_estimated_lines / self.lines_per_full_page)

            # 3. Фиксация ошибки, если заполнено менее 70%
            if fill_ratio < self.min_fill_ratio:
                last_text = block_paragraphs[-1].text.strip()[:30]
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Конец раздела перед абзацем {end_idx} ('{last_text}...')",
                        message=(
                            f"При окончании раздела страница заполнена ориентировочно на {int(fill_ratio * 100)}% "
                            f"(~{total_estimated_lines} из {self.lines_per_full_page} строк). "
                            f"Допустимо заполнение листа не менее, чем на {int(self.min_fill_ratio * 100)}%."
                        )
                    )
                )

        return violations

class ReferencesListRule(BaseRule):
    """
    Правило проверки списка источников и литературы:
    1. Формат нумерованного списка.
    2. Соблюдение алфавитного порядка.
    3. Запрет периодических изданий.
    4. Единый стиль записи автора (Фамилия И.О.).
    5. Электронные ресурсы размещаются ПОСЛЕ печатных изданий.
    6. Наличие даты обращения у электронных ресурсов.
    """

    HEADER_REGEX = r'^(СПИСОК\s+(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ|ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ))'
    URL_REGEX = r'(https?://[^\s]+|www\.[^\s]+|электронный\s+ресурс)'
    ACCESS_DATE_REGEX = r'дата\s+обращения:?\s*\d{2}\.\d{2}\.\d{4}'
    PERIODICAL_KEYWORDS = ['ЖУРНАЛ', 'ВЕСТНИК', 'ИЗВЕСТИЯ', 'ПЕРИОДИЧЕСКОЕ ИЗДАНИЕ', '№', 'ВЫП.']

    def _get_sort_key(self, text: str) -> str:
        """Удаляет порядковую нумерацию в начале строки для сравнения по алфавиту."""
        cleaned = re.sub(r'^\d+[\.\)]\s*', '', text).strip()
        cleaned = re.sub(r'^\s*\[\d+\]\s*', '', cleaned)
        return cleaned.lower()

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "ReferencesListRule"

        # 1. Нахождение заголовка "Список источников и литературы"
        ref_header_idx = None
        for idx, p in enumerate(doc.paragraphs):
            text = p.text.strip().upper()
            if re.match(self.HEADER_REGEX, text):
                ref_header_idx = idx
                break

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

        # 2. Сбор всех элементов списка литературы
        items = []
        for p_idx in range(ref_header_idx + 1, len(doc.paragraphs)):
            p = doc.paragraphs[p_idx]
            text = p.text.strip()
            if not text:
                continue

            # Завершение списка при старте нового раздела (например, Приложение)
            if text.upper().startswith("ПРИЛОЖЕНИЕ"):
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

        # 3. Анализ каждого источника
        for p_num, text in items:
            loc = f"Список источников, абзац {p_num} ('{text[:30]}...')"

            # А) Проверка нумерации
            if not re.match(r'^\d+[\.\)]\s*', text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Источник должен быть частью нумерованного списка (начинаться с '1. ')."
                    )
                )

            is_electronic = bool(re.search(self.URL_REGEX, text, re.IGNORECASE))

            # Б) Проверка последовательности (Электронные ресурсы ПОСЛЕ печатных)
            if is_electronic:
                seen_electronic = True
                electronic_items.append((p_num, text))

                # Проверка наличия даты обращения
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
                            message="Печатное издание находится после электронного ресурса. Электронные ресурсы должны прописываться строго ПОСЛЕ печатных изданий."
                        )
                    )
                printed_items.append((p_num, text))

            # В) Запрет периодических изданий
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

            # Г) Проверка стиля написания автора (Фамилия И.О.)
            clean_item_text = re.sub(r'^\d+[\.\)]\s*', '', text).strip()
            author_match = re.match(r'^[А-ЯA-Z][а-яa-z]+\s+[А-ЯA-Z]\.[А-ЯA-Z]\.', clean_item_text)

            # Если не электронный ресурс и запись начинается не с названия
            if not author_match and not is_electronic and not clean_item_text.startswith("Developer"):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message="Проверьте стиль автора. Должен использоваться единый формат: Фамилия И.О. (например, 'Андреев А.В.')."
                    )
                )

        # 4. Проверка алфавитного порядка для печатных изданий
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

        # 5. Проверка алфавитного порядка для электронных ресурсов
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
