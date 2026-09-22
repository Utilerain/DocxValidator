from dataclasses import dataclass
from operator import le
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from argparse import ArgumentParser
from enum import Enum
from typing import List, Optional
from pathlib import Path
from colorama import Fore

parser = ArgumentParser(description="Validate a Word document with diplom standards.")
parser.add_argument("input_file", help="Path to the input Word document.")

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
                    location=f"Раздел{i}",
                    message=f"Ожидалось: Л: {self.left}cm, П: {self.right}cm, В: {self.top}cm, Н: {self.bottom}\n" +
                            f"Фактически: Л:{left_cm}см, П:{right_cm}см, В:{top_cm}см, Н:{bottom_cm}см."
                ))
        return violations

class DocxValidator:
    def __init__(self, rules: List[BaseRule]) -> None:
        self.rules = rules
    def validate(self, file_path: str) -> List[Violation]:
        doc = Document(file_path)
        violations = []
        for rule in self.rules:
            violations.extend(rule.check(doc))
        return violations

if __name__ == "__main__":
    print("".rjust(70, "-"))
    print()
    print("DocxValidator".rjust(40, " "))
    print()
    print("".rjust(70, "-"))
    validator = DocxValidator(rules=[
        PageMarginsRule()
    ])
    docx_path = Path("./docs/test.docx")
    report = validator.validate(str(docx_path.absolute()))
    if report == []:
        print(Fore.GREEN + "Все в норме!" + Fore.WHITE)
        __import__("sys").exit(0)

    print(Fore.RED + f"Найдено нарушений: {len(report)}\n" + Fore.WHITE)
    for item in report:
        print(f"[{Fore.RED + item.severity.value + Fore.WHITE}] {item.rule_name} | {item.location}")
        print(f"    -> {item.message}")
