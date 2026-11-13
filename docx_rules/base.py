"""Базовые типы всех правил: серьёзность нарушения, структура нарушения, базовый класс."""

from dataclasses import dataclass
from enum import Enum
from typing import List

from docx import Document


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
