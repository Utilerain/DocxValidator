from docx_rules import *
from argparse import ArgumentParser, Namespace
from colors import print_color
from pathlib import Path
from colorama import Fore

class DocxValidator:
    def __init__(self, rules: List[BaseRule]) -> None:
        self.rules = rules
    def validate(self, file_path: str) -> List[Violation]:
        doc = Document(file_path)
        violations = []
        for rule in tqdm(self.rules, desc="Проверка документа на соответствия правилам..."):
            violations.extend(rule.check(doc))
        return violations

parser = ArgumentParser(description="Проверяет Word документ на правила, указанные на написание документации или диплома")
parser.add_argument("input_file", help="Путь до Word документа")
parser.print_help()

if __name__ == "__main__":
    args = parser.parse_args()

    print("".rjust(70, "-"))
    print()
    print("DocxValidator".rjust(40, " "))
    print()
    print("".rjust(70, "-"))
    validator = DocxValidator(rules=[
        PageMarginsRule(),
        MainTextFormatRule(),
        PageNumberingRule(),
        TitlePageRule(),
        SectionFormatRule(),
        SectionPageFillRule()
    ])
    docx_path = Path(args.input_file)
    report = validator.validate(str(docx_path.absolute()))
    if report == []:
        print_color("Все в норме!", Fore.GREEN)
        __import__("sys").exit(0)

    print_color(f"Найдено нарушений: {len(report)}\n", Fore.RED)
    for item in report:
        print_color(f"[{item.severity.value}] {item.rule_name} | {item.location}", Fore.RED)
        print(f"    -> {item.message}")
    __import__("sys").exit(1)
