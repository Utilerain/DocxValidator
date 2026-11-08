from docx_rules import *
from docx_notes import make_notes_docx
from docx_fixed import create_fixed_docx
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

def default_rules():
    return [
        PageMarginsRule(),
        TextGridRule(),
        TableOfContentsRule(),
        MainTextFormatRule(),
        PageNumberingRule(),
        TitlePageRule(),
        SectionFormatRule(),
        SectionPageFillRule(),
        ReferencesListRule(),
        IntroductionContentRule(),
        ParagraphConclusionRule(),
        SourceReferencesRule(),
        AppendixFormatRule(),
        FiguresRule(),
        TablesFormatRule(),
        FormulasRule(),
        ListsFormatRule(),
        DashesRule(),
        WhiteFillRule()
    ]

if __name__ == "__main__":
    parser = ArgumentParser(description="Проверяет Word документ на правила, указанные на написание документации или диплома")
    parser.add_argument("input_file", help="Путь до Word документа")
    parser.add_argument("--make-notes", action="store_true",
                        help="Создать рядом копию документа с замечаниями в виде комментариев Word (оригинал не изменяется)")
    parser.add_argument("--create-fixed", action="store_true",
                        help="Создать рядом исправленную версию документа по требованиям оформления (оригинал не изменяется)")
    args = parser.parse_args()

    print("".rjust(70, "-"))
    print()
    print("DocxValidator".rjust(40, " "))
    print()
    print("".rjust(70, "-"))
    validator = DocxValidator(default_rules())
    docx_path = Path(args.input_file)
    if not docx_path.is_file():
        print_color(f"Файл не найден: {docx_path.absolute()}", Fore.RED)
        __import__("sys").exit(2)
    try:
        report = validator.validate(str(docx_path.absolute()))
    except Exception as error:
        print_color(f"Не удалось открыть или разобрать документ '{docx_path}': "
                    f"{type(error).__name__}: {error}", Fore.RED)
        print_color("Убедитесь, что это корректный файл Word (.docx), и он не поврежден.", Fore.RED)
        __import__("sys").exit(2)
    if report == []:
        print_color("Все в норме!", Fore.GREEN)
        if args.make_notes:
            print_color("Замечаний нет — копия с комментариями не требуется.", Fore.CYAN)
        if args.create_fixed:
            fixed_path = docx_path.with_name(docx_path.stem + "_исправленный.docx")
            create_fixed_docx(str(docx_path.absolute()), str(fixed_path))
            print_color(f"Исправленная версия сохранена: {fixed_path}", Fore.GREEN)
        __import__("sys").exit(0)

    print_color(f"Найдено нарушений: {len(report)}\n", Fore.RED)
    for item in report:
        print_color(f"[{item.severity.value}] {item.rule_name} | {item.location}", Fore.RED if item.severity.value == "ERR" else Fore.YELLOW)
        print(f"    -> {item.message}")

    if args.make_notes:
        notes_path = docx_path.with_name(docx_path.stem + "_замечания.docx")
        placed, unresolved = make_notes_docx(str(docx_path.absolute()), report, str(notes_path))
        print_color(f"\nСоздана копия с замечаниями: {notes_path} (комментариев: {placed}" +
                    (f", не привязано к абзацам: {unresolved}" if unresolved else "") + ")", Fore.GREEN)

    if args.create_fixed:
        fixed_path = docx_path.with_name(docx_path.stem + "_исправленный.docx")
        stats = create_fixed_docx(str(docx_path.absolute()), str(fixed_path))
        print_color(f"\nСоздана исправленная версия: {fixed_path}", Fore.GREEN)
        for name, count in stats.items():
            print(f"    исправлено «{name}»: {count}")
        fixed_report = validator.validate(str(fixed_path))
        print_color(f"Нарушений было: {len(report)}, в исправленной версии осталось: {len(fixed_report)}", Fore.CYAN)

    print_color("Внимание! Программа может ошибаться в поиске ошибок в формате документа! Если вы считаете, что нарушения ошибочны, то можете спокойно игнорировать", Fore.CYAN)
    __import__("sys").exit(1)
