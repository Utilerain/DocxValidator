"""Наборы правил валидатора для стандартного и ГОСТ-профилей."""
from typing import Callable, Dict, List

from docx_rules import (
    AppendixFormatRule,
    BaseRule,
    BreaksAndPaginationRule,
    DashesRule,
    FiguresRule,
    FormulasRule,
    IntroductionContentRule,
    ListsFormatRule,
    MainTextFormatRule,
    PageMarginsRule,
    PageNumberingRule,
    ParagraphConclusionRule,
    ReferencesListRule,
    SectionFormatRule,
    SectionPageFillRule,
    SourceReferencesRule,
    TableOfContentsRule,
    TablesFormatRule,
    Tz34StructureRule,
    Tz34TextFormatRule,
    Tz34TitlePageRule,
    TextGridRule,
    TitlePageRule,
    WhiteFillRule,
)
from gost import (
    AnnotationRule,
    Gost19201PageNumberingRule,
    Gost19201StructureRule,
    TechnicalSpecificationRequisitesRule,
    TechnicalSpecificationTablesRule,
)


PROFILE_STANDARD = "standard"
PROFILE_GOST_19_201 = "gost-19-201"
PROFILE_TZ_34_602 = "tz-34-602"
PROFILE_NAMES = (PROFILE_STANDARD, PROFILE_GOST_19_201, PROFILE_TZ_34_602)


def standard_rules() -> List[BaseRule]:
    """Возвращает новый набор общих правил при каждом вызове."""
    return [
        PageMarginsRule(),
        TextGridRule(),
        TableOfContentsRule(),
        MainTextFormatRule(),
        PageNumberingRule(),
        TitlePageRule(),
        SectionFormatRule(),
        BreaksAndPaginationRule(),
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
        WhiteFillRule(),
    ]


def gost_19_201_rules() -> List[BaseRule]:
    """Стандартный набор с профильными правилами ГОСТ 19.201-78.

    Правила курсовой работы, не предусмотренные структурой ТЗ
    (введение с объектом/предметом, список источников, ссылки [N],
    выводы «Таким образом» в конце параграфов), в этом профиле не
    запускаются: по ГОСТ 19.201-78 они обязательными разделами не являются.
    """
    coursework_only = (
        IntroductionContentRule,
        ReferencesListRule,
        SourceReferencesRule,
        ParagraphConclusionRule,
    )
    return [rule for rule in standard_rules()
            if type(rule) not in coursework_only] + [
        AnnotationRule(),
        Gost19201StructureRule(),
        TechnicalSpecificationRequisitesRule(),
        TechnicalSpecificationTablesRule(),
        Gost19201PageNumberingRule(),
    ]


def tz_34_602_rules() -> List[BaseRule]:
    """Оформление ТЗ на автоматизированную систему по образцу ПР2.

    Титул по ГОСТ 19.104 («УТВЕРЖДЕНО», «Листов N», «Литера») без учебной шапки;
    структура и порядок разделов по ГОСТ 34.602; текст Times New Roman 14 пт,
    интервал 1,5, по ширине, красная строка 1,25 см; поля правое 1 см; без
    обязательной сетки текста и нумерации страниц (в образце их нет).
    """
    return [
        PageMarginsRule(left=3.0, right=1.0, top=2.0, bottom=2.0),
        Tz34TitlePageRule(),
        Tz34StructureRule(),
        Tz34TextFormatRule(),
        BreaksAndPaginationRule(),
        # Списки в образце — автоматические, тире «–» на красной строке 1,25 см;
        # ручные «N. …» не проверяются: одноуровневые «N. …» в ТЗ — заголовки разделов.
        ListsFormatRule(auto_dash_indent_cm=1.25, manual_indent_cm=1.25, check_manual_items=False),
        DashesRule(),
        FiguresRule(),
        TablesFormatRule(require_header_repeat=False),
        WhiteFillRule(),
        AppendixFormatRule(),
    ]


_PROFILE_FACTORIES: Dict[str, Callable[[], List[BaseRule]]] = {
    PROFILE_STANDARD: standard_rules,
    PROFILE_GOST_19_201: gost_19_201_rules,
    PROFILE_TZ_34_602: tz_34_602_rules,
}


def get_rules(profile: str = PROFILE_STANDARD) -> List[BaseRule]:
    """Возвращает правила профиля или сообщает о неизвестном имени."""
    try:
        factory = _PROFILE_FACTORIES[profile]
    except KeyError as error:
        choices = ", ".join(PROFILE_NAMES)
        raise ValueError(f"Неизвестный профиль '{profile}'. Доступны: {choices}.") from error
    return factory()


def rules_for_profile(profile: str = PROFILE_STANDARD) -> List[BaseRule]:
    """Обратимо читаемый псевдоним для интеграции CLI."""
    return get_rules(profile)
