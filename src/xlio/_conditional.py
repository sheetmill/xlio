"""Conditional formats: the sheet's rules, joined with their Excel 2010 (``x14``) halves.

A rule Excel 2010 extended is written twice: the base rule in ``conditionalFormatting``, whose
own ``extLst`` holds an ``x14:id``, and an ``x14:cfRule`` with that id in the sheet's ``extLst``
under the extension URI :data:`xlio.ns.EXT_CONDITIONAL_FORMATTINGS`. Excel reads that extension
only under that URI, compared exactly (``407d``, not the ``407D`` [MS-XLSX] prints); an entry
under any other spelling stays in :attr:`xlio.Sheet.extensions`, unread, as Excel leaves it.
Rules that exist only in the extension (new icon sets, cross-sheet operands) are their own
rules. Which rule wins, and what a rule's formulas mean, is the consumer's.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Literal

from xlio import ns
from xlio._styles import DifferentialFormat, _dxf
from xlio._types import Color, choice, flag, integer, local, ranges, unsigned

__all__ = [
    "CfType",
    "Cfvo",
    "CfvoType",
    "ColorScale",
    "ConditionalFormat",
    "DataBar",
    "Extension",
    "IconSet",
    "Rule",
    "read_conditional_formats",
]

CfType = Literal[
    "expression", "cellIs", "colorScale", "dataBar", "iconSet", "top10", "uniqueValues",
    "duplicateValues", "containsText", "notContainsText", "beginsWith", "endsWith",
    "containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors", "timePeriod",
    "aboveAverage",
]  # fmt: skip
_TYPES: tuple[CfType, ...] = CfType.__args__  # type: ignore[attr-defined]
CfvoType = Literal["num", "percent", "max", "min", "formula", "percentile", "autoMin", "autoMax"]
_CFVO_TYPES: tuple[CfvoType, ...] = CfvoType.__args__  # type: ignore[attr-defined]
Direction = Literal["context", "leftToRight", "rightToLeft"]
_DIRECTIONS: tuple[Direction, ...] = ("context", "leftToRight", "rightToLeft")
AxisPosition = Literal["automatic", "middle", "none"]
_AXES: tuple[AxisPosition, ...] = ("automatic", "middle", "none")


@dataclass(frozen=True, slots=True)
class Extension:
    """An entry of a sheet's ``extLst``: its URI exactly as written, and the element whole."""

    uri: str
    element: ET.Element = field(repr=False)


@dataclass(frozen=True, slots=True)
class Cfvo:
    """A threshold of a color scale, data bar or icon set.

    Attributes:
        type: How the value is taken (``num``, ``percent``, ``percentile``, ``formula``, ``min``,
            ``max``, and in the extension ``autoMin`` and ``autoMax``); ``None`` when the
            written type is none of them.
        value: The value or formula as written (``val``, or the extension's ``xm:f``), or
            ``None``.
        gte: The threshold includes its value (icon sets).
    """

    type: CfvoType | None
    value: str | None = None
    gte: bool = True


@dataclass(frozen=True, slots=True)
class ColorScale:
    """A color scale: its thresholds and their colors, in order."""

    cfvos: tuple[Cfvo, ...] = ()
    colors: tuple[Color, ...] = ()


@dataclass(frozen=True, slots=True)
class DataBar:
    """A data bar, the 2007 one or the extension's (whose fields beyond ``color`` it fills).

    Attributes:
        cfvos: The minimum and maximum thresholds.
        color: The bar's color (the 2007 ``color``, or the extension's ``fillColor``).
        min_length: The shortest bar, in percent of the cell.
        max_length: The longest bar, in percent of the cell.
        show_value: The cell's value is shown beside the bar.
        border: The bar has a border (extension).
        gradient: The bar has a gradient fill (extension; the 2007 bar always has).
        direction: ``context``, ``leftToRight`` or ``rightToLeft`` (extension).
        negative_bar_color_same_as_positive: Negative bars use ``color`` (extension).
        negative_bar_border_color_same_as_positive: Negative bars use ``border_color``.
        axis_position: ``automatic``, ``middle`` or ``none`` (extension).
        border_color: The border's color, or ``None``.
        negative_fill_color: Negative bars' fill, or ``None``.
        negative_border_color: Negative bars' border, or ``None``.
        axis_color: The axis's color, or ``None``.
    """

    cfvos: tuple[Cfvo, ...] = ()
    color: Color | None = None
    min_length: int = 10
    max_length: int = 90
    show_value: bool = True
    border: bool = False
    gradient: bool = True
    direction: Direction = "context"
    negative_bar_color_same_as_positive: bool = False
    negative_bar_border_color_same_as_positive: bool = True
    axis_position: AxisPosition = "automatic"
    border_color: Color | None = None
    negative_fill_color: Color | None = None
    negative_border_color: Color | None = None
    axis_color: Color | None = None


@dataclass(frozen=True, slots=True)
class IconSet:
    """An icon set.

    Attributes:
        icon_set: The set's name (``3TrafficLights1`` when absent), as written: the extension
            adds ``3Stars``, ``3Triangles``, ``5Boxes`` and ``NoIcons``.
        cfvos: The thresholds.
        show_value: The value is shown beside the icon.
        percent: Thresholds are percentages.
        reverse: The icons are in reverse order.
        custom: The icons are chosen one by one (``icons``).
        icons: The custom icons as (set, index) pairs.
    """

    icon_set: str = "3TrafficLights1"
    cfvos: tuple[Cfvo, ...] = ()
    show_value: bool = True
    percent: bool = True
    reverse: bool = False
    custom: bool = False
    icons: tuple[tuple[str, int], ...] = ()


@dataclass(frozen=True, slots=True)
class Rule:
    """A conditional formatting rule.

    Attributes:
        type: The rule's type, or ``None`` when absent or none of the 18 (``source`` has it).
        priority: Its priority (lower wins), or ``None``.
        formulas: Its formulas as written, without ``=``: the base rule's ``formula`` or the
            extension's ``xm:f``. Text, blank and error rules carry their standard formulas
            here, and Excel follows the formula, not the attributes.
        dxf: Its format: the differential format ``dxf_id`` names, or the extension's inline
            ``dxf``; ``None`` when there is none or the index leads nowhere.
        dxf_id: The index written, or ``None``.
        stop_if_true: Lower-priority rules are skipped where this one applies.
        operator: The ``cellIs`` or text operator as written, or ``None``.
        text: The text of a text rule, or ``None``.
        time_period: The period of a ``timePeriod`` rule, as written, or ``None``.
        rank: The N of a ``top10`` rule, or ``None``.
        percent: ``rank`` is a percentage.
        bottom: Bottom N rather than top.
        above_average: Above rather than below (``aboveAverage``).
        equal_average: The average itself is included.
        std_dev: Standard deviations from the average, or ``None``.
        color_scale: The color scale, or ``None``.
        data_bar: The data bar, or ``None``.
        icon_set: The icon set, or ``None``.
        id: The ``x14:id`` that pairs a base rule with its extension rule (or the extension
            rule's own ``id``), or ``None``.
        x14: For a base rule, its extension rule, joined by ``id``; ``None`` otherwise.
        extension: The rule comes from the sheet's ``extLst``.
        source: The element as written.
    """

    type: CfType | None
    priority: int | None = None
    formulas: tuple[str, ...] = ()
    dxf: DifferentialFormat | None = None
    dxf_id: int | None = None
    stop_if_true: bool = False
    operator: str | None = None
    text: str | None = None
    time_period: str | None = None
    rank: int | None = None
    percent: bool = False
    bottom: bool = False
    above_average: bool = True
    equal_average: bool = False
    std_dev: int | None = None
    color_scale: ColorScale | None = None
    data_bar: DataBar | None = None
    icon_set: IconSet | None = None
    id: str | None = None
    x14: Rule | None = None
    extension: bool = False
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class ConditionalFormat:
    """The rules over one range (``conditionalFormatting``).

    Attributes:
        sqref: The ranges, space-separated, as written.
        rules: The rules, in file order.
        pivot: The format belongs to a pivot table.
        extension: The format comes from the sheet's ``extLst`` (rules Excel 2010 added, with
            no base rule).
        source: The element as written.
    """

    sqref: str
    rules: tuple[Rule, ...] = ()
    pivot: bool = False
    extension: bool = False
    source: ET.Element | None = field(default=None, compare=False, repr=False)


def _children(element: ET.Element, name: str) -> list[ET.Element]:
    return [c for c in element if local(c.tag) == name]


def _first(element: ET.Element, name: str) -> ET.Element | None:
    return next((c for c in element if local(c.tag) == name), None)


def _cfvo(element: ET.Element) -> Cfvo:
    kind = element.get("type")
    f = _first(element, "f")
    value = f.text if f is not None else element.get("val")
    return Cfvo(
        choice(kind, _CFVO_TYPES, "num") if kind in _CFVO_TYPES else None,
        value,
        flag(element.get("gte"), True),
    )


def _color(element: ET.Element, name: str) -> Color | None:
    found = _first(element, name)
    return Color.read(found) if found is not None else None


def _data_bar(element: ET.Element) -> DataBar:
    get = element.get
    return DataBar(
        tuple(_cfvo(c) for c in _children(element, "cfvo")),
        _color(element, "color") or _color(element, "fillColor"),
        unsigned(get("minLength"), 10),
        unsigned(get("maxLength"), 90),
        flag(get("showValue"), True),
        flag(get("border"), False),
        flag(get("gradient"), True),
        choice(get("direction"), _DIRECTIONS, "context"),
        flag(get("negativeBarColorSameAsPositive"), False),
        flag(get("negativeBarBorderColorSameAsPositive"), True),
        choice(get("axisPosition"), _AXES, "automatic"),
        _color(element, "borderColor"),
        _color(element, "negativeFillColor"),
        _color(element, "negativeBorderColor"),
        _color(element, "axisColor"),
    )


def _icon_set(element: ET.Element) -> IconSet:
    get = element.get
    return IconSet(
        get("iconSet", "3TrafficLights1"),
        tuple(_cfvo(c) for c in _children(element, "cfvo")),
        flag(get("showValue"), True),
        flag(get("percent"), True),
        flag(get("reverse"), False),
        flag(get("custom"), False),
        tuple(
            (c.get("iconSet", ""), unsigned(c.get("iconId"), 0))
            for c in _children(element, "cfIcon")
        ),
    )


def _x14_id(element: ET.Element) -> str | None:
    """The ``x14:id`` in a base rule's ``extLst``, under its URI."""
    for ext_list in _children(element, "extLst"):
        for ext in _children(ext_list, "ext"):
            if ext.get("uri") == ns.EXT_CONDITIONAL_FORMATTING_ID:
                found = _first(ext, "id")
                if found is not None and found.text:
                    return found.text.strip()
    return None


def _rule(element: ET.Element, dxfs: Sequence[DifferentialFormat], *, extension: bool) -> Rule:
    get = element.get
    kind = get("type")
    if extension:
        formulas = tuple(f.text or "" for f in _children(element, "f"))
        inline = _first(element, "dxf")
        dxf, dxf_id = (_dxf(inline) if inline is not None else None), None
        rule_id = get("id")
    else:
        formulas = tuple(f.text or "" for f in _children(element, "formula"))
        dxf_id = integer(get("dxfId"))
        dxf = dxfs[dxf_id] if dxf_id is not None and 0 <= dxf_id < len(dxfs) else None
        rule_id = _x14_id(element)
    scale = _first(element, "colorScale")
    bar = _first(element, "dataBar")
    icons = _first(element, "iconSet")
    return Rule(
        choice(kind, _TYPES, "expression") if kind in _TYPES else None,
        integer(get("priority")),
        formulas,
        dxf,
        dxf_id,
        flag(get("stopIfTrue"), False),
        get("operator"),
        get("text"),
        get("timePeriod"),
        integer(get("rank")),
        flag(get("percent"), False),
        flag(get("bottom"), False),
        flag(get("aboveAverage"), True),
        flag(get("equalAverage"), False),
        integer(get("stdDev")),
        ColorScale(
            tuple(_cfvo(c) for c in _children(scale, "cfvo")),
            tuple(Color.read(c) for c in _children(scale, "color")),
        )
        if scale is not None
        else None,
        _data_bar(bar) if bar is not None else None,
        _icon_set(icons) if icons is not None else None,
        rule_id,
        None,
        extension,
        element,
    )


def _sqref(element: ET.Element) -> str:
    found = _first(element, "sqref")
    return (found.text or "").strip() if found is not None else element.get("sqref", "")


def read_conditional_formats(
    formats: Sequence[ET.Element],
    extensions: Sequence[Extension],
    dxfs: Sequence[DifferentialFormat],
) -> tuple[ConditionalFormat, ...]:
    """The sheet's conditional formats: the base ones in file order, then the extension's own.

    Args:
        formats: The sheet's ``conditionalFormatting`` elements.
        extensions: The sheet's ``extLst`` entries.
        dxfs: The workbook's differential formats.
    """
    extended: list[tuple[ET.Element, list[Rule]]] = []
    by_id: dict[str, Rule] = {}
    for ext in extensions:
        if ext.uri != ns.EXT_CONDITIONAL_FORMATTINGS:
            continue
        for group in ext.element:
            for cf in _children(group, "conditionalFormatting"):
                rules = [_rule(r, dxfs, extension=True) for r in _children(cf, "cfRule")]
                extended.append((cf, rules))
                by_id.update((r.id, r) for r in rules if r.id is not None)
    joined: set[str] = set()
    out = []
    for cf in formats:
        if ranges(_sqref(cf)) is None:
            continue  # Excel drops a format whose sqref is not a range list (capture repair)
        rules = []
        for element in _children(cf, "cfRule"):
            rule = _rule(element, dxfs, extension=False)
            if rule.id is not None and rule.id in by_id:
                joined.add(rule.id)
                rule = _with_x14(rule, by_id[rule.id])
            rules.append(rule)
        out.append(
            ConditionalFormat(_sqref(cf), tuple(rules), flag(cf.get("pivot"), False), False, cf)
        )
    for cf, rules in extended:
        own = tuple(r for r in rules if r.id is None or r.id not in joined)
        if own:
            out.append(ConditionalFormat(_sqref(cf), own, flag(cf.get("pivot"), False), True, cf))
    return tuple(out)


def _with_x14(rule: Rule, x14: Rule) -> Rule:
    import dataclasses

    return dataclasses.replace(rule, x14=x14)
