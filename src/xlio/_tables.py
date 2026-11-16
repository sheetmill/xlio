"""Tables, table styles, autofilters and sort state.

Absent flags of a table's style info read as false, as Excel draws them (seen for all 60
built-in table styles under three themes). Differential formats are given by index
(``dxf_id``), with the format resolved where the styles are at hand (table styles).
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from xlio._styles import DifferentialFormat
from xlio._types import Unparsed, choice, finite, flag, integer, local, unsigned

__all__ = [
    "AutoFilter",
    "ColorFilter",
    "CustomFilter",
    "DateGroup",
    "DynamicFilter",
    "FilterColumn",
    "IconFilter",
    "SortCondition",
    "SortState",
    "Table",
    "TableColumn",
    "TableStyle",
    "TableStyleElement",
    "TableStyleElementType",
    "TableStyleInfo",
    "TableStyles",
    "Top10",
    "read_autofilter",
    "read_sort_state",
    "read_table",
    "read_table_styles",
]

TableStyleElementType = Literal[
    "wholeTable", "headerRow", "totalRow", "firstColumn", "lastColumn", "firstRowStripe",
    "secondRowStripe", "firstColumnStripe", "secondColumnStripe", "firstHeaderCell",
    "lastHeaderCell", "firstTotalCell", "lastTotalCell", "firstSubtotalColumn",
    "secondSubtotalColumn", "thirdSubtotalColumn", "firstSubtotalRow", "secondSubtotalRow",
    "thirdSubtotalRow", "blankRow", "firstColumnSubheading", "secondColumnSubheading",
    "thirdColumnSubheading", "firstRowSubheading", "secondRowSubheading", "thirdRowSubheading",
    "pageFieldLabels", "pageFieldValues",
]  # fmt: skip
_ELEMENT_TYPES: tuple[TableStyleElementType, ...] = TableStyleElementType.__args__  # type: ignore[attr-defined]
SortBy = Literal["value", "cellColor", "fontColor", "icon"]
_SORT_BY: tuple[SortBy, ...] = ("value", "cellColor", "fontColor", "icon")


def _at(items: Sequence[DifferentialFormat], index: int | None) -> DifferentialFormat | None:
    return items[index] if index is not None and 0 <= index < len(items) else None


# Table styles ---------------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TableStyleElement:
    """One element of a table style: which part of a table it formats, and how.

    Attributes:
        type: The part of the table.
        dxf: The differential format it applies, or ``None`` when ``dxf_id`` leads nowhere.
        dxf_id: The format's index as written, or ``None``. A workbook's formats count from 0;
            the preset file's count from 1 for table styles and from 0 for pivot styles, which
            ``dxf`` already accounts for.
        size: The band size, for stripes (rows or columns per band).
    """

    type: TableStyleElementType
    dxf: DifferentialFormat | None
    dxf_id: int | None = None
    size: int = 1


@dataclass(frozen=True, slots=True)
class TableStyle:
    """A table or pivot table style: the file's own, or one of Excel's built-in presets.

    Attributes:
        name: Its name.
        pivot: Usable for pivot tables.
        table: Usable for tables.
        elements: Its elements, in file order.
        builtin: One of Excel's presets (:mod:`xlio.presets`).
        unparsed: Elements of an unknown type, whole.
        source: The element as written.
    """

    name: str
    pivot: bool = True
    table: bool = True
    elements: tuple[TableStyleElement, ...] = ()
    builtin: bool = False
    unparsed: tuple[Unparsed, ...] = ()
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    def element(self, type: TableStyleElementType) -> TableStyleElement | None:  # noqa: A002
        """The element for a part of the table, or ``None`` when the style has none."""
        return next((e for e in self.elements if e.type == type), None)


@dataclass(frozen=True, slots=True)
class TableStyles:
    """The workbook's table styles (``tableStyles``).

    Attributes:
        default_table_style: The style a new table gets, or ``None``.
        default_pivot_style: The style a new pivot table gets, or ``None``.
        styles: The workbook's own styles.
    """

    default_table_style: str | None = None
    default_pivot_style: str | None = None
    styles: tuple[TableStyle, ...] = ()


def _table_style(
    element: ET.Element, dxfs: Sequence[DifferentialFormat], *, builtin: bool
) -> TableStyle:
    # The preset file numbers its formats from 1 for table styles and from 0 for pivot styles
    # (all 144 checked); a workbook's styles part numbers them from 0 (ST_DxfId).
    base = 1 if builtin and not flag(element.get("pivot"), True) else 0
    elements = []
    unparsed = []
    for child in element:
        kind = child.get("type")
        if local(child.tag) != "tableStyleElement" or kind not in _ELEMENT_TYPES:
            unparsed.append(Unparsed(f"tableStyle/{local(child.tag)}", child))
            continue
        dxf_id = integer(child.get("dxfId"))
        elements.append(
            TableStyleElement(
                choice(kind, _ELEMENT_TYPES, "wholeTable"),
                _at(dxfs, dxf_id - base if dxf_id is not None else None),
                dxf_id,
                unsigned(child.get("size"), 1),
            )
        )
    return TableStyle(
        element.get("name", ""),
        flag(element.get("pivot"), True),
        flag(element.get("table"), True),
        tuple(elements),
        builtin,
        tuple(unparsed),
        element,
    )


def read_table_styles(
    element: ET.Element, dxfs: Sequence[DifferentialFormat], *, builtin: bool = False
) -> TableStyles:
    """A ``tableStyles`` element, its formats taken from ``dxfs``."""
    return TableStyles(
        element.get("defaultTableStyle"),
        element.get("defaultPivotStyle"),
        tuple(
            _table_style(s, dxfs, builtin=builtin) for s in element if local(s.tag) == "tableStyle"
        ),
    )


# Autofilters and sorting ----------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DateGroup:
    """A date group of a values filter (``dateGroupItem``): a year, or down to a second."""

    grouping: Literal["year", "month", "day", "hour", "minute", "second"]
    year: int
    month: int | None = None
    day: int | None = None
    hour: int | None = None
    minute: int | None = None
    second: int | None = None


@dataclass(frozen=True, slots=True)
class CustomFilter:
    """A comparison of a custom filter: its operator and the value, as written."""

    operator: Literal[
        "equal", "lessThan", "lessThanOrEqual", "notEqual", "greaterThanOrEqual", "greaterThan"
    ] = "equal"
    value: str | None = None


@dataclass(frozen=True, slots=True)
class Top10:
    """A top or bottom N (or N percent) filter."""

    top: bool = True
    percent: bool = False
    value: float = 10.0
    filter_value: float | None = None


@dataclass(frozen=True, slots=True)
class DynamicFilter:
    """A dynamic filter (above average, this month, ...): its type and the values computed."""

    type: str
    value: float | None = None
    max_value: float | None = None


@dataclass(frozen=True, slots=True)
class ColorFilter:
    """A filter by cell or font color, the color given by a differential format."""

    dxf_id: int | None
    cell_color: bool = True


@dataclass(frozen=True, slots=True)
class IconFilter:
    """A filter by conditional-format icon."""

    icon_set: str
    icon_id: int | None = None


@dataclass(frozen=True, slots=True)
class FilterColumn:
    """A filtered column of an autofilter (``filterColumn``).

    Attributes:
        column_id: The column, 0-based from the filter's first column.
        hidden_button: The filter button is hidden.
        show_button: The filter button is shown.
        values: The values kept by a values filter (``filters/filter``), as written.
        blank: A values filter keeps blanks.
        date_groups: The date groups a values filter keeps.
        custom: The comparisons of a custom filter.
        custom_and: Both comparisons must hold (else either).
        top10: A top or bottom filter, or ``None``.
        dynamic: A dynamic filter, or ``None``.
        color: A color filter, or ``None``.
        icon: An icon filter, or ``None``.
        source: The element as written.
    """

    column_id: int
    hidden_button: bool = False
    show_button: bool = True
    values: tuple[str, ...] = ()
    blank: bool = False
    date_groups: tuple[DateGroup, ...] = ()
    custom: tuple[CustomFilter, ...] = ()
    custom_and: bool = False
    top10: Top10 | None = None
    dynamic: DynamicFilter | None = None
    color: ColorFilter | None = None
    icon: IconFilter | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class SortCondition:
    """One key of a sort (``sortCondition``)."""

    ref: str
    descending: bool = False
    sort_by: SortBy = "value"
    custom_list: str | None = None
    dxf_id: int | None = None
    icon_set: str | None = None
    icon_id: int | None = None


@dataclass(frozen=True, slots=True)
class SortState:
    """The sort last applied (``sortState``).

    Attributes:
        ref: The range sorted.
        conditions: The sort keys, in order.
        column_sort: Columns were sorted (left to right) rather than rows.
        case_sensitive: The sort was case-sensitive.
        sort_method: ``pinYin``, ``stroke`` or ``None``.
        source: The element as written.
    """

    ref: str
    conditions: tuple[SortCondition, ...] = ()
    column_sort: bool = False
    case_sensitive: bool = False
    sort_method: Literal["pinYin", "stroke"] | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class AutoFilter:
    """An autofilter: its range, its filtered columns and the sort applied in it.

    Attributes:
        ref: The range, header row included; ``None`` when not written.
        columns: The columns with a filter or a button setting.
        sort_state: The sort inside the filter, or ``None``.
        source: The element as written.
    """

    ref: str | None
    columns: tuple[FilterColumn, ...] = ()
    sort_state: SortState | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


_OPERATORS = (
    "equal", "lessThan", "lessThanOrEqual", "notEqual", "greaterThanOrEqual", "greaterThan",
)  # fmt: skip
_GROUPINGS = ("year", "month", "day", "hour", "minute", "second")


def _optional_float(text: str | None) -> float | None:
    value = finite(text, float("nan"))
    return None if value != value else value


def _filter_column(element: ET.Element) -> FilterColumn:
    values: list[str] = []
    groups: list[DateGroup] = []
    custom: list[CustomFilter] = []
    blank = custom_and = False
    top10 = dynamic = color = icon = None
    for child in element:
        name = local(child.tag)
        if name == "filters":
            blank = flag(child.get("blank"), False)
            for item in child:
                if local(item.tag) == "filter":
                    values.append(item.get("val", ""))
                elif local(item.tag) == "dateGroupItem":
                    groups.append(
                        DateGroup(
                            choice(item.get("dateTimeGrouping"), _GROUPINGS, "year"),  # type: ignore[arg-type]
                            integer(item.get("year"), 0) or 0,
                            integer(item.get("month")),
                            integer(item.get("day")),
                            integer(item.get("hour")),
                            integer(item.get("minute")),
                            integer(item.get("second")),
                        )
                    )
        elif name == "customFilters":
            custom_and = flag(child.get("and"), False)
            custom.extend(
                CustomFilter(choice(c.get("operator"), _OPERATORS, "equal"), c.get("val"))  # type: ignore[arg-type]
                for c in child
                if local(c.tag) == "customFilter"
            )
        elif name == "top10":
            top10 = Top10(
                flag(child.get("top"), True),
                flag(child.get("percent"), False),
                finite(child.get("val"), 10.0),
                _optional_float(child.get("filterVal")),
            )
        elif name == "dynamicFilter":
            dynamic = DynamicFilter(
                child.get("type", ""),
                _optional_float(child.get("val")),
                _optional_float(child.get("maxVal")),
            )
        elif name == "colorFilter":
            color = ColorFilter(integer(child.get("dxfId")), flag(child.get("cellColor"), True))
        elif name == "iconFilter":
            icon = IconFilter(child.get("iconSet", ""), integer(child.get("iconId")))
    return FilterColumn(
        unsigned(element.get("colId"), 0),
        flag(element.get("hiddenButton"), False),
        flag(element.get("showButton"), True),
        tuple(values),
        blank,
        tuple(groups),
        tuple(custom),
        custom_and,
        top10,
        dynamic,
        color,
        icon,
        element,
    )


def read_sort_state(element: ET.Element) -> SortState:
    """A ``sortState`` element."""
    conditions = tuple(
        SortCondition(
            c.get("ref", ""),
            flag(c.get("descending"), False),
            choice(c.get("sortBy"), _SORT_BY, "value"),
            c.get("customList"),
            integer(c.get("dxfId")),
            c.get("iconSet"),
            integer(c.get("iconId")),
        )
        for c in element
        if local(c.tag) == "sortCondition"
    )
    method = element.get("sortMethod")
    return SortState(
        element.get("ref", ""),
        conditions,
        flag(element.get("columnSort"), False),
        flag(element.get("caseSensitive"), False),
        "pinYin" if method == "pinYin" else "stroke" if method == "stroke" else None,
        element,
    )


def read_autofilter(element: ET.Element) -> AutoFilter:
    """An ``autoFilter`` element."""
    columns = tuple(_filter_column(c) for c in element if local(c.tag) == "filterColumn")
    sort = next((read_sort_state(c) for c in element if local(c.tag) == "sortState"), None)
    return AutoFilter(element.get("ref"), columns, sort, element)


# Tables ---------------------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TableColumn:
    """A table's column.

    Attributes:
        id: Its id, unique in the table.
        name: Its name, the header text.
        totals_row_function: ``sum``, ``average``, ``count``, ``custom``, ... or ``None``.
        totals_row_label: The totals row's text for this column, or ``None``.
        totals_row_formula: The formula of a ``custom`` totals function, or ``None``.
        calculated_formula: The column's calculated formula, or ``None``.
        data_dxf_id: The differential format of its data cells, or ``None``.
        header_row_dxf_id: Of its header cell, or ``None``.
        totals_row_dxf_id: Of its totals cell, or ``None``.
    """

    id: int
    name: str
    totals_row_function: str | None = None
    totals_row_label: str | None = None
    totals_row_formula: str | None = None
    calculated_formula: str | None = None
    data_dxf_id: int | None = None
    header_row_dxf_id: int | None = None
    totals_row_dxf_id: int | None = None


@dataclass(frozen=True, slots=True)
class TableStyleInfo:
    """A table's style and which of its options are on; an absent option is off.

    Attributes:
        name: The table style's name (a preset or the workbook's own), or ``None``.
        show_first_column: First-column formatting.
        show_last_column: Last-column formatting.
        show_row_stripes: Banded rows.
        show_column_stripes: Banded columns.
    """

    name: str | None = None
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = False
    show_column_stripes: bool = False


@dataclass(frozen=True, slots=True)
class Table:
    """A table (a table part, through the sheet's ``tableParts``).

    Attributes:
        name: Its name.
        display_name: The name formulas use.
        ref: Its range, header and totals rows included.
        id: Its id, unique in the workbook.
        header_row_count: Header rows (0 or 1).
        totals_row_count: Totals rows (0 or 1).
        totals_row_shown: The totals row has been shown.
        columns: Its columns, in order.
        style: Its style and options, or ``None`` when it has none.
        autofilter: Its filter, or ``None``.
        sort_state: Its sort, or ``None``.
        table_type: ``worksheet``, ``xml`` or ``queryTable``.
        header_row_dxf_id: The header row's differential format, or ``None``.
        data_dxf_id: The data area's, or ``None``.
        totals_row_dxf_id: The totals row's, or ``None``.
        part: The table part's name.
        unparsed: The part's elements xlio does not model.
        source: The part's root element.
    """

    name: str
    display_name: str
    ref: str
    id: int | None = None
    header_row_count: int = 1
    totals_row_count: int = 0
    totals_row_shown: bool = True
    columns: tuple[TableColumn, ...] = ()
    style: TableStyleInfo | None = None
    autofilter: AutoFilter | None = None
    sort_state: SortState | None = None
    table_type: Literal["worksheet", "xml", "queryTable"] = "worksheet"
    header_row_dxf_id: int | None = None
    data_dxf_id: int | None = None
    totals_row_dxf_id: int | None = None
    part: str = ""
    unparsed: tuple[Unparsed, ...] = ()
    source: ET.Element | None = field(default=None, compare=False, repr=False)


_TABLE_MODELED = frozenset({"autoFilter", "sortState", "tableColumns", "tableStyleInfo"})


def _text_of(element: ET.Element, name: str) -> str | None:
    found = next((c for c in element if local(c.tag) == name), None)
    return found.text if found is not None else None


def read_table(root: ET.Element, part: str) -> Table:
    """A table part's root element."""
    get = root.get
    columns: list[TableColumn] = []
    style = autofilter = sort = None
    unparsed = []
    for child in root:
        name = local(child.tag)
        if name == "tableColumns":
            for c in child:
                if local(c.tag) != "tableColumn":
                    continue
                columns.append(
                    TableColumn(
                        unsigned(c.get("id"), 0),
                        c.get("name", ""),
                        c.get("totalsRowFunction"),
                        c.get("totalsRowLabel"),
                        _text_of(c, "totalsRowFormula"),
                        _text_of(c, "calculatedColumnFormula"),
                        integer(c.get("dataDxfId")),
                        integer(c.get("headerRowDxfId")),
                        integer(c.get("totalsRowDxfId")),
                    )
                )
        elif name == "tableStyleInfo":
            style = TableStyleInfo(
                child.get("name"),
                flag(child.get("showFirstColumn"), False),
                flag(child.get("showLastColumn"), False),
                flag(child.get("showRowStripes"), False),
                flag(child.get("showColumnStripes"), False),
            )
        elif name == "autoFilter":
            autofilter = read_autofilter(child)
        elif name == "sortState":
            sort = read_sort_state(child)
        if name not in _TABLE_MODELED:
            unparsed.append(Unparsed(f"table/{name}", child))
    return Table(
        get("name", ""),
        get("displayName", get("name", "")),
        get("ref", ""),
        integer(get("id")),
        unsigned(get("headerRowCount"), 1),
        unsigned(get("totalsRowCount"), 0),
        flag(get("totalsRowShown"), True),
        tuple(columns),
        style,
        autofilter,
        sort,
        choice(get("tableType"), ("worksheet", "xml", "queryTable"), "worksheet"),
        integer(get("headerRowDxfId")),
        integer(get("dataDxfId")),
        integer(get("totalsRowDxfId")),
        part,
        tuple(unparsed),
        root,
    )
