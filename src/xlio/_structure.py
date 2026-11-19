"""A worksheet's structure: everything in its part but the cells, from its head or one pass.

The cells are skipped as they stream past; only ``<row>`` attributes are kept. Every other child
of the worksheet is built as an element tree (they are small) and read into records; the children
not modeled are kept whole as :class:`Unparsed`.

Sizes follow Excel's reading, captured (``package``, with a Normal font of Calibri 11): a column
width is kept in points modulo 65,536 at 6 points a digit, a row height is its magnitude and
wraps in twips past 409.5, and the sheet's default sizes are ignored or refused as Excel ignores
or refuses them.
"""

from __future__ import annotations

import math
import xml.etree.ElementTree as ET
import xml.parsers.expat
from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import BinaryIO, Literal

from xlio._cells import ROW_ATTRIBUTES, _row
from xlio._conditional import Extension
from xlio._package import PackageError
from xlio._print import PageBreak, PageSetup, read_breaks, read_page_setup
from xlio._tables import AutoFilter, SortState, read_autofilter, read_sort_state
from xlio._types import (
    CellRef,
    Color,
    Unparsed,
    cell_ref,
    choice,
    finite,
    flag,
    integer,
    local,
    outline_level,
    ranges,
    sheet_outline_level,
    unsigned,
)

__all__ = [
    "ColumnRange",
    "Hyperlink",
    "Pane",
    "PaneName",
    "RowFormat",
    "Selection",
    "SheetFormat",
    "SheetProperties",
    "SheetView",
    "Structure",
    "read_structure",
]

_CHUNK = 1 << 16
_DIGIT_POINTS = 6  # at Calibri 11, the only font the width rule was captured with
_WIDTH_WRAP = 65_536  # points: Excel keeps a column's width in 16 bits
_WIDEST_FILE_WIDTH = 10_922.0  # the widest that does not wrap
_ONE_POINT_WIDTH = 0.1  # a file width of one point: what Excel makes of 1e9 to 7e305
_WIDTH_UNITS = 256  # a width overflowing in 256ths of a character is 0 (capture edges)
_INT_RANGE = 2**31
_ENORMOUS = 1e30  # past this, a height or a default width is 0; 1e30 and 1e308 are captured
_TALLEST = 409.5
_TWIPS = 20
_HEIGHT_WRAP = 65_536  # twips
_WIDEST_DEFAULT = 255.0  # 256 and past are refused (capture attributes)
_WIDEST_BASE = 255
_BASE_OF_255 = 10  # Excel shows columns of 10 characters for a baseColWidth of 255
_MODELED = frozenset(
    {
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "mergeCells",
        "hyperlinks",
        "autoFilter",
        "sortState",
        "tableParts",
        "conditionalFormatting",
        "extLst",
        "drawing",
        "legacyDrawing",
        "picture",
        "rowBreaks",
        "colBreaks",
        "pageSetup",
        "pageMargins",
        "printOptions",
        "headerFooter",
    }
)

PaneName = Literal["bottomRight", "topRight", "bottomLeft", "topLeft"]
_PANES: tuple[PaneName, ...] = ("bottomRight", "topRight", "bottomLeft", "topLeft")


def _measure(text: str) -> float | None:
    """A size attribute as Excel parses it, or ``None`` when it is not a number.

    ``nan``, ``inf`` and ``INF``, in any case, are taken (captured ``package``).
    """
    value = text.strip()
    if "_" in value or value.lower().startswith(("0x", "+0x", "-0x")):
        return None
    try:
        return float(value)
    except ValueError:
        return None


def column_width(text: str) -> float | None:
    """A ``<col width>`` as Excel takes it, in the file's units (captured ``package``).

    Kept in points modulo 65,536: a negative width or one past 10,922 wraps; 1e9 to 7e305, and
    -1e9 to -7e305, is one point; from 8e305 (where the width in 256ths of a character overflows
    a double), NaN and either infinity it is 0, a hidden column (capture ``edges``), and so is
    text that is not a number (captured ``attributes``).
    """
    width = _measure(text)
    if width is None:
        return 0.0  # a hidden column (capture attributes: "abc", "", "20pt")
    if not math.isfinite(width * _WIDTH_UNITS):
        return 0.0  # 8e305 and past, either sign (capture edges)
    if 0 <= width <= _WIDEST_FILE_WIDTH:
        return width
    points = width * _DIGIT_POINTS
    if abs(points) >= _INT_RANGE:
        return _ONE_POINT_WIDTH
    wrapped = math.fmod(points, _WIDTH_WRAP) % _WIDTH_WRAP
    return (wrapped + (1 if width < 0 else 0)) / _DIGIT_POINTS


def row_height(text: str) -> float | None:
    """A ``<row ht>`` as Excel takes it, in points, or ``None`` when Excel ignores it (captured).

    Its magnitude; past 409.5 it wraps in twips modulo 65,536 and is capped at 409.5; NaN and
    either infinity are ignored; past 1e30 it is 0. Empty text is 0; other text that is not a
    number is ignored (``None``), as is a number with space around it (captured ``attributes``).
    """
    if text == "":
        return 0.0  # Excel reads an empty height as 0, a hidden row (capture attributes)
    if text != text.strip():
        return None  # " 30" is ignored (capture attributes)
    value = _measure(text)
    if value is None:
        return None
    height = abs(value)
    if not math.isfinite(height):
        return None
    if height > _ENORMOUS:
        return 0.0
    if height <= _TALLEST:
        return height
    twips = height * _TWIPS
    if twips >= _INT_RANGE:
        return _TALLEST
    return min(_TALLEST, math.floor((twips % _HEIGHT_WRAP) / _TWIPS * 4) / 4)


@dataclass(frozen=True, slots=True)
class SheetProperties:
    """A worksheet's ``sheetPr``.

    Attributes:
        tab_color: The sheet tab's color, or ``None``.
        summary_below: Outline summary rows below their detail (``outlinePr``).
        summary_right: Outline summary columns right of their detail.
        apply_styles: Outline styles applied (``outlinePr applyStyles``).
        fit_to_page: Print scaled to fit pages (``pageSetUpPr fitToPage``).
        filter_mode: A filter is hiding rows.
        code_name: The sheet's VBA code name, or ``None``.
        published: Whether the sheet is published (``published``).
        show_outline_symbols: Outline symbols shown (``outlinePr showOutlineSymbols``).
        show_outline_symbols_text: That attribute as written, or ``None`` when the file leaves
            it out (``show_outline_symbols`` is then the default).
        source: The element as written, or ``None`` when the sheet has none.
    """

    tab_color: Color | None = None
    summary_below: bool = True
    summary_right: bool = True
    apply_styles: bool = False
    fit_to_page: bool = False
    filter_mode: bool = False
    code_name: str | None = None
    published: bool = True
    show_outline_symbols: bool = True
    show_outline_symbols_text: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element) -> SheetProperties:
        tab = outline = setup = None
        for child in element:
            name = local(child.tag)
            if name == "tabColor":
                tab = Color.read(child)
            elif name == "outlinePr":
                outline = child
            elif name == "pageSetUpPr":
                setup = child
        get = (outline if outline is not None else ET.Element("outlinePr")).get
        return cls(
            tab,
            flag(get("summaryBelow"), True),
            flag(get("summaryRight"), True),
            flag(get("applyStyles"), False),
            flag(setup.get("fitToPage"), False) if setup is not None else False,
            flag(element.get("filterMode"), False),
            element.get("codeName"),
            flag(element.get("published"), True),
            flag(get("showOutlineSymbols"), True),
            get("showOutlineSymbols"),
            element,
        )


@dataclass(frozen=True, slots=True)
class Pane:
    """A split or frozen pane (``pane``).

    Attributes:
        x_split: Columns frozen, or the horizontal split position in twentieths of a point.
        y_split: Rows frozen, or the vertical split position.
        top_left_cell: The top-left visible cell of the bottom-right pane, or ``None``.
        active_pane: The pane with the active cell.
        state: ``split``, ``frozen`` or ``frozenSplit``.
    """

    x_split: float = 0.0
    y_split: float = 0.0
    top_left_cell: str | None = None
    active_pane: PaneName = "topLeft"
    state: Literal["split", "frozen", "frozenSplit"] = "split"


_A1 = CellRef(1, 1)


@dataclass(frozen=True, slots=True)
class Selection:
    """A selection in a pane (``selection``): its active cell and the ranges selected.

    Attributes:
        pane: The pane it is in.
        active_cell: The active cell as Excel reads it (capture ``views``): ``activeCell`` when
            it is a cell inside ``sqref``; otherwise the first cell of ``sqref``'s first range,
            or A1 if ``sqref`` is not a list of cells and ranges.
        sqref: The cells and ranges selected, as written (``A1`` when absent).
        raw_active_cell: ``activeCell`` as written, or ``None``.
    """

    pane: PaneName = "topLeft"
    active_cell: CellRef = _A1
    sqref: str = "A1"
    raw_active_cell: str | None = None

    @classmethod
    def read(cls, element: ET.Element) -> Selection:
        raw = element.get("activeCell")
        sqref = element.get("sqref", "A1")
        selected = ranges(sqref)
        active = cell_ref(raw)
        if selected is None:
            active = _A1
        elif active is None or not any(
            a.row <= active.row <= b.row and a.column <= active.column <= b.column
            for a, b in selected
        ):
            active = selected[0][0]
        return cls(choice(element.get("pane"), _PANES, "topLeft"), active, sqref, raw)


@dataclass(frozen=True, slots=True)
class SheetView:
    """A window's view of the sheet (``sheetView``).

    Attributes:
        workbook_view_id: The workbook view (window) this belongs to.
        tab_selected: The sheet's tab is selected.
        view: ``normal``, ``pageBreakPreview`` or ``pageLayout``.
        zoom_scale: Zoom, in percent, as written, as Excel keeps it (capture ``views``; what it
            draws past 10 to 400 is not captured); 100 when absent or not a whole number.
        zoom_scale_normal: Zoom of the normal view, 0 when absent.
        zoom_scale_page_layout_view: Zoom of the page layout view, 0 when absent.
        zoom_scale_sheet_layout_view: Zoom of the page break preview, 0 when absent.
        top_left_cell: The top-left visible cell, or ``None``.
        show_grid_lines: Gridlines shown.
        default_grid_color: Gridlines in the automatic color; else ``color_id`` names it.
        color_id: The gridline color's palette index.
        show_row_col_headers: Row and column headings shown.
        show_zeros: Zero values shown.
        show_formulas: Formulas shown instead of values.
        show_outline_symbols: Outline symbols shown.
        show_ruler: The ruler shown in page layout view.
        show_white_space: White space shown in page layout view.
        right_to_left: Columns run right to left.
        window_protection: Panes cannot be moved.
        pane: The split or frozen pane, or ``None``.
        selections: The selections, one per pane at most.
        show_outline_symbols_text: ``showOutlineSymbols`` as written, or ``None`` when the file
            leaves it out (``show_outline_symbols`` is then the default).
        source: The element as written.
    """

    workbook_view_id: int = 0
    tab_selected: bool = False
    view: Literal["normal", "pageBreakPreview", "pageLayout"] = "normal"
    zoom_scale: int = 100
    zoom_scale_normal: int = 0
    zoom_scale_page_layout_view: int = 0
    zoom_scale_sheet_layout_view: int = 0
    top_left_cell: str | None = None
    show_grid_lines: bool = True
    default_grid_color: bool = True
    color_id: int = 64
    show_row_col_headers: bool = True
    show_zeros: bool = True
    show_formulas: bool = False
    show_outline_symbols: bool = True
    show_ruler: bool = True
    show_white_space: bool = True
    right_to_left: bool = False
    window_protection: bool = False
    pane: Pane | None = None
    selections: tuple[Selection, ...] = ()
    show_outline_symbols_text: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @property
    def active_cell(self) -> CellRef:
        """The active cell of the active pane's selection, else of the first selection, else A1."""
        active = self.pane.active_pane if self.pane is not None else "topLeft"
        chosen = next((s for s in self.selections if s.pane == active), None)
        if chosen is None:
            chosen = next(iter(self.selections), None)
        return chosen.active_cell if chosen is not None else _A1

    @classmethod
    def read(cls, element: ET.Element) -> SheetView:
        get = element.get
        pane = None
        selections = []
        for child in element:
            name = local(child.tag)
            if name == "pane":
                pane = Pane(
                    finite(child.get("xSplit"), 0.0),
                    finite(child.get("ySplit"), 0.0),
                    child.get("topLeftCell"),
                    choice(child.get("activePane"), _PANES, "topLeft"),
                    choice(child.get("state"), ("split", "frozen", "frozenSplit"), "split"),
                )
            elif name == "selection":
                selections.append(Selection.read(child))
        if pane is None and any(s.pane != "topLeft" for s in selections):
            # Excel refuses a selection for a pane the view does not have (capture views).
            raise PackageError("a selection for a pane the sheet view does not have")
        return cls(
            unsigned(get("workbookViewId"), 0),
            flag(get("tabSelected"), False),
            choice(get("view"), ("normal", "pageBreakPreview", "pageLayout"), "normal"),
            unsigned(get("zoomScale"), 100),
            unsigned(get("zoomScaleNormal"), 0),
            unsigned(get("zoomScalePageLayoutView"), 0),
            unsigned(get("zoomScaleSheetLayoutView"), 0),
            get("topLeftCell"),
            flag(get("showGridLines"), True),
            flag(get("defaultGridColor"), True),
            unsigned(get("colorId"), 64),
            flag(get("showRowColHeaders"), True),
            flag(get("showZeros"), True),
            flag(get("showFormulas"), False),
            flag(get("showOutlineSymbols"), True),
            flag(get("showRuler"), True),
            flag(get("showWhiteSpace"), True),
            flag(get("rightToLeft"), False),
            flag(get("windowProtection"), False),
            pane,
            tuple(selections),
            get("showOutlineSymbols"),
            element,
        )


@dataclass(frozen=True, slots=True)
class SheetFormat:
    """The sheet's default sizes and outline levels (``sheetFormatPr``), as Excel reads them.

    Attributes:
        base_col_width: Characters of the default column width before padding; 8 when absent,
            negative or past an integer (captured), or not an integer; 10 when written 255, as
            Excel shows it, and as written from 0 to 254 (capture ``edges``).
        default_col_width: The default column width, or ``None`` when absent or ignored (NaN,
            infinity); past 1e30 it is 0 (captured).
        default_row_height: The default row height in points (magnitude, wrapped as a row
            height), or ``None`` when absent. Excel uses it only with ``custom_height``.
        custom_height: The default row height is set by hand.
        zero_height: Rows are hidden by default.
        thick_top: Rows have a thick top border by default.
        thick_bottom: Rows have a thick bottom border by default.
        outline_level_row: The deepest row outline level.
        outline_level_col: The deepest column outline level.
        raw_default_col_width: ``defaultColWidth`` as written.
        raw_default_row_height: ``defaultRowHeight`` as written.
        source: The element as written, or ``None``.
    """

    base_col_width: int = 8
    default_col_width: float | None = None
    default_row_height: float | None = None
    custom_height: bool = False
    zero_height: bool = False
    thick_top: bool = False
    thick_bottom: bool = False
    outline_level_row: int = 0
    outline_level_col: int = 0
    raw_default_col_width: str | None = None
    raw_default_row_height: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element) -> SheetFormat:
        """Raises ``PackageError`` for the values Excel refuses to open (captured)."""
        get = element.get
        base_text = get("baseColWidth")
        base = 8
        if base_text is not None:
            value = integer(base_text)
            if value is not None and 0 <= value < _INT_RANGE:
                if value > _WIDEST_BASE:
                    raise PackageError(f"baseColWidth of {base_text[:40]!r}, which Excel refuses")
                base = _BASE_OF_255 if value == _WIDEST_BASE else value
        width_text = get("defaultColWidth")
        width = None
        if width_text is not None:
            value_f = _measure(width_text)
            if value_f is not None and not (math.isnan(value_f) or value_f == math.inf):
                if value_f < 0 or _ENORMOUS >= value_f > _WIDEST_DEFAULT:
                    raise PackageError(
                        f"defaultColWidth of {width_text[:40]!r}, which Excel refuses"
                    )
                width = 0.0 if value_f > _ENORMOUS else value_f
        height_text = get("defaultRowHeight")
        height = None
        if height_text is not None:
            parsed = _measure(height_text)
            height = row_height(height_text)
            if parsed is not None and height is None:  # NaN or infinity
                raise PackageError(f"defaultRowHeight of {height_text[:40]!r}, which Excel refuses")
        return cls(
            base,
            width,
            height,
            flag(get("customHeight"), False),
            flag(get("zeroHeight"), False),
            flag(get("thickTop"), False),
            flag(get("thickBottom"), False),
            sheet_outline_level(get("outlineLevelRow"), "outlineLevelRow"),
            sheet_outline_level(get("outlineLevelCol"), "outlineLevelCol"),
            width_text,
            height_text,
            element,
        )


@dataclass(frozen=True, slots=True)
class ColumnRange:
    """A ``<col>``: the format of columns ``min`` to ``max`` (1-based, inclusive).

    Attributes:
        min: The first column.
        max: The last column.
        width: The width as Excel takes it, derived for a Normal font of Calibri 11 (kept in
            points modulo 65,536 at its 6-point digit; NaN, infinity and 8e305 and past are 0),
            or ``None`` when absent or not a number. ``raw_width`` is the value stored.
        style: The index of the columns' cell format.
        hidden: The columns are hidden.
        best_fit: The width was set to fit the contents.
        custom_width: The width was set by hand.
        phonetic: Phonetic guides shown.
        outline_level: The outline level.
        collapsed: The outline is collapsed here.
        raw_width: ``width`` as written, the value Excel stored; the wrap loses it.
        source: The element as written.
    """

    min: int
    max: int
    width: float | None = None
    style: int = 0
    hidden: bool = False
    best_fit: bool = False
    custom_width: bool = False
    phonetic: bool = False
    outline_level: int = 0
    collapsed: bool = False
    raw_width: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class RowFormat:
    """What a ``<row>`` sets beside its cells: for a row that sets anything, or holds no cell.

    Attributes:
        height: The height in points as Excel takes it (its magnitude, wrapped and capped at
            409.5), or ``None`` when absent or ignored (NaN, infinity).
        custom_height: The height was set by hand.
        hidden: The row is hidden.
        style: The row's cell format index (``s``), applied with ``custom_format``; ``None``
            when absent.
        custom_format: The row's format applies to its empty cells.
        outline_level: The outline level.
        collapsed: The outline is collapsed here.
        thick_top: A thick top border.
        thick_bottom: A thick bottom border.
        phonetic: Phonetic guides shown.
        raw_height: ``ht`` as written; the wrap loses it.
    """

    height: float | None = None
    custom_height: bool = False
    hidden: bool = False
    style: int | None = None
    custom_format: bool = False
    outline_level: int = 0
    collapsed: bool = False
    thick_top: bool = False
    thick_bottom: bool = False
    phonetic: bool = False
    raw_height: str | None = None


def row_format(attrs: Mapping[str, str]) -> RowFormat:
    """A ``<row>``'s format, from its attributes."""
    ht = attrs.get("ht")
    s = attrs.get("s")
    return RowFormat(
        row_height(ht) if ht is not None else None,
        flag(attrs.get("customHeight"), False),
        flag(attrs.get("hidden"), False),
        unsigned(s, 0) if s is not None else None,
        flag(attrs.get("customFormat"), False),
        outline_level(attrs.get("outlineLevel")),
        flag(attrs.get("collapsed"), False),
        flag(attrs.get("thickTop"), False),
        flag(attrs.get("thickBot"), False),
        flag(attrs.get("ph"), False),
        ht,
    )


@dataclass(frozen=True, slots=True)
class Hyperlink:
    """A hyperlink on a cell or range (``hyperlink``).

    Attributes:
        ref: The cell or range it covers.
        target: The external target (a URL or file path) its relationship gives, or ``None``
            for a link inside the workbook or a relationship that is missing.
        location: A place in the workbook (``Sheet2!A1``, a defined name), or ``None``.
        tooltip: The tooltip, or ``None``.
        display: The display text, or ``None``.
        relationship_id: The ``r:id``, or ``None``.
        source: The element as written.
    """

    ref: str
    target: str | None = None
    location: str | None = None
    tooltip: str | None = None
    display: str | None = None
    relationship_id: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class Structure:
    """Everything one pass reads from a worksheet part, cells aside."""

    properties: SheetProperties
    dimension: str | None
    views: tuple[SheetView, ...]
    format: SheetFormat
    columns: tuple[ColumnRange, ...]
    rows: Mapping[int, RowFormat]
    merges: tuple[str, ...]
    hyperlinks: tuple[Hyperlink, ...]
    unparsed: tuple[Unparsed, ...]
    autofilter: AutoFilter | None = None
    sort_state: SortState | None = None
    table_ids: tuple[str, ...] = ()
    conditional_formats: tuple[ET.Element, ...] = ()
    extensions: tuple[Extension, ...] = ()
    drawing_id: str | None = None
    legacy_drawing_id: str | None = None
    picture_id: str | None = None
    row_breaks: tuple[PageBreak, ...] = ()
    column_breaks: tuple[PageBreak, ...] = ()
    page_setup: PageSetup = field(default_factory=PageSetup)
    root: str = ""
    """The root element's name, ``{namespace}local``."""


def _rel_id(element: ET.Element) -> str | None:
    for key, value in element.attrib.items():
        if key.endswith("}id"):
            return value
    return None


class _HeadReadError(Exception):
    """Raised by the parser's handler at ``<sheetData>`` to end a read of the head."""


def read_structure(
    stream: BinaryIO, part: str, targets: Mapping[str, str | None], *, head: bool = False
) -> Structure:
    """Read a worksheet part's structure; ``targets`` maps relationship ids to external targets.

    With ``head``, only what comes before ``<sheetData>`` is read (properties, dimension,
    views, format, columns); the rest of the structure is left empty. A record of the head
    written after ``<sheetData>`` is refused by the full read, as Excel refuses it; the head's
    read stops before it.

    Raises:
        PackageError: The part is not well-formed, a row is off the sheet, or a value is one
            Excel refuses.
    """
    children: list[ET.Element] = []
    rows: dict[int, RowFormat] = {}
    builder: ET.TreeBuilder | None = None
    level = 0
    in_data = seen_data = False
    after: list[int] = []  # indexes in children of the records written after <sheetData>
    row = empty_row = 0
    root = ""

    def clark(name: str) -> str:
        return "{" + name if "}" in name else name

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal builder, level, in_data, seen_data, row, root, empty_row
        level += 1
        if level == 1:
            root = clark(name)
        if in_data:
            if level == 3 and name.rpartition("}")[2] == "row":
                r = attrs.get("r")
                row = _row(r) if r else row + 1
                if row > 1_048_576:
                    raise PackageError(f"row {row} is off the sheet")
                if not ROW_ATTRIBUTES.isdisjoint(attrs):
                    rows[row] = row_format(attrs)
                else:
                    empty_row = row  # recorded at its end if it holds no cell
            elif level == 4 and name.rpartition("}")[2] == "c":
                empty_row = 0
            return
        if level == 2:
            if name.rpartition("}")[2] == "sheetData":
                if head:
                    raise _HeadReadError
                in_data = True
                seen_data = True
                return
            builder = ET.TreeBuilder()
        if builder is not None:
            builder.start(clark(name), {clark(k): v for k, v in attrs.items()})

    def end(name: str) -> None:
        nonlocal builder, level, in_data, empty_row
        level -= 1
        if in_data:
            if level == 1:
                in_data = False
            elif level == 2 and empty_row:  # a row without format attributes or cells
                rows[empty_row] = RowFormat()
                empty_row = 0
            return
        if builder is not None:
            builder.end(clark(name))
            if level == 1:
                if seen_data:
                    after.append(len(children))
                children.append(builder.close())
                builder = None

    def text(data: str) -> None:
        if builder is not None:
            builder.data(data)

    p = xml.parsers.expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    try:
        while chunk := stream.read(_CHUNK):
            p.Parse(chunk, False)
        p.Parse(b"", True)
    except _HeadReadError:
        pass
    except xml.parsers.expat.ExpatError as exc:
        raise PackageError(f"{part}: {exc}") from exc
    return _records(children, rows, targets, frozenset(after), root)


_HEAD = frozenset({"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols"})


def _records(
    children: list[ET.Element],
    rows: dict[int, RowFormat],
    targets: Mapping[str, str | None],
    after: frozenset[int] = frozenset(),
    root: str = "",
) -> Structure:
    properties = SheetProperties()
    dimension = None
    views: list[SheetView] = []
    sheet_format = SheetFormat()
    columns: list[ColumnRange] = []
    merges: list[str] = []
    links: list[Hyperlink] = []
    unparsed: list[Unparsed] = []
    autofilter = sort_state = None
    table_ids: list[str] = []
    formats: list[ET.Element] = []
    ids: dict[str, str | None] = {}
    row_breaks: tuple[PageBreak, ...] = ()
    column_breaks: tuple[PageBreak, ...] = ()
    printing: dict[str, ET.Element] = {}
    extensions: list[Extension] = []
    for i, child in enumerate(children):
        name = local(child.tag)
        if name in _HEAD and i in after:
            # Excel refuses a head record written after the cells (capture views).
            raise PackageError(f"worksheet/{name} after the cells, which Excel refuses")
        if name not in _MODELED:
            unparsed.append(Unparsed(f"worksheet/{name}", child))
        elif name == "sheetPr":
            properties = SheetProperties.read(child)
        elif name == "dimension":
            dimension = child.get("ref")
        elif name == "sheetViews":
            for view in child:
                if local(view.tag) == "sheetView":
                    views.append(SheetView.read(view))
                else:
                    unparsed.append(Unparsed(f"worksheet/sheetViews/{local(view.tag)}", view))
        elif name == "sheetFormatPr":
            sheet_format = SheetFormat.read(child)
        elif name == "cols":
            for col in child:
                if local(col.tag) != "col":
                    unparsed.append(Unparsed(f"worksheet/cols/{local(col.tag)}", col))
                    continue
                first, last = integer(col.get("min")), integer(col.get("max"))
                if first is None or last is None or first > last:
                    # Excel refuses the file (capture attributes: min="x", min past max).
                    raise PackageError(
                        f"a column range of {col.get('min')!r} to {col.get('max')!r}, "
                        "which Excel refuses"
                    )
                width_text = col.get("width")
                columns.append(
                    ColumnRange(
                        first,
                        last,
                        column_width(width_text) if width_text is not None else None,
                        unsigned(col.get("style"), 0),
                        flag(col.get("hidden"), False),
                        flag(col.get("bestFit"), False),
                        flag(col.get("customWidth"), False),
                        flag(col.get("phonetic"), False),
                        outline_level(col.get("outlineLevel")),
                        flag(col.get("collapsed"), False),
                        width_text,
                        col,
                    )
                )
        elif name == "rowBreaks":
            row_breaks = read_breaks(child)
        elif name == "colBreaks":
            column_breaks = read_breaks(child)
        elif name in ("pageSetup", "pageMargins", "printOptions", "headerFooter"):
            printing.setdefault(name, child)
        elif name in ("drawing", "legacyDrawing", "picture"):
            ids.setdefault(name, _rel_id(child))
        elif name == "conditionalFormatting":
            formats.append(child)
        elif name == "extLst":
            for ext in child:
                if local(ext.tag) == "ext":
                    extensions.append(Extension(ext.get("uri", ""), ext))
                else:
                    unparsed.append(Unparsed(f"worksheet/extLst/{local(ext.tag)}", ext))
        elif name == "autoFilter":
            autofilter = read_autofilter(child)
        elif name == "sortState":
            sort_state = read_sort_state(child)
        elif name == "tableParts":
            table_ids.extend(
                rid
                for t in child
                if (rid := _rel_id(t)) is not None and local(t.tag) == "tablePart"
            )
        elif name == "mergeCells":
            merges.extend(m.get("ref", "") for m in child if local(m.tag) == "mergeCell")
        elif name == "hyperlinks":
            for link in child:
                rid = _rel_id(link)
                links.append(
                    Hyperlink(
                        link.get("ref", ""),
                        targets.get(rid) if rid is not None else None,
                        link.get("location"),
                        link.get("tooltip"),
                        link.get("display"),
                        rid,
                        link,
                    )
                )
    return Structure(
        properties,
        dimension,
        tuple(views),
        sheet_format,
        tuple(columns),
        rows,
        tuple(merges),
        tuple(links),
        tuple(unparsed),
        autofilter,
        sort_state,
        tuple(table_ids),
        tuple(formats),
        tuple(extensions),
        ids.get("drawing"),
        ids.get("legacyDrawing"),
        ids.get("picture"),
        row_breaks,
        column_breaks,
        read_page_setup(
            printing.get("pageSetup"),
            printing.get("pageMargins"),
            printing.get("printOptions"),
            printing.get("headerFooter"),
        ),
        root,
    )
