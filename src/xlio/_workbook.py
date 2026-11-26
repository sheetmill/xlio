"""The public reading layer: ``open``, ``check``, ``Workbook``, ``Sheet``, ``Parts``."""

from __future__ import annotations

import functools
import io
import os
import posixpath
import re
import xml.etree.ElementTree as ET
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from enum import StrEnum
from types import TracebackType
from typing import BinaryIO, Literal

from xlio import ns
from xlio._book import Book, CalcProperties, DefinedName, SheetInfo
from xlio._cells import _PLAIN_ERRORS, Cell
from xlio._comments import (
    Comment,
    Person,
    ThreadedComment,
    read_comments,
    read_persons,
    read_threaded,
)
from xlio._conditional import ConditionalFormat, Extension, read_conditional_formats
from xlio._drawing import Drawing, Vml, read_drawing, read_vml
from xlio._formulas import FormulaCell
from xlio._package import LimitError, Package, PackageError, ReadLimits
from xlio._print import PageBreak, PageSetup, PrintArea, PrintTitles, print_area, print_titles
from xlio._rich import RichString
from xlio._richvalue import RichValue, read_rich_values
from xlio._structure import (
    ColumnRange,
    Hyperlink,
    RowFormat,
    SheetFormat,
    SheetProperties,
    SheetView,
    Structure,
    read_structure,
    row_format,
)
from xlio._styles import Styles, check_font_sizes, count_cell_formats, read_styles
from xlio._tables import AutoFilter, SortState, Table, read_table
from xlio._theme import Theme, read_theme
from xlio._types import Unparsed, choice, flag, integer, local, unsigned

__all__ = [
    "BookView",
    "CellMetadata",
    "ErrorCode",
    "Parts",
    "Relationship",
    "Sheet",
    "SheetKind",
    "SheetState",
    "Workbook",
    "WorkbookKind",
    "WorkbookProperties",
    "check",
    "open",
]

SheetKind = Literal["worksheet", "chartsheet", "dialogsheet", "macrosheet", "other"]
SheetState = Literal["visible", "hidden", "veryHidden"]
WorkbookKind = Literal["workbook", "macro-enabled", "template", "macro-enabled template"]

_KINDS: dict[str, SheetKind] = {
    "worksheet": "worksheet",
    "chartsheet": "chartsheet",
    "dialogsheet": "dialogsheet",
    "xlMacrosheet": "macrosheet",
    "xlIntlMacrosheet": "macrosheet",
}
_CELL_SHEETS: tuple[SheetKind, ...] = ("worksheet", "macrosheet")  # sheets that hold cells
_STATES: dict[str, SheetState] = {
    "visible": "visible",
    "hidden": "hidden",
    "veryHidden": "veryHidden",
}
_WORKBOOK_KINDS: dict[str, WorkbookKind] = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbook",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "macro-enabled",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "template",
    "application/vnd.ms-excel.template.macroEnabled.main+xml": "macro-enabled template",
}
_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})")
_MAX_ROW = 1_048_576
_MAX_COL = 16_384


class ErrorCode(StrEnum):
    """An error value, as Excel shows it (capture ``errors``).

    A file stores the first seven, ``GETTING_DATA`` and ``BUSY`` as their text; Excel refuses a
    file with any other code written that way. It caches the others as ``#VALUE!`` with a rich
    value behind the cell's ``vm`` whose ``errorType`` names the error, and xlio reads them
    through it. ``EXTERNAL`` and ``PYTHON`` are errors Excel defines but has not been seen to
    show (Excel 16.115 for Mac): no ``errorType`` is known to stand for them.
    """

    NULL = "#NULL!"
    DIV0 = "#DIV/0!"
    VALUE = "#VALUE!"
    REF = "#REF!"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NA = "#N/A"
    GETTING_DATA = "#GETTING_DATA"
    """Excel reads it as written, in a cell's ``<v>`` or as a CSV field, and shows it while the
    workbook is open (``ERROR.TYPE`` 8; captures ``errors`` and ``csv_errors``). When Excel saves
    the file it writes it as ``#N/A``."""
    SPILL = "#SPILL!"
    """Rich value ``errorType`` 8."""
    CONNECT = "#CONNECT!"
    """Rich value ``errorType`` 9."""
    BLOCKED = "#BLOCKED!"
    """Rich value ``errorType`` 10."""
    UNKNOWN = "#UNKNOWN!"
    """Rich value ``errorType`` 11 (and 18, which Excel shows as this)."""
    FIELD = "#FIELD!"
    """Rich value ``errorType`` 12."""
    CALC = "#CALC!"
    """Rich value ``errorType`` 13, and 14, which Excel shows as this (its ``ERROR.TYPE`` is 15,
    not 14: this reading does not tell the two apart; capture ``errors_shown``)."""
    BUSY = "#BUSY!"
    """Rich value ``errorType`` 15, or written as text."""
    EXTERNAL = "#EXTERNAL!"
    """Not seen in a file, and no ``errorType`` is known to stand for it."""
    PYTHON = "#PYTHON!"
    """Not seen in a file, and no ``errorType`` is known to stand for it."""
    TIMEOUT = "#TIMEOUT!"
    """Rich value ``errorType`` 19."""


_ERRORS = {e.value: e for e in ErrorCode}


@dataclass(frozen=True, slots=True)
class Relationship:
    """A relationship from a part (or the package) to another part or an external target.

    Attributes:
        id: The relationship's id (``rId3``).
        type: Its type URI (constants in :mod:`xlio.ns`).
        target: The target as written.
        external: Whether the target is outside the package (``TargetMode="External"``).
        target_part: The target part's name, resolved against the source part's folder, or
            ``None`` for an external target.
    """

    id: str
    type: str
    target: str
    external: bool
    target_part: str | None


@dataclass(frozen=True, slots=True)
class CellMetadata:
    """What the workbook's metadata part says about cells.

    Attributes:
        dynamic_arrays: The ``cm`` indexes that mark a dynamic-array formula.
        value_errors: ``vm`` index → the error a rich value holds, for the errors a file caches
            as ``#VALUE!`` (a blocked spill is ``#SPILL!``). :meth:`Sheet.cells` already reads
            such a cell as its error.
    """

    dynamic_arrays: frozenset[int]
    value_errors: Mapping[int, ErrorCode]


Visibility = Literal["visible", "hidden", "veryHidden"]


@dataclass(frozen=True, slots=True)
class BookView:
    """A workbook window (``bookViews/workbookView``).

    Attributes:
        active_tab: The sheet the window shows, as an index into :attr:`Workbook.sheets`; 0
            when the value leads nowhere. The file counts ``<sheets>`` as written, a sheet
            without a part included; the index is converted.
        first_sheet: The first sheet whose tab is shown, converted likewise (the next sheet
            :attr:`Workbook.sheets` holds, if the one written has no part; 0 past them).
        show_sheet_tabs: Sheet tabs shown.
        show_horizontal_scroll: The horizontal scroll bar shown.
        show_vertical_scroll: The vertical scroll bar shown.
        tab_ratio: The tab bar's share of the scroll bar's width, in thousandths.
        visibility: ``visible``, ``hidden`` or ``veryHidden``.
        minimized: The window is minimized.
        auto_filter_date_grouping: Dates grouped in filter lists.
        x_window: The window's left edge in twentieths of a point, or ``None``.
        y_window: Its top edge, or ``None``.
        window_width: Its width, or ``None``.
        window_height: Its height, or ``None``.
        source: The element as written.
    """

    active_tab: int = 0
    first_sheet: int = 0
    show_sheet_tabs: bool = True
    show_horizontal_scroll: bool = True
    show_vertical_scroll: bool = True
    tab_ratio: int = 600
    visibility: Visibility = "visible"
    minimized: bool = False
    auto_filter_date_grouping: bool = True
    x_window: int | None = None
    y_window: int | None = None
    window_width: int | None = None
    window_height: int | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element, positions: Sequence[int]) -> BookView:
        """``positions``: each listed sheet's place in ``<sheets>`` as written."""
        get = element.get
        tab = unsigned(get("activeTab"), 0)
        first = unsigned(get("firstSheet"), 0)
        return cls(
            positions.index(tab) if tab in positions else 0,
            next((i for i, p in enumerate(positions) if p >= first), 0),
            flag(get("showSheetTabs"), True),
            flag(get("showHorizontalScroll"), True),
            flag(get("showVerticalScroll"), True),
            unsigned(get("tabRatio"), 600),
            choice(get("visibility"), ("visible", "hidden", "veryHidden"), "visible"),
            flag(get("minimized"), False),
            flag(get("autoFilterDateGrouping"), True),
            integer(get("xWindow")),
            integer(get("yWindow")),
            integer(get("windowWidth")),
            integer(get("windowHeight")),
            element,
        )


@dataclass(frozen=True, slots=True)
class WorkbookProperties:
    """The workbook's ``workbookPr``.

    Attributes:
        date1904: Dates count from 1904.
        show_objects: ``all``, ``placeholders`` or ``none``: how drawing objects are shown.
        code_name: The workbook's VBA code name, or ``None``.
        filter_privacy: Personal information removed on save.
        default_theme_version: The theme version the file was written for, or ``None``.
        update_links: ``userSet``, ``never`` or ``always``: external links updated on open.
        source: The element as written, or ``None`` when the workbook has none.
    """

    date1904: bool = False
    show_objects: Literal["all", "placeholders", "none"] = "all"
    code_name: str | None = None
    filter_privacy: bool = False
    default_theme_version: int | None = None
    update_links: Literal["userSet", "never", "always"] = "userSet"
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element) -> WorkbookProperties:
        get = element.get
        return cls(
            flag(get("date1904"), False),
            choice(get("showObjects"), ("all", "placeholders", "none"), "all"),
            get("codeName"),
            flag(get("filterPrivacy"), False),
            integer(get("defaultThemeVersion")),
            choice(get("updateLinks"), ("userSet", "never", "always"), "userSet"),
            element,
        )


_WORKBOOK_MODELED = frozenset({"sheets", "definedNames", "calcPr", "bookViews", "workbookPr"})


class Parts:
    """Every part of the package, read through the same checks and limits as the rest."""

    def __init__(self, package: Package) -> None:
        self._package = package
        self._types: tuple[dict[str, str], dict[str, str]] | None = None

    def __iter__(self) -> Iterator[str]:
        """Part names in archive order."""
        return (info.filename for info in self._package.infos())

    def __len__(self) -> int:
        return len(self._package.names)

    def __contains__(self, name: object) -> bool:
        return name in self._package.names

    def size(self, name: str) -> int:
        """A part's decompressed size, as the archive states it (and reading verifies)."""
        if name not in self._package.names:
            raise KeyError(name)
        return next(i.file_size for i in self._package.infos() if i.filename == name)

    def open(self, name: str) -> BinaryIO:
        """A part as a stream; ``KeyError`` when there is none by that name."""
        return self._package.part(name)

    def read(self, name: str) -> bytes:
        """A part's bytes."""
        return self._package.read(name)

    def xml(self, name: str) -> ET.Element:
        """A part parsed whole into an element tree (a DTD refused, its size bounded)."""
        return self._package.xml(name)

    def content_type(self, name: str) -> str | None:
        """A part's content type from ``[Content_Types].xml``: its override, else its default.

        Part names and extensions compare without case, as the packaging conventions say.
        """
        if self._types is None:
            overrides: dict[str, str] = {}
            defaults: dict[str, str] = {}
            if "[Content_Types].xml" in self._package.names:
                for element in self._package.xml("[Content_Types].xml"):
                    local = element.tag.rpartition("}")[2]
                    if local == "Override":
                        key = element.get("PartName", "").lstrip("/").lower()
                        overrides[key] = element.get("ContentType", "")
                    elif local == "Default":
                        key = element.get("Extension", "").lower()
                        defaults[key] = element.get("ContentType", "")
            self._types = overrides, defaults
        overrides, defaults = self._types
        found = overrides.get(name.lower())
        if found is not None:
            return found
        return defaults.get(posixpath.splitext(name)[1].lstrip(".").lower())

    def relationships(self, source: str) -> tuple[Relationship, ...]:
        """The relationships of a part, in file order; ``""`` for the package's own."""
        folder, base = posixpath.split(source)
        rels = posixpath.join(folder, "_rels", f"{base}.rels")
        if rels not in self._package.names:
            return ()
        out = []
        for rel in self._package.xml(rels).iter(f"{{{ns.PACKAGE_RELATIONSHIPS}}}Relationship"):
            target = rel.get("Target", "")
            external = rel.get("TargetMode") == "External"
            part = None if external else _resolve(folder, target)
            out.append(Relationship(rel.get("Id", ""), rel.get("Type", ""), target, external, part))
        return tuple(out)


def _resolve(folder: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(folder, target))


class Sheet:
    """A sheet of a workbook: its name, state and kind; a worksheet's or macro sheet's cells."""

    def __init__(self, workbook: Workbook, info: SheetInfo, index: int) -> None:
        self._workbook = workbook
        self._info = info
        self.name: str = info.name
        """The sheet's name."""
        self.index: int = index
        """Its position in :attr:`Workbook.sheets`, 0-based."""
        self.position: int = info.position
        """Its place in the workbook's ``<sheets>`` as written, 0-based, a sheet without a part
        counted: what a sheet-local defined name's ``sheet`` (``localSheetId``) refers to."""
        self.state: SheetState = _STATES.get(info.state, "visible")
        """``visible``, ``hidden`` or ``veryHidden``; another value reads as ``visible``."""
        self.kind: SheetKind = _KINDS.get(info.kind, "other")
        """``worksheet``, ``chartsheet``, ``dialogsheet``, ``macrosheet``, or ``other`` for a
        relationship type xlio does not know."""
        self.part: str = info.part
        """The name of the sheet's part."""
        self._row_formats: dict[int, RowFormat] | None = None  # gathered by a whole cells() pass

    def __repr__(self) -> str:
        return f"<Sheet {self.name!r} {self.kind} {self.state}>"

    @functools.cached_property
    def relationships(self) -> tuple[Relationship, ...]:
        """The relationships of the sheet's part, every type."""
        return self._workbook.parts.relationships(self.part)

    def _read_structure(self, head: bool) -> Structure:
        parts = self._workbook.parts
        if self.part not in parts:
            return read_structure(io.BytesIO(b"<worksheet/>"), self.part, {})
        # hyperlinks, whose targets these are, come after <sheetData>: a head read needs none
        targets = {} if head else {r.id: r.target for r in self.relationships if r.external}
        with parts.open(self.part) as f:
            return read_structure(f, self.part, targets, head=head)

    @functools.cached_property
    def _structure(self) -> Structure:
        structure = self._read_structure(head=False)
        if self._row_formats is None:
            self._row_formats = dict(structure.rows)
        return structure

    @functools.cached_property
    def _drops_cells(self) -> bool:
        """Whether Excel drops the sheet's columns, rows and cells (capture ``macrosheet``).

        So it opens a macro sheet whose root is not its own ``xm:macrosheet`` (one in
        SpreadsheetML's namespace, say), keeping the rest.
        """
        return self.kind == "macrosheet" and self._head.root != "{" + ns.XM + "}macrosheet"

    @functools.cached_property
    def _head(self) -> Structure:
        if "_structure" in self.__dict__:
            return self._structure
        return self._read_structure(head=True)

    # The structure. What comes before the cells (properties, views, format, columns) is read
    # from the part's head, up to <sheetData>; the rest takes one pass over the part on first
    # access, cells skipped. Row formats are gathered by a whole pass of cells() as well.

    @property
    def properties(self) -> SheetProperties:
        """The sheet's properties (``sheetPr``): tab color, outline directions, code name."""
        return self._head.properties

    @property
    def dimension(self) -> str | None:
        """The used range the file states (``dimension``), as written; not checked or used."""
        return self._head.dimension

    @property
    def views(self) -> tuple[SheetView, ...]:
        """The sheet's views (``sheetViews``), one per workbook window."""
        return self._head.views

    @property
    def format(self) -> SheetFormat:
        """Default sizes and outline levels (``sheetFormatPr``)."""
        return self._head.format

    @property
    def columns(self) -> tuple[ColumnRange, ...]:
        """Column formats (``cols``), in file order."""
        return () if self._drops_cells else self._head.columns

    @property
    def rows(self) -> Mapping[int, RowFormat]:
        """Row number (1-based) → what its ``<row>`` says, for rows that say anything.

        A ``<row>`` that holds no cell and sets nothing (``<row r="3"/>``) is here too, with
        an empty :class:`RowFormat` (every field its default): Excel counts the row.

        After a whole pass of :meth:`cells` this costs nothing: the rows were read with the cells.
        """
        if self._drops_cells:
            return {}
        if self._row_formats is None:
            return self._structure.rows
        return self._row_formats

    @property
    def merges(self) -> tuple[str, ...]:
        """Merged ranges (``mergeCells``), as written."""
        return self._structure.merges

    @property
    def hyperlinks(self) -> tuple[Hyperlink, ...]:
        """Hyperlinks, with external targets taken from their relationships."""
        return self._structure.hyperlinks

    @property
    def unparsed(self) -> tuple[Unparsed, ...]:
        """The sheet part's elements xlio does not model, whole and in file order."""
        return self._structure.unparsed

    def _target(self, rid: str | None) -> str | None:
        part = next((r.target_part for r in self.relationships if r.id == rid), None)
        return part if part is not None and part in self._workbook.parts else None

    def _related(self, suffix: str) -> list[str]:
        parts = self._workbook.parts
        return [
            r.target_part
            for r in self.relationships
            if r.type.endswith(suffix) and r.target_part is not None and r.target_part in parts
        ]

    @property
    def drawing_id(self) -> str | None:
        """The relationship id the sheet's ``<drawing>`` names (``r:id``), or ``None``.

        :attr:`relationships` gives its target. Unlike a lookup of the relationships by type,
        it names only a relationship the sheet's element refers to.
        """
        return self._structure.drawing_id

    @property
    def legacy_drawing_id(self) -> str | None:
        """The relationship id the sheet's ``<legacyDrawing>`` names (``r:id``), or ``None``."""
        return self._structure.legacy_drawing_id

    @property
    def picture_id(self) -> str | None:
        """The relationship id the sheet's ``<picture>`` (its background) names, or ``None``."""
        return self._structure.picture_id

    @property
    def table_ids(self) -> tuple[str, ...]:
        """The relationship ids the sheet's ``<tableParts>`` names, in file order."""
        return self._structure.table_ids

    @functools.cached_property
    def drawing(self) -> Drawing | None:
        """The sheet's drawing (``drawing``): its anchored objects in z-order, or ``None``.

        An anchor Excel would repair is read as repaired (``Anchor.repaired``).
        """
        part = self._target(self._structure.drawing_id)
        if part is None:
            return None
        parts = self._workbook.parts
        rels = {r.id: (r.target, r.target_part, r.external) for r in parts.relationships(part)}
        types = {t: parts.content_type(t) for _, t, external in rels.values() if t and not external}
        return read_drawing(parts.read(part), part, rels, types)

    @functools.cached_property
    def legacy_drawing(self) -> Vml | None:
        """The sheet's VML drawing (``legacyDrawing``: note boxes, form controls), or ``None``.

        Raises:
            LimitError: The part is larger than ``ReadLimits.max_vml_bytes``.
        """
        part = self._target(self._structure.legacy_drawing_id)
        if part is None:
            return None
        limit = self._workbook._book.package.limits.max_vml_bytes
        if self._workbook.parts.size(part) > limit:
            raise LimitError(f"{part}: more than the limit of {limit} bytes of VML")
        return read_vml(self._workbook.parts.read(part), part, limit)

    @property
    def background(self) -> str | None:
        """The part of the sheet's background picture (``picture``), or ``None``."""
        return self._target(self._structure.picture_id)

    @functools.cached_property
    def comments(self) -> tuple[Comment, ...]:
        """The sheet's notes, in file order; their boxes are in :attr:`legacy_drawing`."""
        parts = self._workbook.parts
        return tuple(
            c for part in self._related("/comments") for c in read_comments(parts.xml(part))
        )

    @functools.cached_property
    def threaded_comments(self) -> tuple[ThreadedComment, ...]:
        """The sheet's threaded comments, in file order, with their authors."""
        parts = self._workbook.parts
        persons = self._workbook.persons
        return tuple(
            c
            for part in self._related("/threadedComment")
            for c in read_threaded(parts.xml(part), persons)
        )

    @property
    def row_breaks(self) -> tuple[PageBreak, ...]:
        """Page breaks between rows (``rowBreaks``), in file order."""
        return self._structure.row_breaks

    @property
    def column_breaks(self) -> tuple[PageBreak, ...]:
        """Page breaks between columns (``colBreaks``), in file order."""
        return self._structure.column_breaks

    @property
    def page_setup(self) -> PageSetup:
        """Page setup, margins, print options, headers and footers."""
        return self._structure.page_setup

    @property
    def print_area(self) -> PrintArea | None:
        """The sheet's print area (``_xlnm.Print_Area`` scoped to it), or ``None``."""
        return print_area(self._workbook.defined_names, self.name, self._info.position)

    @property
    def print_titles(self) -> PrintTitles | None:
        """The sheet's print titles (``_xlnm.Print_Titles`` scoped to it), or ``None``."""
        return print_titles(self._workbook.defined_names, self.name, self._info.position)

    @property
    def extensions(self) -> tuple[Extension, ...]:
        """The sheet's ``extLst`` entries, every one, with their URIs exactly as written."""
        return self._structure.extensions

    @functools.cached_property
    def conditional_formats(self) -> tuple[ConditionalFormat, ...]:
        """The sheet's conditional formats, base rules joined with their ``x14`` rules.

        The formats in ``conditionalFormatting`` come first, in file order, then those that
        exist only in the extension. Reading them reads the workbook's styles (for the
        differential formats).
        """
        if not self._structure.conditional_formats and not self._structure.extensions:
            return ()
        return read_conditional_formats(
            self._structure.conditional_formats,
            self._structure.extensions,
            self._workbook.styles.differential_formats,
        )

    @property
    def autofilter(self) -> AutoFilter | None:
        """The sheet's autofilter, or ``None``; a table's own filter is on the table."""
        return self._structure.autofilter

    @property
    def sort_state(self) -> SortState | None:
        """The sort last applied to the sheet outside an autofilter, or ``None``."""
        return self._structure.sort_state

    @functools.cached_property
    def tables(self) -> tuple[Table, ...]:
        """The sheet's tables (``tableParts``), each read from its part, in file order.

        A table part the sheet names through a relationship that leads nowhere is left out.
        """
        targets = {r.id: r.target_part for r in self.relationships}
        parts = self._workbook.parts
        out = []
        for rid in self._structure.table_ids:
            part = targets.get(rid)
            if part is not None and part in parts:
                out.append(read_table(parts.xml(part), part))
        return tuple(out)

    def cells(self, ref: str | None = None, *, rich: bool = False) -> Iterator[Cell]:
        """The sheet's cells, streamed in file order, values as Excel reads them.

        Args:
            ref: A cell or range (``"B2"``, ``"B2:H40"``); only cells inside it are given.
            rich: Give string cells their runs and phonetic guides (:attr:`Cell.rich`); the
                shared strings are then read with their runs, once.

        Raises:
            ValueError: The sheet is neither a worksheet nor a macro sheet, or ``ref`` is not a
                cell or a range.
            LimitError: The range asked for, or without one the range the cells span, covers
                more cells than ``ReadLimits.max_cells``.
            PackageError: The part is refused.
        """
        if self.kind not in _CELL_SHEETS:
            raise ValueError(f"{self.name!r} is a {self.kind}, not a worksheet")
        limit = self._workbook._book.package.limits.max_cells
        if ref is not None:
            r1, c1, r2, c2 = _range(ref)
            if (r2 - r1 + 1) * (c2 - c1 + 1) > limit:
                raise LimitError(f"{ref} covers more than the limit of {limit} cells")
            for cell in self._read(rich):
                if r1 <= cell.row <= r2 and c1 <= cell.column <= c2:
                    yield cell
            return
        if limit >= _MAX_ROW * _MAX_COL:  # no sheet can pass the limit: nothing to count
            yield from self._read(rich)
            return
        top = left = _MAX_ROW + _MAX_COL
        bottom = right = 0
        for cell in self._read(rich):
            row, col = cell.row, cell.column
            if row < top or row > bottom or col < left or col > right:
                top, bottom = min(top, row), max(bottom, row)
                left, right = min(left, col), max(right, col)
                if (bottom - top + 1) * (right - left + 1) > limit:
                    raise LimitError(f"{self.name!r}: more than the limit of {limit} cells")
            yield cell

    def formulas(self) -> Iterator[FormulaCell]:
        """The sheet's formula cells only, streamed in file order, read without the others.

        For a caller that needs the formulas before, or without, the values: no other cell is
        read, and no value or style is parsed, so it costs a fraction of :meth:`cells` (the
        part is decompressed once more if :meth:`cells` reads it later). The fraction is small
        where values make most of the sheet, and large where formulas do: each formula cell
        costs about what :meth:`cells` spends on a cell. The dependent cells
        of a shared formula are given as :meth:`cells` gives them, with the ``si`` of their
        master and no text: xlio does not parse formulas, and a dependent's text is its
        master's moved to the cell, which takes a formula parser.

        What :meth:`cells` refuses about these cells is refused here too: a part that is not
        well-formed, a formula cell off the sheet, an ``si`` that is not a number, a ``cm`` past
        the cell-metadata table, and, under ``ReadLimits.max_cells``, formula cells spanning
        more cells than the limit. The other cells are not read, so what Excel refuses in them
        (a value it does not read, for instance) is refused by :meth:`cells` and
        :func:`xlio.check`, not here.

        Raises:
            ValueError: The sheet is neither a worksheet nor a macro sheet.
            LimitError: The formula cells span more cells than ``ReadLimits.max_cells``.
            PackageError: The part is refused.
        """
        if self.kind not in _CELL_SHEETS:
            raise ValueError(f"{self.name!r} is a {self.kind}, not a worksheet")
        if self._drops_cells:
            return
        book = self._workbook._book
        limit = book.package.limits.max_cells
        cells = book.formulas(self._info, book.metadata_counts[0])
        if limit >= _MAX_ROW * _MAX_COL:
            yield from cells
            return
        top = left = _MAX_ROW + _MAX_COL
        bottom = right = 0
        for cell in cells:
            row, col = cell.row, cell.column
            if row < top or row > bottom or col < left or col > right:
                top, bottom = min(top, row), max(bottom, row)
                left, right = min(left, col), max(right, col)
                if (bottom - top + 1) * (right - left + 1) > limit:
                    raise LimitError(f"{self.name!r}: more than the limit of {limit} cells")
            yield cell

    def _read(self, rich: bool) -> Iterator[Cell]:
        """The cells, row formats gathered on the way and kept once the part is read whole."""
        if self._drops_cells:
            return
        book = self._workbook._book
        formats: dict[int, RowFormat] = {}

        def on_row(row: int, attrs: dict[str, str]) -> None:
            formats[row] = row_format(attrs)

        styles = self._workbook._cell_format_count
        cms, vms = book.metadata_counts
        yield from book.cells(
            self._info,
            book.rich_strings if rich else None,
            on_row=on_row if self._row_formats is None else None,
            errors=_ERRORS,
            value_errors=book.value_errors,
            bounds=(styles or _MAX_ROW * _MAX_COL, cms, vms),
            plain_errors=self._workbook._plain_errors,
        )
        if self._row_formats is None:
            self._row_formats = formats


def _range(ref: str) -> tuple[int, int, int, int]:
    """``"B2:H40"`` → (2, 2, 40, 8), 1-based and ordered."""
    ends = ref.split(":")
    if not 1 <= len(ends) <= 2:
        raise ValueError(f"not a cell or a range: {ref!r}")
    points = []
    for end in ends:
        m = _REF.fullmatch(end.strip())
        if m is None:
            raise ValueError(f"not a cell or a range: {ref!r}")
        col = 0
        for ch in m[1].upper():
            col = col * 26 + ord(ch) - 64
        row = int(m[2])
        if not (1 <= row <= _MAX_ROW and 1 <= col <= _MAX_COL):
            raise ValueError(f"off the sheet: {ref!r}")
        points.append((row, col))
    (ra, ca), (rb, cb) = points[0], points[-1]
    return min(ra, rb), min(ca, cb), max(ra, rb), max(ca, cb)


class Workbook:
    """An xlsx workbook opened for reading. Use as a context manager, or call :meth:`close`.

    Only the workbook part is parsed when it is opened; other parts are read when asked for,
    then kept.
    """

    def __init__(self, book: Book) -> None:
        self._book = book
        self._closed = False
        # The error codes a cell may hold as text: Excel's, or every one for read_csv's workbook
        self._plain_errors = _PLAIN_ERRORS
        self.parts: Parts = Parts(book.package)
        """Every part of the package, raw."""
        self.sheets: tuple[Sheet, ...] = tuple(
            Sheet(self, info, k) for k, info in enumerate(book.sheets)
        )
        """Every sheet in workbook order, of every kind."""
        self.date1904: bool = book.date1904
        """Whether dates count from 1904 (``workbookPr date1904``)."""
        self.calc: CalcProperties = book.calc
        """The calculation settings (``calcPr``)."""
        self.defined_names: tuple[DefinedName, ...] = tuple(book.defined_names)
        """The defined names, in file order."""
        self.cell_metadata: CellMetadata = CellMetadata(
            book.dynamic_cm,
            {k: _ERRORS[v] for k, v in book.value_errors.items() if v in _ERRORS},
        )
        """Dynamic-array and rich-value metadata."""
        self.workbook_part: str = book.workbook_part
        """The name of the workbook part."""
        content_type = self.parts.content_type(book.workbook_part) or ""
        self.kind: WorkbookKind = _WORKBOOK_KINDS.get(content_type, "workbook")
        """``workbook``, ``macro-enabled``, ``template`` or ``macro-enabled template``, from the
        workbook part's content type (``workbook`` when it names none of them)."""
        self.has_vba: bool = any(
            r.type.endswith("/vbaProject") for r in self.parts.relationships(book.workbook_part)
        )
        """Whether the workbook has a VBA project (a macro-enabled file's macros)."""

    def __enter__(self) -> Workbook:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        """Close the file (a stream given to :func:`open` is left open).

        Closing a closed workbook does nothing.
        """
        self._book.close()
        self._closed = True

    @property
    def closed(self) -> bool:
        """The workbook is closed, by :meth:`close` or at the end of a ``with`` block.

        As a file's ``closed`` says: what is still to be read from its parts can no longer be read.
        """
        return self._closed

    def __len__(self) -> int:
        return len(self.sheets)

    def __iter__(self) -> Iterator[Sheet]:
        return iter(self.sheets)

    def __getitem__(self, key: str | int) -> Sheet:
        return self.sheet(key)

    def sheet(self, key: str | int) -> Sheet:
        """A sheet by name (compared without case, as Excel does) or by 0-based index.

        Raises:
            KeyError: No sheet has that name.
            IndexError: No sheet has that index.
        """
        if isinstance(key, int):
            return self.sheets[key]
        folded = key.casefold()
        for sheet in self.sheets:
            if sheet.name.casefold() == folded:
                return sheet
        raise KeyError(f"no sheet named {key!r}")

    @functools.cached_property
    def views(self) -> tuple[BookView, ...]:
        """The workbook's windows (``bookViews``)."""
        found = self._children("bookViews")
        return tuple(
            BookView.read(v, [s._info.position for s in self.sheets])
            for e in found
            for v in e
            if local(v.tag) == "workbookView"
        )

    @functools.cached_property
    def properties(self) -> WorkbookProperties:
        """The workbook's properties (``workbookPr``)."""
        found = self._children("workbookPr")
        return WorkbookProperties.read(found[0]) if found else WorkbookProperties()

    @functools.cached_property
    def unparsed(self) -> tuple[Unparsed, ...]:
        """The workbook part's elements xlio does not model, whole and in file order."""
        return tuple(
            Unparsed(f"workbook/{local(e.tag)}", e)
            for e in self._book.root
            if local(e.tag) not in _WORKBOOK_MODELED
        )

    @functools.cached_property
    def persons(self) -> Mapping[str, Person]:
        """The people of threaded comments, by id (the ``persons`` part)."""
        part = self._related("/person")
        return read_persons(self.parts.xml(part)) if part is not None else {}

    @functools.cached_property
    def external_links(self) -> tuple[str, ...]:
        """The parts of the workbook's links to other workbooks (``externalLink``).

        In the order of the workbook's relationships; their contents are not read.
        """
        return tuple(
            r.target_part
            for r in self.parts.relationships(self.workbook_part)
            if r.type.endswith("/externalLink") and r.target_part is not None
        )

    @functools.cached_property
    def theme(self) -> Theme | None:
        """The workbook's theme, read on first access; ``None`` when it has none."""
        part = self._related("/theme")
        return read_theme(self.parts.xml(part), part) if part is not None else None

    @functools.cached_property
    def rich_values(self) -> Mapping[int, RichValue]:
        """The rich value behind each value-metadata index (a cell's ``vm``), read on first access.

        A picture placed in a cell is an :class:`ImageValue`: the cell caches ``#VALUE!``, as
        :meth:`Sheet.cells` gives it, and Excel draws the picture instead. An error rich value
        is what :meth:`Sheet.cells` gives for its cell, whatever the cell caches. Other rich
        values (data types, the errors :class:`CellMetadata` already maps) leave the cell's
        value as cached and are :class:`RichValue` records, raw. Not to be confused with rich
        text (:attr:`Cell.rich`).
        """
        parts = self.parts
        found: dict[str, str] = {}
        for r in parts.relationships(self.workbook_part):
            if r.target_part is not None and not r.external:
                found.setdefault(r.type.rsplit("/", 1)[-1].lower(), r.target_part)
        relations: list[str | None] = []
        rel_part = found.get("richvaluerel")
        if rel_part is not None and rel_part in parts:
            targets = {r.id: r.target_part for r in parts.relationships(rel_part)}
            for rel in parts.xml(rel_part):
                rid = next((v for k, v in rel.attrib.items() if k.endswith("}id")), None)
                relations.append(targets.get(rid) if rid is not None else None)
        structures, values = found.get("rdrichvaluestructure"), found.get("rdrichvalue")
        records = read_rich_values(
            parts.xml(structures) if structures is not None and structures in parts else None,
            parts.xml(values) if values is not None and values in parts else None,
            relations,
        )
        return {vm: records[i] for vm, i in self._book.value_rich.items() if 0 <= i < len(records)}

    @functools.cached_property
    def _cell_format_count(self) -> int:
        """How many cell formats the styles hold: what reading cells needs, and no more."""
        if "styles" in self.__dict__:
            return len(self.styles.cell_formats)
        part = self._related("/styles")
        if part is None or part not in self.parts:
            return 0
        with self.parts.open(part) as f:
            return count_cell_formats(f, part)

    @functools.cached_property
    def styles(self) -> Styles:
        """The workbook's styles, read on first access."""
        part = self._related("/styles")
        return read_styles(self.parts.xml(part), part) if part is not None else Styles()

    def _children(self, name: str) -> list[ET.Element]:
        return [e for e in self._book.root if local(e.tag) == name]

    def _related(self, suffix: str) -> str | None:
        for r in self.parts.relationships(self.workbook_part):
            if (
                r.type.endswith(suffix)
                and r.target_part is not None
                and r.target_part in self.parts
            ):
                return r.target_part
        return None

    @property
    def shared_strings(self) -> Sequence[str]:
        """The shared string table, plain text (phonetic runs left out), read once."""
        return self._book.shared_strings

    @property
    def rich_strings(self) -> Sequence[RichString]:
        """The shared string table with runs and phonetic guides, read once.

        Read apart from :attr:`shared_strings`, so the plain table never holds runs.
        """
        return self._book.rich_strings


def open(  # noqa: A001 - the module-level open() of zipfile's neighbors gzip and tarfile
    source: str | os.PathLike[str] | BinaryIO, *, limits: ReadLimits | None = None
) -> Workbook:
    """Open an xlsx, xlsm, xltx or xltm file for reading.

    The package is checked before anything is parsed (part count and sizes, compression
    methods, encryption, names used twice); then the workbook part is read.

    Args:
        source: A path, or a binary file open for reading and seekable (left open).
        limits: What reading may cost; :class:`ReadLimits` by default.

    Raises:
        PackageError: The file is not a readable xlsx package (an ``.xls`` or password-protected
            file included), or a part is refused.
        LimitError: The package exceeds ``limits``.
    """
    return Workbook(Book(source, limits))


def check(source: str | os.PathLike[str] | BinaryIO, *, limits: ReadLimits | None = None) -> None:
    """Check that Excel would open a package as it is, without a repair.

    Every part is inflated once as a stream: it must have exactly the size and checksum the
    archive states and must not declare a DTD; the workbook part must be there; the styles'
    font sizes and font colors must be ones Excel takes. Then every worksheet and macro sheet
    is read whole: its structure, its cells (with their style and metadata indexes), its
    drawing's anchors and its conditional formats, refusing what Excel refuses or would only
    open after a repair.
    This parses every sheet, so it costs about as much as reading every cell; :func:`open`
    reads lazily and refuses only what it reads.

    A known limit: xlio does not parse formulas, so a file Excel repairs because a formula does
    not parse (a conditional format's, for one: Excel removes the format) passes this check.

    Raises:
        PackageError: The file is refused, or Excel would repair it.
        LimitError: The package exceeds ``limits``.
    """
    with Book(source, limits) as book:
        package = book.package
        for name in sorted(package.names):
            with package.part(name) as f:
                while f.read(1 << 16):
                    pass
        workbook = Workbook(book)
        styles = workbook._related("/styles")
        if styles is not None:
            with package.part(styles) as f:
                check_font_sizes(f, styles)
        for sheet in workbook.sheets:
            if sheet.kind in _CELL_SHEETS:
                _check_sheet(sheet)


def _check_sheet(sheet: Sheet) -> None:
    """Read a worksheet or macro sheet whole, refusing what Excel refuses or repairs (captures)."""
    sheet.unparsed  # noqa: B018 - the whole structure
    for _ in sheet.cells():
        pass
    drawing = sheet.drawing
    if drawing is not None and any(o.anchor.repaired for o in drawing.objects):
        raise PackageError(
            f"{sheet.name!r}: a two-cell anchor ends before it starts, which Excel repairs"
        )
    for cf in sheet.conditional_formats:
        for rule in cf.rules:
            if cf.extension:
                continue
            if rule.type is None:
                raise PackageError(
                    f"{sheet.name!r}: a conditional format of a type Excel repairs away"
                )
            if rule.dxf_id is not None and rule.dxf is None:
                raise PackageError(
                    f"{sheet.name!r}: a conditional format's dxfId past the table, which "
                    "Excel repairs"
                )
