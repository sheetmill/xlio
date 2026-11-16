"""Printing: page breaks, page setup, margins, options, headers and footers, print areas.

Normal view draws none of these; they matter to a page layout or a report. A
break's ``id`` is zero-based: a row break of 10 is the boundary below row 10. The print area and
titles are defined names (``_xlnm.Print_Area``, ``_xlnm.Print_Titles``) scoped to the sheet by
its position in the workbook, hidden sheets included.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Literal

from xlio._book import DefinedName
from xlio._types import choice, finite, flag, integer, local, unsigned

__all__ = [
    "HeaderFooter",
    "PageBreak",
    "PageMargins",
    "PageSetup",
    "PrintArea",
    "PrintOptions",
    "PrintTitles",
    "print_area",
    "print_titles",
    "read_breaks",
    "read_header_footer",
    "read_page_setup",
]


@dataclass(frozen=True, slots=True)
class PageBreak:
    """A page break (``brk``).

    Attributes:
        id: The boundary: a row break of 10 falls below row 10, a column break of 6 right of
            column F.
        manual: Set by hand (``man``); breaks Excel computed may be stored too.
        min: The first cell the break spans, 0 when absent.
        max: The last, 0 when absent.
    """

    id: int
    manual: bool = False
    min: int = 0
    max: int = 0


def read_breaks(element: ET.Element) -> tuple[PageBreak, ...]:
    """A ``rowBreaks`` or ``colBreaks`` element's breaks, in file order."""
    return tuple(
        PageBreak(
            unsigned(b.get("id"), 0),
            flag(b.get("man"), False),
            unsigned(b.get("min"), 0),
            unsigned(b.get("max"), 0),
        )
        for b in element
        if local(b.tag) == "brk"
    )


@dataclass(frozen=True, slots=True)
class PageMargins:
    """Margins in inches (``pageMargins``).

    ``None`` for one the file leaves out (the schema requires all six).
    """

    left: float | None = None
    right: float | None = None
    top: float | None = None
    bottom: float | None = None
    header: float | None = None
    footer: float | None = None


@dataclass(frozen=True, slots=True)
class PrintOptions:
    """What is printed besides the cells (``printOptions``)."""

    grid_lines: bool = False
    headings: bool = False
    horizontal_centered: bool = False
    vertical_centered: bool = False


@dataclass(frozen=True, slots=True)
class HeaderFooter:
    """Headers and footers, in Excel's code form (``&L``, ``&P``, ...), as written."""

    odd_header: str | None = None
    odd_footer: str | None = None
    even_header: str | None = None
    even_footer: str | None = None
    first_header: str | None = None
    first_footer: str | None = None
    different_odd_even: bool = False
    different_first: bool = False
    scale_with_doc: bool = True
    align_with_margins: bool = True


@dataclass(frozen=True, slots=True)
class PageSetup:
    """A sheet's page setup.

    Attributes:
        paper_size: The paper size code (1 is Letter, 9 is A4), or ``None``.
        orientation: ``default``, ``portrait`` or ``landscape``.
        scale: The print scale in percent.
        fit_to_width: Pages wide, when fitting (``SheetProperties.fit_to_page``); 0 is automatic.
        fit_to_height: Pages tall, when fitting.
        first_page_number: The first page's number, used with ``use_first_page_number``.
        use_first_page_number: ``first_page_number`` applies.
        page_order: ``downThenOver`` or ``overThenDown``.
        black_and_white: Printed in black and white.
        draft: Draft quality.
        cell_comments: ``none``, ``asDisplayed`` or ``atEnd``.
        errors: How errors print: ``displayed``, ``blank``, ``dash`` or ``NA``.
        copies: Copies.
        margins: The margins, or ``None`` when absent.
        options: The print options, or ``None`` when absent.
        header_footer: Headers and footers, or ``None`` when absent.
        source: The ``pageSetup`` element as written, or ``None``.
    """

    paper_size: int | None = None
    orientation: Literal["default", "portrait", "landscape"] = "default"
    scale: int = 100
    fit_to_width: int = 1
    fit_to_height: int = 1
    first_page_number: int = 1
    use_first_page_number: bool = False
    page_order: Literal["downThenOver", "overThenDown"] = "downThenOver"
    black_and_white: bool = False
    draft: bool = False
    cell_comments: Literal["none", "asDisplayed", "atEnd"] = "none"
    errors: Literal["displayed", "blank", "dash", "NA"] = "displayed"
    copies: int = 1
    margins: PageMargins | None = None
    options: PrintOptions | None = None
    header_footer: HeaderFooter | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


def read_header_footer(element: ET.Element) -> HeaderFooter:
    """A ``headerFooter`` element."""
    texts = {local(c.tag): c.text or "" for c in element}
    get = element.get
    return HeaderFooter(
        texts.get("oddHeader"),
        texts.get("oddFooter"),
        texts.get("evenHeader"),
        texts.get("evenFooter"),
        texts.get("firstHeader"),
        texts.get("firstFooter"),
        flag(get("differentOddEven"), False),
        flag(get("differentFirst"), False),
        flag(get("scaleWithDoc"), True),
        flag(get("alignWithMargins"), True),
    )


def _margin(text: str | None) -> float | None:
    value = finite(text, float("nan"))
    return None if value != value else value


def read_page_setup(
    setup: ET.Element | None,
    margins: ET.Element | None,
    options: ET.Element | None,
    header_footer: ET.Element | None,
) -> PageSetup:
    """The page setup from its four elements, any of which may be absent."""
    get = setup.get if setup is not None else ET.Element("x").get
    first = integer(get("firstPageNumber"))
    return PageSetup(
        integer(get("paperSize")),
        choice(get("orientation"), ("default", "portrait", "landscape"), "default"),
        unsigned(get("scale"), 100),
        unsigned(get("fitToWidth"), 1),
        unsigned(get("fitToHeight"), 1),
        first if first is not None else 1,
        flag(get("useFirstPageNumber"), False),
        choice(get("pageOrder"), ("downThenOver", "overThenDown"), "downThenOver"),
        flag(get("blackAndWhite"), False),
        flag(get("draft"), False),
        choice(get("cellComments"), ("none", "asDisplayed", "atEnd"), "none"),
        choice(get("errors"), ("displayed", "blank", "dash", "NA"), "displayed"),
        unsigned(get("copies"), 1),
        PageMargins(
            *(
                _margin(margins.get(k))
                for k in ("left", "right", "top", "bottom", "header", "footer")
            )
        )
        if margins is not None
        else None,
        PrintOptions(
            flag(options.get("gridLines"), False),
            flag(options.get("headings"), False),
            flag(options.get("horizontalCentered"), False),
            flag(options.get("verticalCentered"), False),
        )
        if options is not None
        else None,
        read_header_footer(header_footer) if header_footer is not None else None,
        setup,
    )


# Print areas and titles ----------------------------------------------------------------------

_RANGE = re.compile(
    r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?[0-9]{1,7}:\$?[0-9]{1,7}"
)
_ROWS = re.compile(r"\$?[0-9]{1,7}:\$?[0-9]{1,7}")
_COLUMNS = re.compile(r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}")


@dataclass(frozen=True, slots=True)
class PrintArea:
    """A sheet's print area (``_xlnm.Print_Area``).

    Attributes:
        ranges: The ranges, in file order, as written without the sheet name (``$A$1:$C$5``,
            ``$A:$C``, ``$1:$5``); empty when none is usable.
        problem: Why the name, or a part of it, was not used (``#REF!``, a formula, another
            sheet's range, a name without a sheet scope), or ``None``.
        text: The name's text as written.
    """

    ranges: tuple[str, ...]
    problem: str | None
    text: str


@dataclass(frozen=True, slots=True)
class PrintTitles:
    """A sheet's print titles (``_xlnm.Print_Titles``): rows and columns repeated on each page.

    Attributes:
        rows: The title rows (``$1:$2``), or ``None``.
        columns: The title columns (``$A:$B``), or ``None``.
        problem: Why a part was not used, or ``None``.
        text: The name's text as written.
    """

    rows: str | None
    columns: str | None
    problem: str | None
    text: str


def _split(text: str) -> list[str]:
    """Split at top-level commas: not inside a quoted sheet name or a function's parentheses."""
    parts: list[str] = []
    current: list[str] = []
    quoted = False
    depth = 0
    for ch in text:
        if ch == "'":
            quoted = not quoted  # '' (an escaped quote) toggles twice
        elif not quoted and ch in "()":
            depth += 1 if ch == "(" else -1
        if ch == "," and not quoted and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
    parts.append("".join(current))
    return [p.strip() for p in parts]


def _sheet_and_range(part: str) -> tuple[str | None, str]:
    sheet, bang, rest = part.rpartition("!")
    if not bang:
        return None, part
    if sheet.startswith("'") and sheet.endswith("'") and len(sheet) >= 2:
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, rest


def _parts(name: DefinedName, sheet: str) -> tuple[list[str], str | None]:
    """The ranges of this sheet in a print name, and the first problem met."""
    ranges: list[str] = []
    problem = None
    for part in _split(name.text):
        if "#REF!" in part:
            problem = problem or "#REF!"
            continue
        owner, ref = _sheet_and_range(part)
        if owner is None or not _RANGE.fullmatch(ref):
            problem = problem or "a formula"
        elif owner.casefold() != sheet.casefold():
            problem = problem or f"a range of another sheet ({owner})"
        else:
            ranges.append(ref)
    return ranges, problem


def _named(names: Sequence[DefinedName], wanted: str, index: int) -> DefinedName | None:
    """The sheet's name of this kind.

    A global one is not valid and is not taken (it stays in ``Workbook.defined_names``).
    """
    folded = wanted.casefold()
    return next((n for n in names if n.name.casefold() == folded and n.sheet == index), None)


def print_area(names: Sequence[DefinedName], sheet: str, index: int) -> PrintArea | None:
    """A sheet's print area from the workbook's defined names, or ``None`` when it has none."""
    own = _named(names, "_xlnm.Print_Area", index)
    if own is None:
        return None
    ranges, problem = _parts(own, sheet)
    return PrintArea(tuple(ranges), problem, own.text)


def print_titles(names: Sequence[DefinedName], sheet: str, index: int) -> PrintTitles | None:
    """A sheet's print titles from the workbook's defined names, or ``None``."""
    own = _named(names, "_xlnm.Print_Titles", index)
    if own is None:
        return None
    ranges, problem = _parts(own, sheet)
    rows = next((r for r in ranges if _ROWS.fullmatch(r)), None)
    columns = next((r for r in ranges if _COLUMNS.fullmatch(r)), None)
    if len(ranges) > (rows is not None) + (columns is not None):
        problem = problem or "a range that is neither whole rows nor whole columns"
    return PrintTitles(rows, columns, problem, own.text)
