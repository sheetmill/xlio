"""Typed readings of attribute values, and the records shared by several parts.

Each reading follows the schema type unless a captured Excel rule says otherwise; a text that
does not parse gives the schema's default, and the element it came from stays on the record
(``source``) with the text as written.
"""

from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import NamedTuple, TypeVar

from xlio._package import PackageError

__all__ = ["CellRef", "Color", "Unparsed"]

_DOUBLE = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?|[+-]?INF|NaN")
_HEX = re.compile(r"[0-9A-Fa-f]+")
_T = TypeVar("_T", bound=str)
_CELL = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})")
_MAX_ROW = 1_048_576
_MAX_COL = 16_384


class CellRef(NamedTuple):
    """A cell's position, 1-based; ``str()`` gives it in A1 form (``CellRef(3, 2)`` is ``B3``).

    Attributes:
        row: The row, 1 to 1,048,576.
        column: The column, 1 to 16,384.
    """

    row: int
    column: int

    def __str__(self) -> str:
        letters = ""
        col = self.column
        while col:
            col, rem = divmod(col - 1, 26)
            letters = chr(65 + rem) + letters
        return f"{letters}{self.row}"


def cell_ref(text: str | None) -> CellRef | None:
    """A cell in A1 form as Excel takes it (any case, ``$`` allowed, on the grid), else ``None``.

    A range, a list, R1C1 form, a sheet name or surrounding space is not a cell (capture
    ``views``).
    """
    if text is None:
        return None
    m = _CELL.fullmatch(text)
    if m is None:
        return None
    col = 0
    for ch in m[1].upper():
        col = col * 26 + ord(ch) - 64
    row = int(m[2])
    if not (1 <= row <= _MAX_ROW and 1 <= col <= _MAX_COL):
        return None
    return CellRef(row, col)


def ranges(sqref: str) -> list[tuple[CellRef, CellRef]] | None:
    """The ranges of a space-separated list (``A1 C3:D4``), each ordered; ``None`` if one is not."""
    out = []
    for item in sqref.split():
        ends = [cell_ref(e) for e in item.split(":")]
        if not 1 <= len(ends) <= 2 or None in ends:
            return None
        first, last = ends[0], ends[-1]
        assert first is not None
        assert last is not None
        out.append(
            (
                CellRef(min(first.row, last.row), min(first.column, last.column)),
                CellRef(max(first.row, last.row), max(first.column, last.column)),
            )
        )
    return out or None


def flag(text: str | None, default: bool) -> bool:
    """An ``xsd:boolean``: ``1``/``true`` or ``0``/``false``; anything else is the default."""
    if text is None:
        return default
    value = text.strip()
    if value in ("1", "true"):
        return True
    if value in ("0", "false"):
        return False
    return default


def integer(text: str | None, default: int | None = None) -> int | None:
    """An ``xsd:int`` or ``xsd:unsignedInt``, or the default when it is not one."""
    if text is None:
        return default
    value = text.strip()
    if value.lstrip("+-").isdigit() and value.isascii():
        return bounded_int(value)
    return default


INT_BOUND = 2**63 - 1
"""The magnitude a longer integer in a file is read as: past every table and limit, and never
converted from its digits (Python refuses to convert thousands of digits)."""


def bounded_int(digits: str) -> int:
    """Signed ASCII digits as an integer, bounded at ``INT_BOUND`` in magnitude."""
    body = digits.lstrip("+-").lstrip("0")
    if len(body) > 19:
        return -INT_BOUND if digits.startswith("-") else INT_BOUND
    value = int(digits)
    return max(-INT_BOUND, min(INT_BOUND, value))


def unsigned(text: str | None, default: int) -> int:
    """A non-negative integer (an index, a count, a level), or the default."""
    value = integer(text)
    return value if value is not None and value >= 0 else default


_DIGITS = re.compile(r"[0-9]+")


def _outline_byte(text: str | None) -> int | None:
    """An outline level as Excel takes it: decimal digits only, kept as a byte; else ``None``.

    The last eight digits are enough, since 10**8 is a multiple of 256.
    """
    if text is None or not _DIGITS.fullmatch(text):
        return None
    return int(text[-8:]) % 256


def outline_level(text: str | None) -> int:
    """A row's or column's ``outlineLevel`` as Excel reads it (capture ``outline``).

    Decimal digits, leading zeros allowed, are kept modulo 8 (9 is 1, 255 is 7, 256 is 0);
    anything else (``-1``, ``1.5``, ``+2``, `` 2``, empty) is no level. Captured up to 256;
    larger values follow the same rule.
    """
    value = _outline_byte(text)
    return value % 8 if value is not None else 0


def sheet_outline_level(text: str | None, name: str) -> int:
    """A sheet's ``outlineLevelRow`` or ``outlineLevelCol`` (capture ``outline``).

    Excel refuses a file stating 8 to 255; it opens one stating 256 or a value that is not
    decimal digits, and counts the levels from the rows and columns. Read so: 0 to 7 as they
    are, 256 as 0 (larger values as a byte, as for rows), anything else 0.

    Raises:
        PackageError: A level from 8 to 255.
    """
    value = _outline_byte(text)
    if value is None or text is None:
        return 0
    if value > 7:
        raise PackageError(f"{name} of {text[:40]!r}, which Excel refuses")
    return value


def double(text: str | None) -> float | None:
    """An ``xsd:double`` (``INF``, ``-INF`` and ``NaN`` included), or ``None``."""
    if text is None:
        return None
    value = text.strip()
    if not _DOUBLE.fullmatch(value):
        return None
    return float(value.replace("INF", "inf"))


def finite(text: str | None, default: float) -> float:
    """A finite double, or the default."""
    value = double(text)
    return value if value is not None and math.isfinite(value) else default


def choice(text: str | None, allowed: tuple[_T, ...], default: _T) -> _T:
    """A value of a closed set, or the default when the text is none of them."""
    for value in allowed:
        if text == value:
            return value
    return default


def local(tag: str) -> str:
    """An element's name without its namespace."""
    return tag.rpartition("}")[2]


@dataclass(frozen=True, slots=True)
class Color:
    """A color as SpreadsheetML gives it (``CT_Color``): one of RGB, theme, palette or automatic.

    Resolving a theme or palette color, and applying the tint, needs the theme and the palette
    (``Theme.colors``, ``Styles.indexed_colors``); the record keeps what the file says.

    Attributes:
        rgb: ``AARRGGBB`` in upper case, as Excel takes the text (fewer digits are a number,
            made opaque; text that is not hex is black), or ``None`` (absent, or more than
            eight digits); the text stays in ``source``.
        theme: The theme color index, or ``None``.
        indexed: The palette index, or ``None``.
        auto: Whether the color is automatic (the system's).
        tint: The tint, from -1 (darker) to 1 (lighter); 0 when absent; one outside wraps as
            Excel takes it.
        source: The element as written.
    """

    rgb: str | None = None
    theme: int | None = None
    indexed: int | None = None
    auto: bool = False
    tint: float = 0.0
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element) -> Color:
        """A color from its element (``color``, ``tabColor``, ``fgColor``, ...)."""
        rgb = element.get("rgb")
        if rgb is not None:
            rgb = _rgb(rgb.strip().upper())
        theme = integer(element.get("theme"))
        indexed = integer(element.get("indexed"))
        return cls(
            rgb,
            theme if theme is not None and theme >= 0 else None,
            indexed if indexed is not None and indexed >= 0 else None,
            flag(element.get("auto"), False),
            _tint(finite(element.get("tint"), 0.0)),
            element,
        )


def _rgb(text: str) -> str | None:
    """An ``rgb`` as Excel takes it (capture ``attributes``).

    Eight hex digits as written; fewer are a number whose low 24 bits are the color, made opaque
    (``FF00`` is ``FF00FF00``); text that is not hex is opaque black; more than eight digits is
    no color.
    """
    if len(text) == 8 and _HEX.fullmatch(text):
        return text
    if len(text) > 8:
        return None
    value = int(text, 16) if _HEX.fullmatch(text) else 0
    return f"FF{value & 0xFFFFFF:06X}"


def _tint(tint: float) -> float:
    """A tint as Excel takes it (capture ``attributes``).

    Within -1 to 1 as written; outside, counted in 32,767ths, truncated and wrapped as a signed
    16-bit number (2 is -2/32,767, -1.5 is 16,386/32,767).
    """
    if -1.0 <= tint <= 1.0:
        return tint
    scaled = tint * 32_767
    if not math.isfinite(scaled):
        return 0.0  # near the largest double: no count of 32,767ths to wrap (not captured)
    steps = int(scaled)
    return ((steps + 32_768) % 65_536 - 32_768) / 32_767


@dataclass(frozen=True, slots=True)
class Unparsed:
    """An element xlio reads but does not model, given whole so a consumer can report it.

    Attributes:
        path: Where it stands, as local names from the part's root (``worksheet/phoneticPr``).
        element: The element and everything inside it.
    """

    path: str
    element: ET.Element = field(repr=False)


def refuse(what: str, text: str) -> PackageError:
    """The error for a value Excel refuses to open."""
    return PackageError(f"{what} of {text[:40]!r}, which Excel refuses")
