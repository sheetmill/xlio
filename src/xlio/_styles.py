"""The workbook's styles part: number formats, fonts, fills, borders and the formats built of them.

Indices are resolved: a cell format holds its font, fill, border and named style, not their
numbers (the numbers stay beside them). An index that leads nowhere gives ``None``. Font sizes
are read as Excel reads them (captures ``package`` and ``edges``): the magnitude, from 1 to below
409.6, and any other size makes Excel refuse the file, so it is refused here too. So are a border
that writes ``left`` before ``start`` or ``right`` before ``end`` (capture ``edges``).
"""

from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
import xml.parsers.expat
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import BinaryIO, Literal, TypeVar

from xlio._package import PackageError
from xlio._tables import TableStyle, TableStyles, read_table_styles
from xlio._types import Color, Unparsed, choice, finite, flag, integer, local, unsigned

__all__ = [
    "Alignment",
    "Border",
    "BorderSide",
    "BorderStyle",
    "CellFormat",
    "DifferentialFormat",
    "Fill",
    "Font",
    "GradientFill",
    "GradientStop",
    "NamedStyle",
    "PatternFill",
    "PatternType",
    "Protection",
    "Styles",
    "Underline",
    "read_styles",
]

_RGB = re.compile(r"[0-9A-Fa-f]{8}")
# Excel keeps a font size in twentieths of a point, cut down, up to 8,191 (409.55): 409.56 is
# kept as 409.55 and 409.6 is refused (capture edges).
_SMALLEST, _TOO_LARGE = 1.0, 409.6
_T = TypeVar("_T")

Underline = Literal["single", "double", "singleAccounting", "doubleAccounting", "none"]
_UNDERLINES: tuple[Underline, ...] = (
    "single", "double", "singleAccounting", "doubleAccounting", "none",
)  # fmt: skip
PatternType = Literal[
    "none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical",
    "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical",
    "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625",
]  # fmt: skip
_PATTERNS: tuple[PatternType, ...] = PatternType.__args__  # type: ignore[attr-defined]
BorderStyle = Literal[
    "none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed",
    "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot",
]  # fmt: skip
_BORDERS: tuple[BorderStyle, ...] = BorderStyle.__args__  # type: ignore[attr-defined]
VertAlign = Literal["superscript", "subscript", "baseline"]
_VERT_ALIGNS: tuple[VertAlign, ...] = ("superscript", "subscript", "baseline")
FontScheme = Literal["major", "minor", "none"]
_SCHEMES: tuple[FontScheme, ...] = ("major", "minor", "none")
Horizontal = Literal[
    "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"
]
_HORIZONTALS: tuple[Horizontal, ...] = Horizontal.__args__  # type: ignore[attr-defined]
Vertical = Literal["top", "center", "bottom", "justify", "distributed"]
_VERTICALS: tuple[Vertical, ...] = Vertical.__args__  # type: ignore[attr-defined]


def font_size(text: str) -> float:
    """A font size as Excel takes it: its magnitude, from 1 to below 409.6.

    Captures ``package`` and ``edges``.

    Excel keeps the size in twentieths of a point, cut down (409.54 is 409.5, capture
    ``edges``); the size is returned as written, since that cut is known only there.

    Raises:
        PackageError: Any other size (0, 0.5, 409.6, NaN, infinity), which makes Excel refuse
            the file.
    """
    try:
        points = abs(float(text.strip())) if "_" not in text else math.nan
    except ValueError:
        points = math.nan
    if not _SMALLEST <= points < _TOO_LARGE:  # NaN compares false
        raise PackageError(f"a font size of {text[:40]!r}, which Excel refuses")
    return points


_RUN_LARGEST = 409.55


_THEME_COLORS = 12


def font_theme(text: str | None) -> None:
    """Refuse a font color's theme index past the theme's 12 colors, as Excel refuses the file.

    Capture ``theme_bound``: 11 opens, 12 and past are refused (a negative index is no color).

    Raises:
        PackageError: The index is 12 or more.
    """
    index = integer(text)
    if index is not None and index >= _THEME_COLORS:
        raise PackageError(f"a font color of theme {index}, which Excel refuses")


def run_font_size(text: str) -> float:
    """A rich-text run's font size as Excel takes it (capture ``attributes``).

    Its magnitude, from 1 to 409.55: 0 and -1 are 1, -2 is 2, 410 is 409.55, a size too large
    for a double is 1. Text that is not a number is read as 1 too (only a number too large was
    captured).
    """
    try:
        points = abs(float(text.strip())) if "_" not in text else math.nan
    except ValueError:
        points = math.nan
    if not math.isfinite(points):
        return 1.0
    return min(max(points, 1.0), _RUN_LARGEST)


def dxf_font_size(text: str) -> float:
    """A differential format's font size as Excel takes it (capture ``attributes``).

    Its magnitude, up to 409.55 (-2 is 2).

    Raises:
        PackageError: 0, a size past 409.55, or one that is not a finite number, which make
            Excel refuse the file.
    """
    try:
        points = abs(float(text.strip())) if "_" not in text else math.nan
    except ValueError:
        points = math.nan
    if not 0 < points <= _RUN_LARGEST:  # NaN compares false
        raise PackageError(f"a differential font size of {text[:40]!r}, which Excel refuses")
    return points


def _rotation(text: str | None) -> int:
    """``textRotation`` as Excel takes it: 0 to 255 as written, anything else 0.

    Capture ``attributes``: 181 to 254 are kept, though the schema stops at 180 and 255.
    """
    value = integer(text, 0) or 0
    return value if 0 <= value <= 255 else 0


def _val_flag(element: ET.Element | None) -> bool:
    """A ``CT_BooleanProperty`` (``<b/>``, ``<b val="0"/>``): present means true."""
    return element is not None and flag(element.get("val"), True)


def _children(element: ET.Element) -> dict[str, ET.Element]:
    found: dict[str, ET.Element] = {}
    for child in element:
        found.setdefault(local(child.tag), child)
    return found


@dataclass(frozen=True, slots=True)
class Font:
    """A font (``font``, or a rich-text run's ``rPr``).

    Attributes:
        name: The typeface (``name``, ``rFont`` in a run), or ``None``.
        size: The size in points as Excel takes it, or ``None`` when absent: in the styles'
            fonts, 1 to below 409.6 (another size is refused); in a run, its magnitude
            clamped to 1 to 409.55; in a differential format, its magnitude up to 409.55 (0, or
            a size past it, is refused; captures ``package``, ``attributes`` and ``edges``).
        bold: Bold.
        italic: Italic.
        underline: The underline, or ``None`` when absent (``<u/>`` alone is ``single``).
        strike: Struck through.
        color: The color, or ``None``.
        vert_align: ``superscript``, ``subscript`` or ``baseline``, or ``None``.
        scheme: ``major`` or ``minor`` for a theme font, ``none``, or ``None`` when absent.
        family: The font family number, or ``None``.
        charset: The character set number, or ``None``.
        outline: Outline (a Mac effect).
        shadow: Shadow (a Mac effect).
        condense: Condensed.
        extend: Extended.
        raw_size: The size as written.
        source: The element as written.
    """

    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: Underline | None = None
    strike: bool = False
    color: Color | None = None
    vert_align: VertAlign | None = None
    scheme: FontScheme | None = None
    family: int | None = None
    charset: int | None = None
    outline: bool = False
    shadow: bool = False
    condense: bool = False
    extend: bool = False
    raw_size: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element, *, kind: Literal["cell", "run", "dxf"] = "cell") -> Font:
        """A font; ``kind`` picks Excel's size rule: the styles' fonts, runs, or dxfs."""
        c = _children(element)
        name = c.get("name", c.get("rFont"))
        sz = c.get("sz")
        raw = sz.get("val") if sz is not None else None
        size = None
        if raw is not None:
            size = {"cell": font_size, "run": run_font_size, "dxf": dxf_font_size}[kind](raw)
        if kind == "cell" and "color" in c:
            font_theme(c["color"].get("theme"))
        u = c.get("u")
        va = c.get("vertAlign")
        scheme = c.get("scheme")
        family = c.get("family")
        charset = c.get("charset")
        return cls(
            name.get("val") if name is not None else None,
            size,
            _val_flag(c.get("b")),
            _val_flag(c.get("i")),
            choice(u.get("val", "single"), _UNDERLINES, "single") if u is not None else None,
            _val_flag(c.get("strike")),
            Color.read(c["color"]) if "color" in c else None,
            choice(va.get("val"), _VERT_ALIGNS, "baseline") if va is not None else None,
            choice(scheme.get("val"), _SCHEMES, "none") if scheme is not None else None,
            integer(family.get("val")) if family is not None else None,
            integer(charset.get("val")) if charset is not None else None,
            _val_flag(c.get("outline")),
            _val_flag(c.get("shadow")),
            _val_flag(c.get("condense")),
            _val_flag(c.get("extend")),
            raw,
            element,
        )


@dataclass(frozen=True, slots=True)
class PatternFill:
    """A pattern fill.

    Attributes:
        pattern: The pattern, or ``None`` when the file names none (in a differential format,
            Excel then draws a solid fill of ``background``).
        foreground: The pattern's color (``fgColor``), or ``None``.
        background: The color behind it (``bgColor``), or ``None``.
        source: The element as written.
    """

    pattern: PatternType | None = None
    foreground: Color | None = None
    background: Color | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class GradientStop:
    """A gradient's stop: its position from 0 to 1, and its color."""

    position: float
    color: Color | None


@dataclass(frozen=True, slots=True)
class GradientFill:
    """A gradient fill.

    Attributes:
        type: ``linear`` or ``path``.
        degree: The angle of a linear gradient.
        left: A path gradient's inner rectangle, as fractions of the cell.
        right: See ``left``.
        top: See ``left``.
        bottom: See ``left``.
        stops: The stops, in file order.
        source: The element as written.
    """

    type: Literal["linear", "path"] = "linear"
    degree: float = 0.0
    left: float = 0.0
    right: float = 0.0
    top: float = 0.0
    bottom: float = 0.0
    stops: tuple[GradientStop, ...] = ()
    source: ET.Element | None = field(default=None, compare=False, repr=False)


Fill = PatternFill | GradientFill


def _fill(element: ET.Element) -> Fill | None:
    for child in element:
        name = local(child.tag)
        if name == "patternFill":
            c = _children(child)
            pattern = child.get("patternType")
            return PatternFill(
                choice(pattern, _PATTERNS, "none") if pattern is not None else None,
                Color.read(c["fgColor"]) if "fgColor" in c else None,
                Color.read(c["bgColor"]) if "bgColor" in c else None,
                child,
            )
        if name == "gradientFill":
            stops = []
            for stop in child:
                if local(stop.tag) == "stop":
                    color = next((Color.read(e) for e in stop if local(e.tag) == "color"), None)
                    stops.append(GradientStop(finite(stop.get("position"), 0.0), color))
            get = child.get
            return GradientFill(
                choice(get("type"), ("linear", "path"), "linear"),
                finite(get("degree"), 0.0),
                finite(get("left"), 0.0),
                finite(get("right"), 0.0),
                finite(get("top"), 0.0),
                finite(get("bottom"), 0.0),
                tuple(stops),
                child,
            )
    return None


@dataclass(frozen=True, slots=True)
class BorderSide:
    """One side of a border: its line style and color."""

    style: BorderStyle = "none"
    color: Color | None = None


@dataclass(frozen=True, slots=True)
class Border:
    """A border (``border``).

    Attributes:
        left: The left side, or ``None`` when absent: ``start`` when the border has one, as
            Excel draws it over a ``left`` written after it (capture ``edges``; a ``left``
            written before ``start`` makes Excel refuse the file).
        right: The right side: ``end`` when the border has one, else ``right``; or ``None``.
        top: The top side, or ``None``.
        bottom: The bottom side, or ``None``.
        diagonal: The diagonal line, drawn where ``diagonal_up`` or ``diagonal_down`` says.
        vertical: Inner vertical lines (differential formats), or ``None``.
        horizontal: Inner horizontal lines, or ``None``.
        diagonal_up: The diagonal runs from bottom left to top right.
        diagonal_down: The diagonal runs from top left to bottom right.
        outline: Applied to the outline of a range only.
        source: The element as written.
    """

    left: BorderSide | None = None
    right: BorderSide | None = None
    top: BorderSide | None = None
    bottom: BorderSide | None = None
    diagonal: BorderSide | None = None
    vertical: BorderSide | None = None
    horizontal: BorderSide | None = None
    diagonal_up: bool = False
    diagonal_down: bool = False
    outline: bool = True
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element) -> Border:
        """Raises ``PackageError`` for a ``left`` before ``start`` or a ``right`` before ``end``."""
        border_order([local(e.tag) for e in element])
        c = _children(element)

        def side(*names: str) -> BorderSide | None:
            found = next((c[n] for n in names if n in c), None)
            if found is None:
                return None
            color = next((Color.read(e) for e in found if local(e.tag) == "color"), None)
            return BorderSide(choice(found.get("style"), _BORDERS, "none"), color)

        get = element.get
        return cls(
            side("start", "left"),
            side("end", "right"),
            side("top"),
            side("bottom"),
            side("diagonal"),
            side("vertical"),
            side("horizontal"),
            flag(get("diagonalUp"), False),
            flag(get("diagonalDown"), False),
            flag(get("outline"), True),
            element,
        )


@dataclass(frozen=True, slots=True)
class Alignment:
    """A cell format's alignment.

    Attributes:
        horizontal: The horizontal alignment, or ``None`` when absent (``general``).
        vertical: The vertical alignment, or ``None`` when absent (``bottom``).
        wrap_text: Text wraps.
        shrink_to_fit: Text shrinks to fit.
        indent: The indent level.
        text_rotation: 0 to 90 counterclockwise, 91 to 180 clockwise (as 90 minus), 255
            vertical text; 181 to 254 are kept as written and anything else reads as 0, as
            Excel keeps them (capture ``attributes``; what it draws for 181 to 254 is not
            captured).
        reading_order: 0 by context, 1 left to right, 2 right to left.
        justify_last_line: The last line of distributed text is justified.
        relative_indent: An indent added in differential formats.
        source: The element as written.
    """

    horizontal: Horizontal | None = None
    vertical: Vertical | None = None
    wrap_text: bool = False
    shrink_to_fit: bool = False
    indent: int = 0
    text_rotation: int = 0
    reading_order: int = 0
    justify_last_line: bool = False
    relative_indent: int = 0
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    @classmethod
    def read(cls, element: ET.Element) -> Alignment:
        get = element.get
        h, v = get("horizontal"), get("vertical")
        return cls(
            choice(h, _HORIZONTALS, "general") if h is not None else None,
            choice(v, _VERTICALS, "bottom") if v is not None else None,
            flag(get("wrapText"), False),
            flag(get("shrinkToFit"), False),
            unsigned(get("indent"), 0),
            _rotation(get("textRotation")),
            unsigned(get("readingOrder"), 0),
            flag(get("justifyLastLine"), False),
            integer(get("relativeIndent"), 0) or 0,
            element,
        )


@dataclass(frozen=True, slots=True)
class Protection:
    """A cell format's protection: locked and hidden apply when the sheet is protected."""

    locked: bool = True
    hidden: bool = False


@dataclass(frozen=True, slots=True)
class CellFormat:
    """A cell format (``cellXfs/xf``) or a named style's format (``cellStyleXfs/xf``).

    Attributes:
        number_format_id: The number format's id.
        number_format: The format code of a custom number format, or ``None`` for a built-in id
            (whose code depends on the locale).
        font: The font, or ``None`` when the index leads nowhere.
        fill: The fill, or ``None``.
        border: The border, or ``None``.
        alignment: The alignment, or ``None`` when absent.
        protection: The protection, or ``None`` when absent.
        style: For a cell format, the named style's format it is based on (``xfId``), or
            ``None``.
        font_id: The font's index.
        fill_id: The fill's index.
        border_id: The border's index.
        xf_id: The named style's format index, or ``None``.
        apply_number_format: ``applyNumberFormat`` as written, ``None`` when absent; Excel's
            reading of the apply flags differs between cell and style formats.
        apply_font: ``applyFont``.
        apply_fill: ``applyFill``.
        apply_border: ``applyBorder``.
        apply_alignment: ``applyAlignment``.
        apply_protection: ``applyProtection``.
        quote_prefix: Text entered with a leading apostrophe.
        pivot_button: Shows a pivot table button.
        source: The element as written.
    """

    number_format_id: int = 0
    number_format: str | None = None
    font: Font | None = None
    fill: Fill | None = None
    border: Border | None = None
    alignment: Alignment | None = None
    protection: Protection | None = None
    style: CellFormat | None = None
    font_id: int = 0
    fill_id: int = 0
    border_id: int = 0
    xf_id: int | None = None
    apply_number_format: bool | None = None
    apply_font: bool | None = None
    apply_fill: bool | None = None
    apply_border: bool | None = None
    apply_alignment: bool | None = None
    apply_protection: bool | None = None
    quote_prefix: bool = False
    pivot_button: bool = False
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class NamedStyle:
    """A named cell style (``cellStyles/cellStyle``), such as Normal or Hyperlink.

    Attributes:
        name: Its name.
        xf_id: The index of its format in the style formats.
        format: That format, or ``None`` when the index leads nowhere.
        builtin_id: The built-in style's id (0 is Normal), or ``None`` for a custom style.
        hidden: Hidden from the styles gallery.
        custom_builtin: A built-in style changed by the user.
    """

    name: str
    xf_id: int
    format: CellFormat | None
    builtin_id: int | None = None
    hidden: bool = False
    custom_builtin: bool = False


@dataclass(frozen=True, slots=True)
class DifferentialFormat:
    """A differential format (``dxfs/dxf``), as conditional formats and table styles use it.

    It holds the parts set over a cell's own format; ``None`` for each part it leaves alone.

    Attributes:
        font: The font properties it sets (its size by the differential-format rule of
            ``Font.size``).
        number_format_id: The number format's id, or ``None``.
        number_format: The number format's code, or ``None``.
        fill: The fill.
        alignment: The alignment.
        border: The border.
        protection: The protection.
        source: The element as written.
    """

    font: Font | None = None
    number_format_id: int | None = None
    number_format: str | None = None
    fill: Fill | None = None
    alignment: Alignment | None = None
    border: Border | None = None
    protection: Protection | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class Styles:
    """The styles part, its indices resolved.

    Attributes:
        number_formats: Custom number formats by id; built-in ids are not listed.
        fonts: The fonts.
        fills: The fills (``None`` for a ``fill`` with neither pattern nor gradient).
        borders: The borders.
        cell_formats: The cell formats a cell's ``style`` indexes (``cellXfs``).
        cell_style_formats: The named styles' formats (``cellStyleXfs``).
        cell_styles: The named styles.
        differential_formats: The differential formats (``dxfs``).
        table_styles: The workbook's own table styles and its default style names.
        indexed_colors: The workbook's own palette (``colors/indexedColors``) as ``AARRGGBB``
            in upper case, or ``None`` when the file keeps Excel's default palette. An entry
            that is not eight hex digits is ``None``.
        mru_colors: The recently used colors (``colors/mruColors``).
        part: The styles part's name, or ``None`` when the workbook has none.
        unparsed: The styles part's elements xlio does not model (``extLst``, for one).
        source: The whole part, or ``None``.
    """

    number_formats: Mapping[int, str] = field(default_factory=dict)
    fonts: tuple[Font, ...] = ()
    fills: tuple[Fill | None, ...] = ()
    borders: tuple[Border, ...] = ()
    cell_formats: tuple[CellFormat, ...] = ()
    cell_style_formats: tuple[CellFormat, ...] = ()
    cell_styles: tuple[NamedStyle, ...] = ()
    differential_formats: tuple[DifferentialFormat, ...] = ()
    table_styles: TableStyles = field(default_factory=TableStyles)
    indexed_colors: tuple[str | None, ...] | None = None
    mru_colors: tuple[Color, ...] = ()
    part: str | None = None
    unparsed: tuple[Unparsed, ...] = ()
    source: ET.Element | None = field(default=None, compare=False, repr=False)

    def cell_format(self, index: int) -> CellFormat | None:
        """The cell format a cell's ``style`` names, or ``None`` when it leads nowhere."""
        return self.cell_formats[index] if 0 <= index < len(self.cell_formats) else None

    def table_style(self, name: str) -> TableStyle | None:
        """A table style by name: the workbook's own first, then Excel's presets."""
        own = next((s for s in self.table_styles.styles if s.name == name), None)
        if own is not None:
            return own
        from xlio import presets

        return presets.table_style(name)


_MODELED = frozenset(
    {"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
     "tableStyles", "colors"}
)  # fmt: skip


def _at(items: Sequence[_T], index: int) -> _T | None:
    return items[index] if 0 <= index < len(items) else None


def _apply(text: str | None) -> bool | None:
    return None if text is None else flag(text, True)


def _protection(element: ET.Element | None) -> Protection | None:
    if element is None:
        return None
    return Protection(flag(element.get("locked"), True), flag(element.get("hidden"), False))


def _xfs(
    element: ET.Element,
    formats: Mapping[int, str],
    fonts: Sequence[Font],
    fills: Sequence[Fill | None],
    borders: Sequence[Border],
    styles: Sequence[CellFormat] | None,
) -> tuple[CellFormat, ...]:
    out = []
    for xf in element:
        if local(xf.tag) != "xf":
            continue
        get = xf.get
        c = _children(xf)
        number_format_id = unsigned(get("numFmtId"), 0)
        font_id, fill_id, border_id = (
            unsigned(get("fontId"), 0),
            unsigned(get("fillId"), 0),
            unsigned(get("borderId"), 0),
        )
        xf_id = integer(get("xfId")) if styles is not None else None
        out.append(
            CellFormat(
                number_format_id,
                formats.get(number_format_id),
                _at(fonts, font_id),
                _at(fills, fill_id),
                _at(borders, border_id),
                Alignment.read(c["alignment"]) if "alignment" in c else None,
                _protection(c.get("protection")),
                _at(styles, xf_id) if styles is not None and xf_id is not None else None,
                font_id,
                fill_id,
                border_id,
                xf_id,
                _apply(get("applyNumberFormat")),
                _apply(get("applyFont")),
                _apply(get("applyFill")),
                _apply(get("applyBorder")),
                _apply(get("applyAlignment")),
                _apply(get("applyProtection")),
                flag(get("quotePrefix"), False),
                flag(get("pivotButton"), False),
                xf,
            )
        )
    return tuple(out)


def _dxf(element: ET.Element) -> DifferentialFormat:
    c = _children(element)
    num = c.get("numFmt")
    return DifferentialFormat(
        Font.read(c["font"], kind="dxf") if "font" in c else None,
        unsigned(num.get("numFmtId"), 0) if num is not None else None,
        num.get("formatCode") if num is not None else None,
        _fill(c["fill"]) if "fill" in c else None,
        Alignment.read(c["alignment"]) if "alignment" in c else None,
        Border.read(c["border"]) if "border" in c else None,
        _protection(c.get("protection")),
        element,
    )


def read_differential_formats(element: ET.Element) -> tuple[DifferentialFormat, ...]:
    """A ``dxfs`` element's formats."""
    return tuple(_dxf(d) for d in element if local(d.tag) == "dxf")


def border_order(names: list[str]) -> None:
    """Refuse a border with ``left`` before ``start``, or ``right`` before ``end``.

    Excel refuses such a file (capture ``edges``).

    Raises:
        PackageError: Such a border.
    """
    for before, after in (("left", "start"), ("right", "end")):
        if before in names and after in names and names.index(before) < names.index(after):
            raise PackageError(f"a border with {before} before {after}, which Excel refuses")


def check_font_sizes(stream: BinaryIO, part: str) -> None:
    """Refuse a styles part whose fonts hold a size Excel refuses, without building the styles.

    The fonts of ``fonts`` (size and theme color) and of differential formats (size) are
    checked, by the rules the captures settled, and so is the order of each border's sides.

    Raises:
        PackageError: A font size or font color Excel refuses, or a part that is not
            well-formed.
    """
    path: list[str] = []
    sides: list[str] = []

    def start(name: str, attrs: dict[str, str]) -> None:
        path.append(name.rpartition("}")[2])
        if len(path) == 4 and path[1:3] == ["borders", "border"]:
            sides.append(path[3])
            border_order(sides)
        elif len(path) == 3 and path[1:] == ["borders", "border"]:
            sides.clear()
        elif path[-3:] == ["fonts", "font", "sz"] and len(path) == 4 and "val" in attrs:
            font_size(attrs["val"])
        elif path[-4:] == ["dxfs", "dxf", "font", "sz"] and len(path) == 5 and "val" in attrs:
            dxf_font_size(attrs["val"])
        elif path[-3:] == ["fonts", "font", "color"] and len(path) == 4:
            font_theme(attrs.get("theme"))

    def end(_: str) -> None:
        path.pop()

    p = xml.parsers.expat.ParserCreate(namespace_separator="}")
    p.StartElementHandler, p.EndElementHandler = start, end
    try:
        while chunk := stream.read(1 << 16):
            p.Parse(chunk, False)
        p.Parse(b"", True)
    except xml.parsers.expat.ExpatError as exc:
        raise PackageError(f"{part}: {exc}") from exc


class _CountedError(Exception):
    """Raised to stop the count at the end of ``cellXfs``."""


def count_cell_formats(stream: BinaryIO, part: str) -> int:
    """The number of cell formats :func:`read_styles` would read, without reading them.

    The part is streamed up to the end of its first ``cellXfs``, whose ``xf`` children are
    counted, as :func:`read_styles` takes them.

    Raises:
        PackageError: The part, as far as it is read, is not well-formed.
    """
    depth = count = 0
    state = 0  # 0 before cellXfs, 1 inside it

    def start(name: str, _: dict[str, str]) -> None:
        nonlocal depth, count, state
        depth += 1
        if depth == 2 and state == 0 and name.rpartition("}")[2] == "cellXfs":
            state = 1
        elif depth == 3 and state == 1 and name.rpartition("}")[2] == "xf":
            count += 1

    def end(_: str) -> None:
        nonlocal depth
        depth -= 1
        if depth == 1 and state == 1:
            raise _CountedError

    p = xml.parsers.expat.ParserCreate(namespace_separator="}")
    p.StartElementHandler, p.EndElementHandler = start, end
    try:
        while chunk := stream.read(1 << 16):
            p.Parse(chunk, False)
        p.Parse(b"", True)
    except _CountedError:
        pass
    except xml.parsers.expat.ExpatError as exc:
        raise PackageError(f"{part}: {exc}") from exc
    return count


def read_styles(root: ET.Element, part: str) -> Styles:
    """The styles from their part's element tree.

    Raises:
        PackageError: A font size Excel refuses.
    """
    c = _children(root)
    formats: dict[int, str] = {}
    if "numFmts" in c:
        for fmt in c["numFmts"]:
            number_id = integer(fmt.get("numFmtId"))
            if number_id is not None and fmt.get("formatCode") is not None:
                formats[number_id] = fmt.get("formatCode", "")
    fonts = tuple(Font.read(f) for f in c["fonts"]) if "fonts" in c else ()
    fills = tuple(_fill(f) for f in c["fills"]) if "fills" in c else ()
    borders = tuple(Border.read(b) for b in c["borders"]) if "borders" in c else ()
    style_formats: tuple[CellFormat, ...] = ()
    if "cellStyleXfs" in c:
        style_formats = _xfs(c["cellStyleXfs"], formats, fonts, fills, borders, None)
    cell_formats: tuple[CellFormat, ...] = ()
    if "cellXfs" in c:
        cell_formats = _xfs(c["cellXfs"], formats, fonts, fills, borders, style_formats)
    named = []
    for style in c.get("cellStyles", ()):
        xf_id = unsigned(style.get("xfId"), 0)
        named.append(
            NamedStyle(
                style.get("name", ""),
                xf_id,
                _at(style_formats, xf_id),
                integer(style.get("builtinId")),
                flag(style.get("hidden"), False),
                flag(style.get("customBuiltin"), False),
            )
        )
    dxfs = read_differential_formats(c["dxfs"]) if "dxfs" in c else ()
    table_styles = (
        read_table_styles(c["tableStyles"], dxfs) if "tableStyles" in c else TableStyles()
    )
    indexed = None
    mru: list[Color] = []
    for group in c.get("colors", ()):
        name = local(group.tag)
        if name == "indexedColors":
            values = []
            for entry in group:
                rgb = entry.get("rgb", "").strip()
                values.append(rgb.upper() if _RGB.fullmatch(rgb) else None)
            indexed = tuple(values)
        elif name == "mruColors":
            mru.extend(Color.read(e) for e in group)
    unparsed = tuple(
        Unparsed(f"styleSheet/{local(e.tag)}", e) for e in root if local(e.tag) not in _MODELED
    )
    return Styles(
        formats,
        fonts,
        fills,
        borders,
        cell_formats,
        style_formats,
        tuple(named),
        dxfs,
        table_styles,
        indexed,
        tuple(mru),
        part,
        unparsed,
        root,
    )
