"""Streaming reads of shared strings and worksheet cells with expat.

Element names are matched by their local part (expat runs without namespace processing, the
fast path), so prefixed and Strict OOXML files read like transitional ones. Text is
unescaped from OOXML's ``_xHHHH_`` form.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import xml.parsers.expat
from collections.abc import Callable, Iterator, Mapping
from typing import BinaryIO, Literal, NamedTuple, cast

from xlio._package import PackageError
from xlio._rich import RichString, rich
from xlio._types import bounded_int

__all__ = [
    "ROW_ATTRIBUTES",
    "Cell",
    "CellType",
    "Formula",
    "read_cells",
    "read_shared_strings",
    "unescape",
]

_CHUNK = 1 << 16
_SPACE = " \t\n\r"
_TRUE = ("1", "true")
# The error codes Excel reads written as text in a <v>; any other makes it refuse the file, as a
# constant or as a formula's cached value (capture errors).
_PLAIN_ERRORS = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA",
     "#BUSY!"}
)  # fmt: skip
_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")
_SURROGATE = re.compile("[\ud800-\udfff]")
_MAX_ROW = 1_048_576
_MAX_COL = 16_384
_INF = float("inf")
_DBL_MIN = 2.2250738585072014e-308  # the smallest normal double


def unescape(text: str) -> str:
    """OOXML's escape of characters XML cannot hold: ``_x000D_`` is a carriage return.

    A high surrogate escaped just before a low one (``_xD83D__xDE00_``, in either case) is one
    character, as UTF-16 holds it (U+1F600). A surrogate without its other half, or after it,
    is one character of its own, a lone surrogate that no encoding writes: Excel holds it so
    and writes it back escaped (captures ``strings`` and ``strings_cached``).
    """
    if "_x" not in text:
        return text
    text = _ESCAPE.sub(lambda m: chr(int(m[1], 16)), text)
    if _SURROGATE.search(text) is None:
        return text
    return text.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")


_LOCAL = {"row": 1, "c": 2, "v": 3, "f": 4, "is": 5, "t": 6, "rPh": 7, "si": 8, "r": 9}


def _codes() -> dict[str, int]:
    """Element name → code, filled per name on first sight (prefixed names included)."""

    class Codes(dict[str, int]):
        def __missing__(self, name: str) -> int:
            code = _LOCAL.get(name.rpartition(":")[2], 0)
            self[name] = code
            return code

    return Codes()


def read_shared_strings(stream: BinaryIO) -> list[str]:
    """The shared string table: each ``<si>``'s text, phonetic runs (``<rPh>``) left out.

    Each ``<t>`` not marked ``xml:space="preserve"`` loses its leading and trailing whitespace,
    as Excel reads it (captured by xlrecalc:
    https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/whitespace.tsv).

    Raises:
        PackageError: A ``<t>`` after a run (``<r>``) of the same string, which makes Excel
            refuse the file (capture ``edges``).
    """
    strings: list[str] = []
    parts: list[str] = []
    codes = _codes()
    in_t = False
    t_start = 0
    preserve = False
    phonetic = in_run = 0
    ran = False  # the string has had a run

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal in_t, t_start, preserve, phonetic, in_run, ran
        code = codes[name]
        if code == 6:
            if ran and not in_run and not phonetic:
                raise PackageError("a string with text after its runs, which Excel refuses")
            in_t = True
            t_start = len(parts)
            preserve = attrs.get("xml:space") == "preserve"
        elif code == 7:
            phonetic += 1
        elif code == 8:
            parts.clear()
            ran = False
        elif code == 9:
            in_run += 1
            ran = True

    def end(name: str) -> None:
        nonlocal in_t, phonetic, in_run
        code = codes[name]
        if code == 6:
            in_t = False
            if not preserve and not phonetic:
                parts[t_start:] = ["".join(parts[t_start:]).strip(_SPACE)]
        elif code == 7:
            phonetic -= 1
        elif code == 8:
            strings.append(unescape("".join(parts)))
        elif code == 9:
            in_run -= 1

    def text(data: str) -> None:
        if in_t and not phonetic:
            parts.append(data)

    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    while chunk := stream.read(_CHUNK):
        p.Parse(chunk, False)
    p.Parse(b"", True)
    return strings


class Formula(NamedTuple):
    """A cell's formula (``<f>``).

    Attributes:
        text: The formula's text as the file stores it (no leading ``=``, ``_xHHHH_`` escapes
            decoded), or ``None`` for the dependent cells of a shared formula.
        kind: ``normal``, ``shared``, ``array`` or ``dataTable``, as written.
        ref: The range of a shared formula's master or of an array formula, or ``None``.
        si: The shared-formula index, or ``None``.
        attributes: Every attribute of ``<f>`` as written.
    """

    text: str | None
    kind: str
    ref: str | None
    si: int | None
    attributes: dict[str, str]


CellType = Literal["n", "s", "str", "inlineStr", "b", "e", "d", "none"]


class Cell(NamedTuple):
    """One cell of a worksheet, as Excel reads it.

    Attributes:
        row: The row, 1-based.
        column: The column, 1-based.
        value: The value Excel reads: a ``float`` for a number (a subnormal reads as 0), a
            ``str`` for text, a ``bool``, an :class:`ErrorCode` for an error (a code Excel
            refuses is refused), the ISO 8601 text of a ``d`` cell, or ``None``.
        type: Where the value comes from: ``n`` (a number), ``s`` (the shared strings), ``str``
            (a formula's text result), ``inlineStr``, ``b`` (a boolean), ``e`` (an error), ``d``
            (an ISO 8601 date) or ``none`` (no value: an uncached formula or a styled empty
            cell).
        formula: The cell's formula, or ``None``.
        style: The index of the cell's format in the workbook's cell formats (``s``); 0 when
            the file gives none or one that is not a non-negative integer.
        cm: The cell-metadata index (``cm``), or ``None``.
        vm: The value-metadata index (``vm``), or ``None``.
        phonetic: Phonetic guides shown (``ph``).
        rich: For a shared or inline string read with ``rich=True``: its runs and phonetic
            guides; otherwise ``None``.
    """

    row: int
    column: int
    value: float | str | bool | None
    type: CellType
    formula: Formula | None
    style: int = 0
    cm: int | None = None
    vm: int | None = None
    phonetic: bool = False
    rich: RichString | None = None


def _column(ref: str) -> tuple[int, int]:
    col = 0
    for i, ch in enumerate(ref):
        o = ord(ch)
        if 65 <= o <= 90:
            col = col * 26 + o - 64
        elif 97 <= o <= 122:
            col = col * 26 + o - 96
        else:
            return int(ref[i:]), col
    raise ValueError(f"not a cell reference: {ref!r}")


def _row(text: str) -> int:
    """A row number from the file; Excel refuses one off its grid (capture ``value_forms``)."""
    try:
        row = int(text)
    except ValueError:
        raise PackageError(f"not a row number: {text!r}") from None
    if not 1 <= row <= _MAX_ROW:
        raise PackageError(f"row {row} is off the sheet")
    return row


def _number(text: str, row: int, col: int) -> tuple[float | str, str]:
    """A ``<v>`` of a number cell as Excel reads it (capture ``value_forms``).

    Python's ``float`` takes more than Excel: ``nan``, ``inf``, ``+5`` and ``1_000``, which make
    Excel refuse the file, are refused here too. A number too large for a double is ``#NULL!``,
    as Excel reads it, and a subnormal is 0.
    """
    try:
        value = float(text)
    except ValueError:
        raise PackageError(f"row {row}, column {col}: not a number: {text[:40]!r}") from None
    if -_INF < value < _INF and "_" not in text and text.lstrip()[:1] != "+":
        # Leading whitespace is allowed; a subnormal reads as 0 (capture value_forms).
        return (0.0 if -_DBL_MIN < value < _DBL_MIN else value), "n"
    if (
        "_" in text
        or text.lstrip()[:1] == "+"
        or any(ch.isalpha() and ch not in "eE" for ch in text)
    ):
        raise PackageError(f"row {row}, column {col}: not a number Excel reads: {text[:40]!r}")
    return "#NULL!", "e"


ROW_ATTRIBUTES = frozenset(
    {"ht", "customHeight", "hidden", "s", "customFormat", "outlineLevel", "collapsed",
     "thickTop", "thickBot", "ph"}
)  # fmt: skip
"""The ``<row>`` attributes that make a row format; ``r``, ``spans`` and the rest do not."""


def read_cells(
    stream: BinaryIO,
    strings: list[str],
    rich_strings: list[RichString] | None = None,
    *,
    on_row: Callable[[int, dict[str, str]], None] | None = None,
    errors: Mapping[str, str] | None = None,
    value_errors: Mapping[int, str | float] | None = None,
    bounds: tuple[int, int, int] | None = None,
) -> Iterator[Cell]:
    """The cells of a worksheet part, in file order.

    Args:
        stream: The part.
        strings: The shared strings.
        rich_strings: The shared strings with their runs, to give string cells their runs.
        on_row: Called with each ``<row>`` that has an attribute of :data:`ROW_ATTRIBUTES`
            (its 1-based number and its attributes), and with each that has none and holds no
            cell (its number and no attributes: Excel counts the row), so row formats are
            gathered in the same pass.
        errors: Error texts mapped to the values given (a text not in it is given as written).
        value_errors: Value-metadata indexes mapped to the error a rich value stands for; an
            error cell with such a ``vm`` is given that error, or that number (a ``float``).
        bounds: The number of cell formats, cell-metadata and value-metadata entries: a cell
            whose ``s``, ``cm`` or ``vm`` is past its table is refused, as Excel refuses it
            (capture ``attributes``).

    The element handlers are the hot path of a load: expat runs without namespace processing
    (names are matched by their local part, so a prefixed file reads the same), and the state
    lives in closure variables. With ``rich_strings``, string cells carry their runs (inline
    strings' are built from their element).
    """
    builder: ET.TreeBuilder | None = None
    inline_rich: RichString | None = None
    out: list[Cell] = []
    buffer: list[str] = []
    codes = _codes()
    row = col = 0
    attrs_c: dict[str, str] = {}
    v: str | None = None
    v_preserve = False
    t_preserve = False
    f_text: str | None = None
    f_attrs: dict[str, str] | None = None
    inline: list[str] | None = None
    collect = 0  # 3 <v>, 4 <f>, 6 <t> inside <is>
    empty_row = 0  # a <row> without format attributes, until a cell shows it holds one
    in_is = phonetic = in_run = 0
    ran = False  # the inline string has had a run

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal row, col, attrs_c, v, v_preserve, t_preserve, f_text, f_attrs, inline, empty_row
        nonlocal collect, in_is, phonetic, builder, inline_rich, in_run, ran
        code = codes[name]
        if builder is not None:
            builder.start(name, attrs)
        if code == 2:
            empty_row = 0
            ref = attrs.get("r")
            if ref:
                try:
                    ref_row, ref_col = _column(ref)
                except ValueError:
                    raise PackageError(f"not a cell reference: {ref[:40]!r}") from None
                if not 1 <= ref_row <= _MAX_ROW:
                    raise PackageError(f"row {ref_row} is off the sheet")
                row = ref_row
                # Excel places a cell whose column is off its grid by position (capture
                # value_forms: XFE1 after nothing is A1).
                col = ref_col if 1 <= ref_col <= _MAX_COL else col + 1
            else:
                col += 1
            if col > _MAX_COL:
                raise PackageError(f"row {row}: more cells than columns")
            attrs_c = attrs
            v = f_text = None
            v_preserve = False
            f_attrs = inline = inline_rich = None
        elif code == 3:
            buffer.clear()
            collect = 3
            v_preserve = attrs.get("xml:space") == "preserve"
        elif code == 4:
            buffer.clear()
            collect = 4
            f_attrs = attrs
        elif code == 1:
            r = attrs.get("r")
            row = _row(r) if r else row + 1
            if row > _MAX_ROW:
                raise PackageError(f"row {row} is off the sheet")
            col = 0
            if on_row is not None:
                if ROW_ATTRIBUTES.isdisjoint(attrs):
                    empty_row = row  # reported at its end if it holds no cell
                else:
                    on_row(row, attrs)
        elif code == 5:
            in_is += 1
            ran = False
            inline = []
            if rich_strings is not None:
                builder = ET.TreeBuilder()
                builder.start(name, attrs)
        elif code == 7:
            phonetic += 1
        elif code == 9 and in_is:
            in_run += 1
            ran = True
        elif code == 6 and in_is:
            if ran and not in_run and not phonetic:
                raise PackageError(
                    f"row {row}: a string with text after its runs, which Excel refuses"
                )
            buffer.clear()
            collect = 6
            t_preserve = attrs.get("xml:space") == "preserve"

    def finish(cell: Cell) -> Cell:
        """A cell with its runs, or the error its rich value stands for: the rarer cells."""
        if rich_strings is not None and cell.type == "s":
            index = int(v or "-1")  # checked by _cell
            if index < len(rich_strings):
                cell = cell._replace(rich=rich_strings[index])
        elif inline_rich is not None and cell.type == "inlineStr":
            cell = cell._replace(rich=inline_rich)
        # An error rich value behind vm wins, whatever <v> holds: Excel shows it over a
        # cached 5, TRUE, text or other error when it does not recalculate (capture
        # errors_stale). Other rich values (pictures, data types) leave the cell as it is.
        rich_value = None
        if value_errors is not None and cell.vm is not None:
            rich_value = value_errors.get(cell.vm)
        if isinstance(rich_value, float):
            cell = cell._replace(value=rich_value, type="n", rich=None)
        elif rich_value is not None:
            error = errors.get(rich_value, rich_value) if errors is not None else rich_value
            cell = cell._replace(value=error, type="e", rich=None)
        elif cell.type == "e":
            text = cast(str, cell.value)
            cell = cell._replace(value=errors.get(text, text) if errors is not None else text)
        return cell

    def end(name: str) -> None:
        nonlocal v, f_text, collect, in_is, phonetic, builder, inline_rich, empty_row, in_run
        code = codes[name]
        if builder is not None:
            builder.end(name)
            if code == 5:
                inline_rich = rich(builder.close())
                builder = None
        if code == 2:
            cell = _cell(row, col, attrs_c, v, v_preserve, f_text, f_attrs, inline, strings, bounds)
            if (
                rich_strings is not None
                or inline_rich is not None
                or cell.vm is not None
                or cell.type == "e"
            ):  # the hot path skips what most cells do not need
                cell = finish(cell)  # leaves style, cm and vm, which _cell checked
            out.append(cell)
        elif code == 3:
            v = "".join(buffer)
            collect = 0
        elif code == 4:
            f_text = "".join(buffer)
            collect = 0
        elif code == 6 and collect == 6:
            if not phonetic and inline is not None:
                run = "".join(buffer)
                inline.append(run if t_preserve else run.strip(_SPACE))
            collect = 0
        elif code == 7:
            phonetic -= 1
        elif code == 5:
            in_is -= 1
        elif code == 9 and in_is:
            in_run -= 1
        elif code == 1 and empty_row:
            on_row(empty_row, {})  # type: ignore[misc]  # set only with on_row
            empty_row = 0

    def text(data: str) -> None:
        if collect:
            buffer.append(data)
        if builder is not None:
            builder.data(data)

    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    while chunk := stream.read(_CHUNK):
        p.Parse(chunk, False)
        yield from out
        out.clear()
    p.Parse(b"", True)
    yield from out


def _cell(
    row: int,
    col: int,
    attrs: dict[str, str],
    v: str | None,
    v_preserve: bool,
    f_text: str | None,
    f_attrs: dict[str, str] | None,
    inline: list[str] | None,
    strings: list[str],
    bounds: tuple[int, int, int] | None = None,
) -> Cell:
    t = attrs.get("t", "n")
    value: float | str | bool | None
    if t == "n":
        if v:
            value, t = _number(v, row, col)
        else:
            value, t = None, "none"
    elif t == "inlineStr":
        value = unescape("".join(inline)) if inline is not None else None
    elif v is None or (v == "" and t != "str"):
        value, t = None, "none"
    elif t == "s":
        try:
            index = int(v)
        except ValueError:
            index = -1
        if not 0 <= index < len(strings):
            raise PackageError(f"row {row}, column {col}: no shared string {v[:40]!r}")
        value = strings[index]
    elif t == "str":
        # Excel drops leading and trailing whitespace of a <v> not marked xml:space="preserve"
        value = unescape(v if v_preserve else v.strip(_SPACE))
    elif t == "b":
        value = v in ("1", "true")
    elif t == "e":
        if v not in _PLAIN_ERRORS:
            raise PackageError(f"row {row}, column {col}: an error Excel refuses: {v[:40]!r}")
        value = v
    else:  # "d"
        value = v
    formula = None
    if f_attrs is not None:
        si = f_attrs.get("si")
        formula = Formula(
            unescape(f_text) if f_text else None,  # Excel decodes escapes in formulas too
            f_attrs.get("t", "normal"),
            f_attrs.get("ref"),
            _shared_index(si, row, col) if si is not None else None,
            f_attrs,
        )
    # The hot path: most cells have no cm, vm or ph, so _index is called only for those written,
    # and a style index too short to need a bound is read directly.
    s, cm, vm = attrs.get("s"), attrs.get("cm"), attrs.get("vm")
    if s is None:
        style = 0
    else:
        style = _STYLES.get(s, -1)
        if style < 0:
            style = _index(s, None) or 0
            if len(_STYLES) < _STYLES_KEPT:
                _STYLES[s] = style
    cm_index = _index(cm, None) if cm is not None else None
    vm_index = _index(vm, None) if vm is not None else None
    if bounds is not None and (
        (style and style >= bounds[0])
        or (cm_index is not None and cm_index > bounds[1])
        or (vm_index is not None and vm_index > bounds[2])
    ):
        raise PackageError(f"row {row}, column {col}: a format or metadata index past its table")
    return Cell(
        row,
        col,
        value,
        t,  # type: ignore[arg-type]  # one of CellType, by the branches above
        formula,
        style,
        cm_index,
        vm_index,
        attrs.get("ph") in _TRUE,
    )


_STYLES: dict[str, int] = {}
"""An ``s`` attribute's text -> the style index read from it: a sheet repeats a few texts."""
_STYLES_KEPT = 65_536  # the cache stops growing there, whatever a file holds


def _index(text: str | None, default: int | None) -> int | None:
    """A 0- or 1-based index attribute; one that is not a non-negative integer is the default."""
    if text is None:
        return default
    if text.isdigit() and text.isascii():
        return bounded_int(text)
    return default


def _shared_index(text: str, row: int, col: int) -> int:
    try:
        return int(text)
    except ValueError:
        raise PackageError(
            f"row {row}, column {col}: not a shared formula index: {text[:40]!r}"
        ) from None
