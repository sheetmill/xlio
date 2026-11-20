"""A worksheet's formula cells only, streamed without reading any other cell.

Two readers give the same items. The scan finds each ``<f>`` and the ``<c>`` that holds it in
the part's bytes, while expat, given no handlers, checks that the part is well-formed. It takes
only parts written plainly, as Excel writes them: UTF-8; among the cells, unprefixed names and
no comment, CDATA section or processing instruction; no cell after ``sheetData``; every formula
cell with its ``r``; no cell longer than 1 MiB. On any other part it stops, and the caller reads
the part again with the second reader, expat with handlers, skipping the items already given.
"""

from __future__ import annotations

import re
import xml.parsers.expat
from collections.abc import Iterator
from typing import BinaryIO, NamedTuple

from xlio._cells import _MAX_COL, _MAX_ROW, Formula, _codes, _column, _index, _row, unescape
from xlio._package import PackageError

__all__ = ["FormulaCell", "NotPlainError", "read_formulas", "scan_formulas"]

_CHUNK = 1 << 16
_TAIL = 1 << 20  # a cell longer than this is read by expat: the scan would copy it per chunk


class FormulaCell(NamedTuple):
    """A cell that holds a formula, as :meth:`Sheet.formulas` gives it.

    Attributes:
        row: The row, 1-based.
        column: The column, 1-based.
        formula: The formula, as :attr:`Cell.formula` gives it: the dependent cells of a shared
            formula carry only its ``si`` (``text`` is ``None``).
        cm: The cell-metadata index (``cm``), or ``None``.
        cached: The cell holds a cached value: :attr:`Cell.value` would not be ``None``.
    """

    row: int
    column: int
    formula: Formula
    cm: int | None
    cached: bool


class NotPlainError(Exception):
    """The part is not written plainly enough for the scan: read it with expat instead."""


def _formula(text: str | None, attrs: dict[str, str], row: int, col: int) -> Formula:
    si = attrs.get("si")
    if si is not None:
        try:
            index: int | None = int(si)
        except ValueError:
            raise PackageError(
                f"row {row}, column {col}: not a shared formula index: {si[:40]!r}"
            ) from None
    else:
        index = None
    return Formula(
        unescape(text) if text else None, attrs.get("t", "normal"), attrs.get("ref"), index, attrs
    )


def _cm(text: str | None, row: int, col: int, cms: int | None) -> int | None:
    cm = _index(text, None) if text is not None else None
    if cm is not None and cms is not None and cm > cms:
        raise PackageError(f"row {row}, column {col}: a format or metadata index past its table")
    return cm


def _position(ref: str) -> tuple[int, int]:
    try:
        row, col = _column(ref)
    except ValueError:
        raise PackageError(f"not a cell reference: {ref[:40]!r}") from None
    if not 1 <= row <= _MAX_ROW:
        raise PackageError(f"row {row} is off the sheet")
    return row, col


# The scan -------------------------------------------------------------------------------------

_F = re.compile(rb"<f(?=[\s/>])([^>]*?)(/>|>([^<]*)</f>)")
_F_START = re.compile(rb"<f[\s/>]")
_C_START = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?c[\s/>]")
_SPECIAL = re.compile(rb"<(?:([!?])|/sheetData>|[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*))")
"""Among the cells, what the scan does not take: a declaration, or a prefixed name; and their
end. Group 1 is ``!`` or ``?``; group 2, a prefixed element's local name."""
_CELL_NAMES = frozenset(
    (b"c", b"f", b"v", b"row", b"is", b"sheetData", b"worksheet", b"macrosheet")
)
_ATTRIBUTE = re.compile(rb"""\s+([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"<]*)"|'([^'<]*)')""")
_ATTRIBUTES = re.compile(rb"""(?:\s+[A-Za-z_][\w:.-]*\s*=\s*(?:"[^"<]*"|'[^'<]*'))*\s*/?""")
_REFERENCE = re.compile(r"&(lt|gt|amp|quot|apos|#[0-9]+|#x[0-9A-Fa-f]+);")
_ENTITIES = {"lt": "<", "gt": ">", "amp": "&", "quot": '"', "apos": "'"}
_DECLARATION = re.compile(rb"""<\?xml[^>]*?encoding\s*=\s*["']([^"']*)["']""")


def _entity(match: re.Match[str]) -> str:
    name = match.group(1)
    if name[0] != "#":
        return _ENTITIES[name]
    return chr(int(name[2:], 16) if name[1] == "x" else int(name[1:]))


def _text(raw: bytes) -> str:
    """Character data as expat gives it: line ends made LF, references decoded."""
    text = raw.decode("utf-8")
    if "\r" in text:
        text = text.replace("\r\n", "\n").replace("\r", "\n")
    return _REFERENCE.sub(_entity, text) if "&" in text else text


def _attributes(raw: bytes) -> dict[str, str]:
    """A start tag's attributes as expat gives them (whitespace in a value made spaces)."""
    if _ATTRIBUTES.fullmatch(raw) is None:
        raise NotPlainError("attribute")
    out = {}
    for name, double, single in _ATTRIBUTE.findall(raw):
        value = (double or single).decode("utf-8")
        if "\t" in value or "\n" in value or "\r" in value:
            value = value.replace("\r\n", " ").translate(_SPACES)
        out[name.decode("utf-8")] = _REFERENCE.sub(_entity, value) if "&" in value else value
    return out


_SPACES = str.maketrans("\t\n\r", "   ")
_KNOWN: dict[bytes, dict[str, str]] = {}
"""A formula's start tag -> its attributes: a sheet repeats a few (``t="shared" si="0"``)."""
_KNOWN_KEPT = 4096  # the cache stops growing there, whatever a file holds


def _formula_attributes(raw: bytes) -> dict[str, str]:
    """A formula's attributes, a dict of its own (a caller may change it)."""
    known = _KNOWN.get(raw)
    if known is None:
        known = _attributes(raw)
        if len(_KNOWN) < _KNOWN_KEPT:
            _KNOWN[raw] = known
    return dict(known)


def _cell_start(data: bytes, before: int) -> int:
    """The start of the last ``<c`` element before ``before``, or -1."""
    while True:
        at = data.rfind(b"<c", 0, before)
        if at < 0 or data[at + 2 : at + 3] in _AFTER_NAME:
            return at
        before = at  # <col, <cfRule, ... or a name cut at the end of the data


_AFTER_NAME = (b" ", b">", b"/", b"\t", b"\n", b"\r")


def _plain_start(head: bytes) -> bool:
    """The part is UTF-8 and opens with a declaration (or none) that says so."""
    if head.startswith(b"\xef\xbb\xbf"):
        head = head[3:]
    if not head.startswith(b"<"):
        return False
    if head.startswith(b"<?xml"):
        found = _DECLARATION.match(head)
        return found is None or found.group(1).lower() in (b"utf-8", b"utf8")
    return True


def scan_formulas(stream: BinaryIO, cms: int | None = None) -> Iterator[FormulaCell]:
    """The formula cells of a worksheet part written plainly, found in its bytes.

    Args:
        stream: The part.
        cms: The number of cell-metadata entries: a formula cell whose ``cm`` is past it is
            refused, as :meth:`Sheet.cells` refuses it.

    Raises:
        NotPlainError: The part is not written plainly; the items given so far are the first items
            :func:`read_formulas` gives.
        PackageError: A formula cell's position, ``si`` or ``cm`` is refused.
        xml.parsers.expat.ExpatError: The part is not well-formed.
    """
    check = xml.parsers.expat.ParserCreate()
    tail = b""
    first = True
    data_done = False
    last_cell = -1

    def found(data: bytes, end: int) -> list[FormulaCell]:
        """The items of ``data[:end]``, all or none: a part found not plain gives none."""
        nonlocal data_done, last_cell
        closed = -1
        at = 0
        while not data_done:  # among the cells: anything that could hide one or name one otherwise
            special = _SPECIAL.search(data, at, end)
            if special is None:
                break
            at = special.end()
            if special.group(1) is not None:  # <! or <?: allowed only for the declaration
                if not (first and special.group(1) == b"?" and special.start() == 0):
                    raise NotPlainError("comment, CDATA section or processing instruction")
            elif special.group(2) is None:
                closed = special.start()
                break
            elif special.group(2) in _CELL_NAMES:
                raise NotPlainError("prefixed element")
        if data_done or closed >= 0:  # after the cells, a cell of any name is read by expat
            after = 0 if data_done else closed
            if _C_START.search(data, after, end):
                raise NotPlainError("cell after sheetData")  # a cell after sheetData
            data_done = True
        items = []
        for m in _F.finditer(data, 0, end):
            at = m.start()
            cell = _cell_start(data, at)
            if cell < 0 or cell == last_cell:
                raise NotPlainError("formula outside a cell, or a second in one")
            last_cell = cell
            excel = _EXCEL.match(data, cell, end)
            if excel is not None and excel.start(5) == at + 2:  # as Excel writes one: quickest
                letters, digits, t, cm_text, f_raw, raw, value = excel.groups()
                col = _LETTERS.get(letters, 0) or _letters(letters)
                row = int(digits)
                if row <= _MAX_ROW and col <= _MAX_COL:
                    cm = int(cm_text) if cm_text is not None else None
                    if cm is not None and cms is not None and cm > cms:
                        raise PackageError(
                            f"row {row}, column {col}: a format or metadata index past its table"
                        )
                    if f_raw:
                        formula = _formula(
                            _text(raw) if raw else None, _formula_attributes(f_raw), row, col
                        )
                    else:
                        text = unescape(_text(raw)) if raw else None
                        formula = Formula(text or None, "normal", None, None, {})
                    items.append(
                        FormulaCell(
                            row,
                            col,
                            formula,
                            cm,
                            value is not None and (len(value) > 7 or t == b"str"),  # <v></v>: 7
                        )
                    )
                    continue
            usual = _USUAL.match(data, cell, end)
            if usual is not None and usual.start(2) == at + 2:  # Excel's shape: the quick way
                attrs_c = _cell_attributes(usual.group(1))
                ref = attrs_c.get("r")
                if not ref:
                    raise NotPlainError("no r")
                row, col = _position(ref)
                if not 1 <= col <= _MAX_COL:
                    raise NotPlainError("column off the grid")
                raw = usual.group(3)
                value = usual.group(4)
                t = attrs_c.get("t", "n")
                items.append(
                    FormulaCell(
                        row,
                        col,
                        _formula(
                            _text(raw) if raw is not None else None,
                            _formula_attributes(usual.group(2)),
                            row,
                            col,
                        ),
                        _cm(attrs_c.get("cm"), row, col, cms),
                        value is not None and (value not in (b"<v/>", b"<v></v>") or t == "str"),
                    )
                )
                continue
            tag_end = data.find(b">", cell)
            tag = data[cell + 2 : tag_end]
            if tag.endswith(b"/") or data.find(b"</c>", tag_end, at) >= 0:
                raise NotPlainError("cell tag")
            attrs_c = _attributes(tag)
            ref = attrs_c.get("r")
            if not ref:
                raise NotPlainError("no r")
            row, col = _position(ref)
            if not 1 <= col <= _MAX_COL:
                raise NotPlainError("column off the grid")  # placed by position: expat counts
            f_attrs = _formula_attributes(m.group(1))
            text = _text(m.group(3)) if m.group(3) is not None else None
            cell_end = data.find(b"</c>", m.end(), end)
            if cell_end < 0 or _F_START.search(data, m.end(), cell_end):
                raise NotPlainError("second formula")  # a second formula in the cell
            t = attrs_c.get("t", "n")
            if t == "inlineStr":
                cached = _C_IS.search(data, tag_end, cell_end) is not None
            else:
                v_at = _V_START.search(data, tag_end, cell_end)
                if v_at is None:
                    cached = False
                else:
                    v = _V.match(data, v_at.start(), cell_end)
                    if v is None:
                        raise NotPlainError("value holding markup")  # a value holding markup
                    cached = bool(v.group(1)) or t == "str"
            items.append(
                FormulaCell(
                    row,
                    col,
                    _formula(text, f_attrs, row, col),
                    _cm(attrs_c.get("cm"), row, col, cms),
                    cached,
                )
            )
        if len(items) != len(_F_START.findall(data, 0, end)):
            raise NotPlainError("formula holding markup")  # a formula holding markup
        return items

    while chunk := stream.read(_CHUNK):
        check.Parse(chunk, False)
        data = tail + chunk
        if first and not _plain_start(data):
            raise NotPlainError("encoding")
        cut = _cell_start(data, len(data))
        if cut < 0:  # no cell: keep only what may begin one
            cut = data.rfind(b"<")
            if cut < 0:
                cut = len(data)
        yield from found(data, cut)
        first = False
        tail = data[cut:]
        if len(tail) > _TAIL:
            raise NotPlainError("long cell")
        last_cell = -1 if last_cell < cut else last_cell - cut
    check.Parse(b"", True)
    if first and tail == b"":
        return
    yield from found(tail, len(tail))


_ATTRS = rb"""((?:\s+[A-Za-z_][\w:.-]*\s*=\s*(?:"[^"<]*"|'[^'<]*'))*)\s*"""
_USUAL = re.compile(
    rb"<c" + _ATTRS + rb"><f" + _ATTRS + rb"(?:/>|>([^<]*)</f>)(<v>[^<]*</v>|<v/>)?</c>"
)
"""A formula cell as Excel writes it: its formula, then its value or none, nothing else."""


_EXCEL = re.compile(
    rb'<c r="([A-Z]{1,3})([1-9][0-9]{0,6})"(?: s="[0-9]{1,10}")?(?: t="([a-z]{1,9})")?'
    rb'(?: cm="([0-9]{1,10})")?(?: vm="[0-9]{1,10}")?><f' + _ATTRS + rb"(?:/>|>([^<]*)</f>)"
    rb"(<v>[^<]*</v>|<v/>)?</c>"
)
"""A formula cell exactly as Excel writes one: attributes in its order, a column in capitals."""
_LETTERS: dict[bytes, int] = {}


def _letters(letters: bytes) -> int:
    col = 0
    for o in letters:
        col = col * 26 + o - 64
    if len(_LETTERS) < _MAX_COL:
        _LETTERS[letters] = col
    return col


def _cell_attributes(raw: bytes) -> dict[str, str]:
    """A cell's attributes, its start tag already found well-formed by :data:`_USUAL`."""
    if b"&" in raw or b"\t" in raw or b"\n" in raw or b"\r" in raw:
        return _attributes(raw)
    return {
        name.decode("utf-8"): (double or single).decode("utf-8")
        for name, double, single in _ATTRIBUTE.findall(raw)
    }


_V = re.compile(rb"<v(?:\s[^>]*)?(?:/>|>([^<]*)</v>)")
_V_START = re.compile(rb"<v[\s/>]")
_C_IS = re.compile(rb"<is[\s/>]")


# The expat reader -----------------------------------------------------------------------------


def read_formulas(stream: BinaryIO, cms: int | None = None) -> Iterator[FormulaCell]:
    """The formula cells of a worksheet part, read with expat.

    It reads any part :func:`scan_formulas` does not take. Elements are matched by their local
    part, as :func:`xlio._cells.read_cells` matches them.

    Raises:
        PackageError: A formula cell's position, ``si`` or ``cm`` is refused.
        xml.parsers.expat.ExpatError: The part is not well-formed.
    """
    codes = _codes()
    out: list[FormulaCell] = []
    buffer: list[str] = []
    row = col = 0
    attrs_c: dict[str, str] = {}
    f_attrs: dict[str, str] | None = None
    f_text: str | None = None
    collect = 0  # 3 <v>, 4 <f>
    v_seen = v_text = inline = False

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal row, col, attrs_c, f_attrs, f_text, collect, v_seen, v_text, inline
        code = codes[name]
        if code == 2:
            ref = attrs.get("r")
            if ref:
                ref_row, ref_col = _position(ref)
                row = ref_row
                col = ref_col if 1 <= ref_col <= _MAX_COL else col + 1
            else:
                col += 1
            attrs_c = attrs
            f_attrs = f_text = None
            v_seen = v_text = inline = False
        elif code == 4:
            buffer.clear()
            collect = 4
            f_attrs = attrs
        elif code == 3:
            collect = 3
            v_seen = True
        elif code == 5:
            inline = True
        elif code == 1:
            r = attrs.get("r")
            row = _row(r) if r else row + 1
            col = 0

    def end(name: str) -> None:
        nonlocal collect, f_text
        code = codes[name]
        if code == 2 and f_attrs is not None:
            if col > _MAX_COL or not 1 <= row <= _MAX_ROW:
                raise PackageError(f"row {row}: more cells than columns")
            t = attrs_c.get("t", "n")
            cached = inline if t == "inlineStr" else v_seen and (v_text or t == "str")
            out.append(
                FormulaCell(
                    row,
                    col,
                    _formula(f_text, f_attrs, row, col),
                    _cm(attrs_c.get("cm"), row, col, cms),
                    cached,
                )
            )
        elif code == 4:
            f_text = "".join(buffer)
            collect = 0
        elif code == 3:
            collect = 0

    def text(data: str) -> None:
        nonlocal v_text
        if collect == 4:
            buffer.append(data)
        elif collect == 3 and data:
            v_text = True

    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    while chunk := stream.read(_CHUNK):
        p.Parse(chunk, False)
        yield from out
        out.clear()
    p.Parse(b"", True)
    yield from out
