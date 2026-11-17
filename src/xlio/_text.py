"""Text formats: a CSV file read as Excel for Mac opens it (File > Open, en_US).

Every rule here is one the captures ``text``, ``text_launch_*``, ``text_extra`` and
``text_limits`` show (``tests/captures/``), unless its comment says it is not captured. The file
becomes a one-sheet workbook built in memory and read by :func:`xlio.open`, so a text file and
an xlsx file give the same :class:`~xlio.Workbook`.
"""

from __future__ import annotations

import calendar
import dataclasses
import datetime
import itertools
import os
import re
import tempfile
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass
from decimal import ROUND_DOWN, Decimal
from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO, Literal
from xml.sax.saxutils import escape

from xlio._package import LimitError, ReadLimits

if TYPE_CHECKING:
    from xlio._workbook import Workbook

__all__ = ["CsvOptions", "TextOptions", "read_csv"]

Conversions = Literal["excel", "keep", "none"]
Formulas = Literal["formula", "text"]


@dataclass(frozen=True, kw_only=True)
class TextOptions:
    """How a text file is read; every default is what Excel for Mac does with File > Open.

    Attributes:
        encoding: The file's encoding. ``None`` is Excel for Mac's detection: a UTF-8 or a
            UTF-16 little-endian byte order mark, else ``fallback_encoding`` (so a UTF-8 file
            without a mark reads as Mac Roman, as Excel shows it).
        fallback_encoding: The encoding of a file without a byte order mark, when ``encoding``
            is ``None``: ``"mac_roman"``, as Excel for Mac reads it.
        conversions: ``"excel"``, Excel's conversions (numbers, dates, times, percentages,
            currency, booleans, errors); ``"keep"``, the conversions Excel makes when its notice
            of default conversions is answered "Don't Convert" (exponent forms, leading zeros,
            more than 15 digits and ``0 1/2`` stay text); ``"none"``, every field text.
        formulas: ``"formula"``: a field Excel takes as a formula (``=1+1``, ``+1+1``,
            ``@SUM(1)``) is a formula cell without a cached value, as Excel holds it before
            calculating; ``"text"``: it stays the text written.
        today: The date a date without a year takes its year from (``1/2``); ``None`` is the
            clock's.
        sheet_name: The sheet's name; ``None`` takes the file's name as Excel does (its stem,
            brackets as parentheses, cut to 31 characters), or ``Sheet1`` for a stream.
    """

    encoding: str | None = None
    fallback_encoding: str = "mac_roman"
    conversions: Conversions = "excel"
    formulas: Formulas = "formula"
    today: datetime.date | None = None
    sheet_name: str | None = None


@dataclass(frozen=True, kw_only=True)
class CsvOptions(TextOptions):
    """How a delimited file is read.

    Attributes:
        delimiter: The field delimiter: a comma, as Excel reads a ``.csv`` file.
        quote: The quote character.
    """

    delimiter: str = ","
    quote: str = '"'

    @classmethod
    def tsv(cls, **changes: object) -> CsvOptions:
        """Options for tab-delimited text, otherwise a CSV file's."""
        return dataclasses.replace(cls(delimiter="\t"), **changes)  # type: ignore[arg-type]


_MAX_ROW, _MAX_COL = 1_048_576, 16_384
_SPOOL = 8 << 20  # a built package larger than this goes to a temporary file
_LINE = 32_759  # a longer line continues on the next row (capture text_limits)


def _decode(data: bytes, options: TextOptions) -> str:
    """The file's text, as Excel for Mac decodes it (capture ``text``)."""
    if options.encoding is not None:
        text = data.decode(options.encoding, errors="replace")
        return text[1:] if text[:1] == "﻿" else text
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace")
    if data.startswith(b"\xff\xfe"):
        return data[2:].decode("utf-16-le", errors="replace")
    # A UTF-16 big-endian mark is not taken: its bytes read as the fallback encoding, and the
    # NUL characters dropped (capture text, case "encoding utf-16-be with BOM").
    return data.decode(options.fallback_encoding, errors="replace").replace("\x00", "")


def _records(text: str, delimiter: str, quote: str) -> Iterator[list[str]]:
    """The rows and fields, as Excel splits them (capture ``text``, the ``structure`` cases).

    A quote opens a quoted run only as a field's first character; inside it ``""`` is a quote,
    and the closing quote ends the run while the field goes on (``"a"b`` is ``ab``). An
    unterminated run takes the rest of the file. CR, LF and CRLF end a row; an empty line is an
    empty row.
    """
    fields: list[str] = []
    field: list[str] = []
    i, n = 0, len(text)
    quoted = start = False
    start = True
    while i < n:
        ch = text[i]
        if quoted:
            if ch == quote:
                if i + 1 < n and text[i + 1] == quote:
                    field.append(quote)
                    i += 2
                    continue
                quoted = False
            else:
                field.append(ch)
            i += 1
            continue
        if start and ch == quote:
            quoted, start = True, False
            i += 1
            continue
        start = False
        if ch == delimiter:
            fields.append("".join(field))
            field = []
            start = True
        elif ch in "\r\n":
            fields.append("".join(field))
            yield fields
            fields, field, start = [], [], True
            if ch == "\r" and i + 1 < n and text[i + 1] == "\n":
                i += 1
        else:
            field.append(ch)
        i += 1
    if field or fields or quoted:
        fields.append("".join(field))
        yield fields


# Conversions --------------------------------------------------------------------------------

_CURRENCY = '"$"#,##0_);[Red]\\("$"#,##0\\)'
_CURRENCY_CENTS = '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)'
_PLAIN = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)")
_EXPONENT = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+")
_THOUSANDS = re.compile(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d*)?")
_PERCENT = re.compile(r"([+-]?(?:\d+\.?\d*|\.\d+))%|%((?:\d+\.?\d*|\.\d+))")
_DOLLARS = re.compile(r"(\()?\$(\d{1,3}(?:,\d{3})+(?:\.\d*)?|\d+\.?\d*|\.\d+)(\))?")
_PARENTHESES = re.compile(r"\((\d+\.?\d*|\.\d+|\d{1,3}(?:,\d{3})+(?:\.\d*)?)\)")
_FRACTION = re.compile(r"(-?)(\d+) (\d+)/(\d+)")
_MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
_MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_name) if m})
_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
# Newer errors as Excel reads them in a CSV file (capture text_forms); #EXTERNAL! stays text.
_NEWER_ERRORS = {"#SPILL!": "#VALUE!", "#CALC!": "#VALUE!", "#FIELD!": "#VALUE!",
                 "#BLOCKED!": "#VALUE!", "#CONNECT!": "#VALUE!", "#BUSY!": "#VALUE!",
                 "#UNKNOWN!": "#VALUE!", "#PYTHON!": "#VALUE!", "#TIMEOUT!": "#VALUE!",
                 "#GETTING_DATA": "#N/A"}  # fmt: skip

Value = tuple[str, object, "int | str"]  # (kind, value, number format: a built-in id or a code)


def _fifteen(text: str) -> float:
    """A decimal as Excel keeps it: cut (not rounded) to 15 significant digits."""
    d = Decimal(text)
    if d == 0:
        return 0.0
    exponent = d.adjusted() - 14
    return float(d.quantize(Decimal(1).scaleb(exponent), rounding=ROUND_DOWN))


def _too_many_digits(text: str) -> bool:
    digits = re.sub(r"\D", "", text.split("e")[0].split("E")[0]).lstrip("0")
    return len(digits.rstrip("0") if "." in text else digits) > 15


def _number(s: str, keep: bool) -> Value | None:
    """A number Excel converts, with the format it assigns; ``None`` when it is not one."""
    if _PLAIN.fullmatch(s):
        if keep and (re.match(r"[+-]?0\d", s) or _too_many_digits(s)):
            return None
        return "n", _fifteen(s) + 0.0, 0
    if _EXPONENT.fullmatch(s):
        value = float(s)
        if value in (float("inf"), float("-inf")) or (
            value == 0 and _fifteen(s.lower().split("e")[0]) != 0
        ):
            return "str", s, 11  # out of a double's range: text, with the format (capture text)
        if value != 0 and abs(value) < 2.2250738585072014e-308:
            return "str", s, 11
        if keep:  # "Don't Convert": an exponent form in range stays text (capture text_launch_dont)
            return None
        return "n", _fifteen(f"{Decimal(s):f}") + 0.0, 11
    if _THOUSANDS.fullmatch(s):
        return "n", _fifteen(s.replace(",", "")), 4 if "." in s else 3
    m = _PERCENT.fullmatch(s)
    if m:
        number = m[1] or m[2]
        return "n", _fifteen(number) / 100, 10 if "." in number else 9
    m = _DOLLARS.fullmatch(s)
    if m and bool(m[1]) == bool(m[3]):
        value = _fifteen(m[2].replace(",", ""))
        return "n", -value if m[1] else value, _CURRENCY_CENTS if "." in m[2] else _CURRENCY
    m = _PARENTHESES.fullmatch(s)
    if m:
        return "n", -_fifteen(m[1].replace(",", "")), 3 if "," in m[1] else 0
    m = _FRACTION.fullmatch(s)
    if m and int(m[4]) != 0:
        if keep and m[2] == "0":
            return None
        value = int(m[2]) + int(m[3]) / int(m[4])
        # "# ?/?" for a one-digit denominator, "# ??/??" for more (capture text_forms)
        return "n", -value if m[1] else value, 12 if int(m[4]) < 10 else 13
    return None


def _serial(day: datetime.date) -> int | None:
    """Excel's serial number of a date (1900 system), or ``None`` outside 1900-9999."""
    if not 1900 <= day.year <= 9999:
        return None
    serial = (day - datetime.date(1899, 12, 30)).days
    return serial if day >= datetime.date(1900, 3, 1) else serial - 1


def _year(text: str) -> int | None:
    if len(text) == 4:
        return int(text)
    if len(text) <= 2:  # 00-49 are 2000-2049, 50-99 1950-1999 (capture text_extra)
        y = int(text)
        return 2000 + y if y < 50 else 1900 + y
    return None


def _day(year: int | None, month: int, day: int) -> int | None:
    if year is None or not 1 <= month <= 12:
        return None
    try:
        return _serial(datetime.date(year, month, day))
    except ValueError:
        return None


_N = r"(\d{1,2})"
_DATES: list[tuple[re.Pattern[str], str, int]] = [
    (re.compile(rf"{_N}[/-]{_N}[/-](\d{{1,4}})"), "mdy", 14),
    (re.compile(r"(\d{4})[/-]" + _N + r"[/-]" + _N), "ymd", 14),
    (re.compile(rf"{_N}[/-]{_N}"), "md", 16),
    (re.compile(rf"{_N}[- ]([A-Za-z]+)[- ](\d{{2,4}})"), "dMy", 15),
    (re.compile(rf"([A-Za-z]+) {_N}, (\d{{4}})"), "Mdy", 15),
    (re.compile(rf"{_N}-([A-Za-z]+)"), "dM", 16),
    (re.compile(rf"([A-Za-z]+)-{_N}"), "Md", 16),
    (re.compile(r"([A-Za-z]+) (\d{4})"), "My", 17),
]


def _date(s: str, today: datetime.date) -> tuple[int, int] | None:
    """A date Excel reads in en_US (month first), as its serial and format."""
    for pattern, order, fmt in _DATES:
        m = pattern.fullmatch(s)
        if m is None:
            continue
        g = m.groups()
        if order == "mdy":
            serial = _day(_year(g[2]), int(g[0]), int(g[1]))
        elif order == "ymd":
            serial = _day(int(g[0]), int(g[1]), int(g[2]))
        elif order == "md":
            serial = _day(today.year, int(g[0]), int(g[1]))
        elif order == "dMy":
            serial = _day(_year(g[2]), _MONTHS.get(g[1].lower(), 0), int(g[0]))
        elif order == "Mdy":
            serial = _day(int(g[2]), _MONTHS.get(g[0].lower(), 0), int(g[1]))
        elif order == "dM":
            serial = _day(today.year, _MONTHS.get(g[1].lower(), 0), int(g[0]))
        elif order == "Md":
            serial = _day(today.year, _MONTHS.get(g[0].lower(), 0), int(g[1]))
        else:  # "My": the first of the month
            serial = _day(int(g[1]), _MONTHS.get(g[0].lower(), 0), 1)
        if serial is not None:
            return serial, fmt
        return None
    return None


_TIME = re.compile(r"(\d{1,4}):(\d{2})(?::(\d{2})(\.\d+)?)?(?: ?([AaPp][Mm]))?")


def _time(s: str) -> tuple[float, int] | None:
    """A time of day as a fraction of a day, and its format."""
    m = _TIME.fullmatch(s)
    if m is None:
        return None
    hours, minutes = int(m[1]), int(m[2])
    seconds = int(m[3]) if m[3] else 0
    fraction = float(m[4]) if m[4] else 0.0
    if minutes > 59 or seconds > 59:
        return None
    if m[5]:
        if not 1 <= hours <= 12:
            return None
        hours = hours % 12 + (12 if m[5].lower() == "pm" else 0)
        fmt = 19 if m[3] else 18
    elif m[4]:
        fmt = 47
    elif hours > 23:
        fmt = 46
    else:
        fmt = 21 if m[3] else 20
    return (hours * 3600 + minutes * 60 + seconds + fraction) / 86400, fmt


def _formula(field: str) -> str | None:
    """The formula a field holds, as Excel takes it, or ``None``.

    ``=``, ``+``, ``-`` and ``@`` start one (``+1+1`` is ``1+1``, ``@SUM(1)`` is ``SUM(1)``);
    ``=`` alone, ``==1``, text ending in an operator (``-5-``, ``=1+``, ``=A1:``) and unbalanced
    parentheses (``=SUM(``, ``=1)``) stay text, while ``=foo bar`` is a formula. Excel parses
    the formula; xlio does not, so this test of its shape follows the cases captured (captures
    ``text`` and ``text_forms``) and may differ from Excel's parser elsewhere.
    """
    if not field or field[0] not in "=+-@":
        return None
    body = field[1:] if field[0] in "=+@" else field
    if not body or body[0] == "=" or body[-1] in "+-*/^&=<>,(:":
        return None
    if body.count("(") != body.count(")"):
        return None
    return body


def convert(field: str, options: TextOptions, today: datetime.date) -> Value | None:
    """A field as Excel holds it: (kind, value, number format), or ``None`` for no cell.

    Kinds: ``n`` (a float), ``str``, ``b`` (a bool), ``e`` (an error's text) and ``f`` (a
    formula's text, without the ``=``).
    """
    if field == "":
        return None
    if options.conversions == "none":
        return "str", field, 0
    s = field.strip(" ")  # spaces around a number are dropped; around text, kept
    upper = s.upper()
    if upper in ("TRUE", "FALSE") and s == field:
        return "b", upper == "TRUE", 0
    if upper in _ERRORS and s == field:
        return "e", upper, 0
    if upper in _NEWER_ERRORS and s == field:
        return "e", _NEWER_ERRORS[upper], 0
    keep = options.conversions == "keep"
    number = _number(s, keep) if s else None
    if number is not None:
        return number
    date = _date(s, today)
    if date is not None:
        return "n", float(date[0]), date[1]
    time = _time(s)
    if time is not None:
        return "n", time[0], time[1]
    if " " in s:
        day_part, _, time_part = s.rpartition(" ")
        if time_part.upper() in ("AM", "PM"):
            day_part, _, clock = day_part.rpartition(" ")
            time_part = f"{clock} {time_part}"
        date, time = _date(day_part, today), _time(time_part)
        if date is not None and time is not None:
            return "n", date[0] + time[0], 22
    if options.formulas == "formula":
        formula = _formula(field)
        if formula is not None:
            return "f", formula, 0
    return "str", field, 0


# The workbook --------------------------------------------------------------------------------

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_TYPES = "application/vnd.openxmlformats-officedocument.spreadsheetml."


def _letters(col: int) -> str:
    out = ""
    while col:
        col, r = divmod(col - 1, 26)
        out = chr(65 + r) + out
    return out


def _text_xml(text: str) -> str:
    """Text as element content: XML's escapes, and OOXML's for characters XML cannot hold."""
    out = escape(text.replace("_x", "_x005F_x"))
    return re.sub(
        r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff￾￿]",
        lambda m: f"_x{ord(m[0]):04X}_",
        out,
    )


def _package(
    rows: Iterator[list[str]], name: str, options: TextOptions, today: datetime.date
) -> BinaryIO:
    """The workbook package, its sheet written row by row as the rows come.

    The package goes to a temporary file past a few megabytes (deleted when it is closed), so
    a large file's package is not held in memory beside its text and shared strings.
    """
    # Each text is written to the shared strings as it comes, without looking for an equal
    # one: a table of the texts met would cost far more than the file (a value reads the same).
    strings = tempfile.SpooledTemporaryFile(max_size=_SPOOL)  # noqa: SIM115
    count = 0
    formats: dict[int | str, int] = {0: 0}
    out = tempfile.SpooledTemporaryFile(max_size=_SPOOL)  # noqa: SIM115 - the workbook keeps it
    z = zipfile.ZipFile(out, "w", zipfile.ZIP_STORED)
    sheet = z.open("xl/worksheets/sheet1.xml", "w", force_zip64=True)
    sheet.write(f'<worksheet xmlns="{_MAIN}"><sheetData>'.encode())
    for r, fields in enumerate(itertools.islice(rows, _MAX_ROW), 1):
        cells = []
        for c, field in enumerate(fields[:_MAX_COL], 1):
            value = convert(field, options, today)
            if value is None:
                continue
            kind, data, fmt = value
            style = formats.setdefault(fmt, len(formats))
            ref = f'r="{_letters(c)}{r}"' + (f' s="{style}"' if style else "")
            if kind == "n":
                cells.append(f"<c {ref}><v>{float(data)!r}</v></c>")  # type: ignore[arg-type]
            elif kind == "b":
                cells.append(f'<c {ref} t="b"><v>{int(bool(data))}</v></c>')
            elif kind == "e":
                cells.append(f'<c {ref} t="e"><v>{escape(str(data))}</v></c>')
            elif kind == "f":
                cells.append(f"<c {ref}><f>{_text_xml(str(data))}</f></c>")
            else:
                index, count = count, count + 1
                strings.write(
                    f'<si><t xml:space="preserve">{_text_xml(str(data))}</t></si>'.encode()
                )
                cells.append(f'<c {ref} t="s"><v>{index}</v></c>')
        if cells:
            sheet.write((f'<row r="{r}">' + "".join(cells) + "</row>").encode("utf-8"))
    sheet.write(b"</sheetData></worksheet>")
    sheet.close()
    custom = {code: 164 + i for i, code in enumerate(f for f in formats if isinstance(f, str))}
    xfs = "".join(
        f'<xf numFmtId="{custom[f] if isinstance(f, str) else f}" fontId="0" fillId="0"'
        ' borderId="0"' + (' applyNumberFormat="1"/>' if f else "/>")
        for f in formats
    )
    num_fmts = "".join(
        f'<numFmt numFmtId="{i}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>'
        for code, i in custom.items()
    )
    styles = (
        f'<styleSheet xmlns="{_MAIN}">'
        + (f'<numFmts count="{len(custom)}">{num_fmts}</numFmts>' if custom else "")
        + '<fonts count="1"><font><sz val="12"/><name val="Aptos Narrow"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(formats)}">{xfs}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
            'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{_TYPES}sheet.main+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{_TYPES}worksheet+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{_TYPES}styles+xml"/>'
            f'<Override PartName="/xl/sharedStrings.xml" ContentType="{_TYPES}sharedStrings+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{_PKG}"><Relationship Id="rId1" '
            f'Type="{_DOC}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        ),
        "xl/workbook.xml": (
            f'<workbook xmlns="{_MAIN}" xmlns:r="{_DOC}"><sheets>'
            f'<sheet name="{escape(name, {chr(34): "&quot;"})}" sheetId="1" r:id="rId1"/>'
            "</sheets></workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{_PKG}">'
            f'<Relationship Id="rId1" Type="{_DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_DOC}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId3" Type="{_DOC}/sharedStrings" Target="sharedStrings.xml"/>'
            "</Relationships>"
        ),
        "xl/styles.xml": styles,
    }
    with z:
        for part, body in parts.items():
            z.writestr(part, body.encode("utf-8"))
        with z.open("xl/sharedStrings.xml", "w", force_zip64=True) as sst:
            sst.write(f'<sst xmlns="{_MAIN}">'.encode())
            strings.seek(0)
            while chunk := strings.read(1 << 20):
                sst.write(chunk)
            sst.write(b"</sst>")
        strings.close()
    out.seek(0)
    return out  # type: ignore[return-value]


def _lines(text: str, options: CsvOptions) -> Iterator[list[str]]:
    """The rows, a line past Excel's length continuing below (capture text_limits)."""
    for row in _records(text, options.delimiter, options.quote):
        if len(row) != 1 or len(row[0]) <= _LINE:
            yield row
        else:
            yield from ([row[0][i : i + _LINE]] for i in range(0, len(row[0]), _LINE))


def sheet_name_of(path: str | os.PathLike[str]) -> str:
    """The sheet's name Excel gives a file: its stem, brackets as parentheses, 31 characters."""
    stem = Path(path).stem
    return stem.replace("[", "(").replace("]", ")")[:31]


def read_csv(
    source: str | os.PathLike[str] | BinaryIO,
    options: CsvOptions | None = None,
    *,
    limits: ReadLimits | None = None,
) -> Workbook:
    """Read a delimited text file as Excel for Mac opens it: a one-sheet :class:`Workbook`.

    Args:
        source: A path, or a binary file open for reading.
        options: How to read it (:class:`CsvOptions`); Excel's defaults when ``None``.
        limits: What reading may cost: the file's size is bounded by ``max_text_bytes``, and the
            workbook built from it by the other limits.

    Returns:
        The workbook, one sheet built in memory: its parts (``Workbook.parts``) are the ones
        xlio synthesizes for it, not a file's, and :func:`xlio.patch` cannot write it.

    Raises:
        LimitError: The file is larger than ``limits.max_text_bytes``.
    """
    from xlio._workbook import open as open_workbook

    options = options or CsvOptions()
    limits = limits or ReadLimits()
    if isinstance(source, (str, os.PathLike)):
        name = options.sheet_name or sheet_name_of(source)
        with Path(source).open("rb") as f:
            data = f.read(limits.max_text_bytes + 1)
    else:
        name = options.sheet_name or "Sheet1"
        data = source.read(limits.max_text_bytes + 1)
    if len(data) > limits.max_text_bytes:
        raise LimitError(f"the file is larger than the limit of {limits.max_text_bytes} bytes")
    text = _decode(data, options)
    del data
    today = options.today or datetime.date.today()
    package = _package(_lines(text, options), name, options, today)
    del text
    return open_workbook(package, limits=limits)
