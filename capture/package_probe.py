"""What Excel does with a package xlio refuses or doubts: which refusals cost anything.

    uv run python -m capture.package_probe build   # write the workbooks only
    uv run python -m capture.package_probe run     # build, open in Excel, write TSV

Each case is one small workbook holding 111 in A1 of its only sheet, with one thing changed in
its worksheet part: a minimal ``<!DOCTYPE>``, with no entity and no external reference; another
compression method; the part stored twice under one name (the second holding 222); sizes in
the central directory that are not the part's; or a number no column, row or font can have.
``package_open.applescript`` opens each one in its own run, under a time limit, and reads A1, the
sizes of columns and rows and the font size back. A control opens first and last: an outcome counts
only between two that open. The output is ``tests/captures/package.tsv``. New cases go before the
last control.
"""

from __future__ import annotations

import datetime as dt
import io
import struct
import subprocess
import sys
import warnings
import zipfile
import zlib
from collections.abc import Callable
from pathlib import Path

import openpyxl

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "package.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "xlio"
    / "package"
)
SHEET = "xl/worksheets/sheet1.xml"
TIME_LIMIT = "60"


def base() -> bytes:
    """A workbook holding 111 in A1, as openpyxl writes it."""
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = 111
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _rewrite(data: bytes, write_sheet: Callable[[zipfile.ZipFile, bytes], None]) -> bytes:
    """A copy of the package with its worksheet part written by ``write_sheet``."""
    out = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as source,
        zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target,
    ):
        for info in source.infolist():
            if info.filename == SHEET:
                write_sheet(target, source.read(SHEET))
            else:
                target.writestr(info.filename, source.read(info.filename))
    return out.getvalue()


def _directory_sizes(data: bytes, *, size: int, crc: int | None = None) -> bytes:
    """Rewrite the worksheet part's uncompressed size, and its CRC, in the central directory."""
    patched = bytearray(data)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        directory = archive.start_dir
    position = directory
    while patched[position : position + 4] == b"PK\x01\x02":
        name_length, extra_length, comment_length = struct.unpack_from(
            "<HHH", patched, position + 28
        )
        name = bytes(patched[position + 46 : position + 46 + name_length])
        if name == SHEET.encode():
            struct.pack_into("<I", patched, position + 24, size)
            if crc is not None:
                struct.pack_into("<I", patched, position + 16, crc)
            return bytes(patched)
        position += 46 + name_length + extra_length + comment_length
    raise KeyError(SHEET)


def control(data: bytes) -> bytes:
    return data


def doctype(data: bytes) -> bytes:
    def write(target: zipfile.ZipFile, xml: bytes) -> None:
        at = xml.index(b"<worksheet")
        target.writestr(SHEET, xml[:at] + b"<!DOCTYPE worksheet>" + xml[at:])

    return _rewrite(data, write)


def _method(method: int) -> Callable[[bytes], bytes]:
    def case(data: bytes) -> bytes:
        return _rewrite(data, lambda target, xml: target.writestr(SHEET, xml, compress_type=method))

    return case


def duplicate(data: bytes) -> bytes:
    def write(target: zipfile.ZipFile, xml: bytes) -> None:
        target.writestr(SHEET, xml)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # zipfile warns of the duplicate name
            target.writestr(SHEET, xml.replace(b"<v>111</v>", b"<v>222</v>"))

    return _rewrite(data, write)


def false_smaller(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        size = archive.getinfo(SHEET).file_size
    return _directory_sizes(data, size=size - 100)


def false_larger(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        size = archive.getinfo(SHEET).file_size
    return _directory_sizes(data, size=size + 1000)


def false_prefix(data: bytes) -> bytes:
    """The directory describes the whole part as far as its XML; 100 KB of spaces follow it."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        xml = archive.read(SHEET)
    padded = _rewrite(data, lambda target, _: target.writestr(SHEET, xml + b" " * 100_000))
    return _directory_sizes(padded, size=len(xml), crc=zlib.crc32(xml))


def _edit(part: str, old: str, new: str) -> Callable[[bytes], bytes]:
    """A case changing one piece of one part's XML."""

    def case(data: bytes) -> bytes:
        out = io.BytesIO()
        with (
            zipfile.ZipFile(io.BytesIO(data)) as source,
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target,
        ):
            for info in source.infolist():
                body = source.read(info.filename)
                if info.filename == part:
                    text = body.decode()
                    assert old in text, old
                    body = text.replace(old, new, 1).encode()
                target.writestr(info.filename, body)
        return out.getvalue()

    return case


def _width(value: str) -> Callable[[bytes], bytes]:
    col = f'<cols><col min="1" max="1" width="{value}" customWidth="1"/></cols><sheetData>'
    return _edit(SHEET, "<sheetData>", col)


def _height(value: str) -> Callable[[bytes], bytes]:
    return _edit(SHEET, '<row r="1"', f'<row r="1" ht="{value}" customHeight="1"')


def _sheet_format(attribute: str, value: str) -> Callable[[bytes], bytes]:
    return _edit(SHEET, "<sheetFormatPr ", f'<sheetFormatPr {attribute}="{value}" ')


def _font_size(value: str) -> Callable[[bytes], bytes]:
    return _edit("xl/styles.xml", '<sz val="11"', f'<sz val="{value}"')


CASES: tuple[tuple[str, Callable[[bytes], bytes]], ...] = (
    ("control", control),
    ("doctype", doctype),
    ("bzip2", _method(zipfile.ZIP_BZIP2)),
    ("lzma", _method(zipfile.ZIP_LZMA)),
    ("duplicate_name", duplicate),
    ("size_smaller", false_smaller),
    ("size_larger", false_larger),
    ("size_prefix", false_prefix),
    ("width_nan", _width("nan")),
    ("width_NaN", _width("NaN")),
    ("width_inf", _width("inf")),
    ("width_INF", _width("INF")),
    ("width_negative", _width("-1")),
    ("width_1e308", _width("1e308")),
    ("height_nan", _height("nan")),
    ("height_NaN", _height("NaN")),
    ("height_inf", _height("inf")),
    ("height_INF", _height("INF")),
    ("height_negative", _height("-1")),
    ("height_1e308", _height("1e308")),
    ("font_size_1e308", _font_size("1e308")),
    ("font_size_nan", _font_size("nan")),
    ("font_size_negative", _font_size("-1")),
    ("default_width_nan", _sheet_format("defaultColWidth", "nan")),
    ("default_width_negative", _sheet_format("defaultColWidth", "-1")),
    ("default_height_nan", _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="nan"')),
    ("default_height_negative", _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="-1"')),
    ("width_256", _width("256")),
    ("width_1e4", _width("1e4")),
    ("width_10922", _width("10922")),
    ("width_1e6", _width("1e6")),
    ("width_1e9", _width("1e9")),
    ("width_99999999999", _width("99999999999")),
    ("width_1e30", _width("1e30")),
    ("width_minus_0.5", _width("-0.5")),
    ("width_minus_2", _width("-2")),
    ("width_minus_100", _width("-100")),
    ("width_minus_inf", _width("-inf")),
    ("height_410", _height("410")),
    ("height_1e4", _height("1e4")),
    ("height_1e6", _height("1e6")),
    ("height_1e9", _height("1e9")),
    ("height_1e30", _height("1e30")),
    ("height_minus_0.5", _height("-0.5")),
    ("height_minus_100", _height("-100")),
    ("height_minus_inf", _height("-inf")),
    ("font_size_inf", _font_size("inf")),
    ("font_size_0", _font_size("0")),
    ("font_size_0.5", _font_size("0.5")),
    ("font_size_410", _font_size("410")),
    ("font_size_1e6", _font_size("1e6")),
    ("font_size_1e30", _font_size("1e30")),
    ("font_size_minus_0.5", _font_size("-0.5")),
    ("default_width_inf", _sheet_format("defaultColWidth", "inf")),
    ("default_width_300", _sheet_format("defaultColWidth", "300")),
    ("default_width_1e30", _sheet_format("defaultColWidth", "1e30")),
    ("default_width_1e308", _sheet_format("defaultColWidth", "1e308")),
    ("default_width_99999999999", _sheet_format("defaultColWidth", "99999999999")),
    ("base_width_300", _edit(SHEET, 'baseColWidth="8"', 'baseColWidth="300"')),
    ("base_width_99999999999", _edit(SHEET, 'baseColWidth="8"', 'baseColWidth="99999999999"')),
    ("base_width_minus_1", _edit(SHEET, 'baseColWidth="8"', 'baseColWidth="-1"')),
    ("default_height_inf", _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="inf"')),
    ("default_height_1e308", _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="1e308"')),
    (
        "custom_default_height_nan",
        _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="nan" customHeight="1"'),
    ),
    (
        "custom_default_height_negative",
        _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="-1" customHeight="1"'),
    ),
    (
        "custom_default_height_1e308",
        _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="1e308" customHeight="1"'),
    ),
    (
        "custom_default_height_20",
        _edit(SHEET, 'defaultRowHeight="15"', 'defaultRowHeight="20" customHeight="1"'),
    ),
    ("control_again", control),
)


def build() -> list[Path]:
    """Write one workbook per case into Excel's container."""
    CONTAINER_DIR.mkdir(parents=True, exist_ok=True)
    data = base()
    paths = []
    for name, case in CASES:
        path = CONTAINER_DIR / f"package_{name}.xlsx"
        path.write_bytes(case(data))
        paths.append(path)
    return paths


def _version() -> str:
    result = subprocess.run(
        [
            "gtimeout",
            "-k",
            "5",
            "20",
            "osascript",
            "-e",
            'tell application "Microsoft Excel" to get version',
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.stdout.strip() or "unknown"


def _open(path: Path) -> str:
    """One case in its own ``osascript``; alerts are restored even when the limit stops it."""
    script = Path(__file__).with_name("package_open.applescript")
    # -k: osascript waiting for an answer from Excel does not stop for SIGTERM.
    result = subprocess.run(
        ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(script), str(path)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode in (124, 137):
        restore = Path(__file__).with_name("restore_alerts.applescript")
        subprocess.run(["gtimeout", "-k", "5", "20", "osascript", str(restore)], check=False)
        return f"timed out after {TIME_LIMIT} s"
    if result.returncode != 0:
        raise SystemExit(f"{path.name}: {result.stderr.strip()}")
    return result.stdout.strip()


def run() -> None:
    """Open every case, print the outcome and write the TSV. A case that hangs ends the run."""
    refuse_open_workbooks()
    edition = edition_line()
    paths = build()
    version = _version()
    rows = []
    finished = True
    for (name, _), path in zip(CASES, paths, strict=True):
        outcome = _open(path)
        print(f"{name}\t{outcome}", flush=True)
        rows.append(f"{name}\t{outcome}")
        if outcome.startswith("timed out"):
            finished = False
            break
    if not finished:
        sys.exit("a case hung: stopped, and nothing written. Look into it before running again.")
    header = [
        f"# suite: package (capture.package_probe); Excel {version} for Mac, en_US",
        edition,
        f"# captured: {dt.date.today().isoformat()}; method: open with alerts off, read A1 (111 "
        "unless the case changes it), the number of sheets, the widths of A and B, the heights "
        "of rows 1 and 2, and A1's font size",
        "case\toutcome\tA1\tsheets\twidth A\theight 1\twidth B\theight 2\tfont size",
    ]
    write_capture(OUTPUT, header + rows)
    print(f"{len(rows)} rows written to {OUTPUT}")
    for path in paths:  # kept only when a case hung, to look into
        path.unlink(missing_ok=True)


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
