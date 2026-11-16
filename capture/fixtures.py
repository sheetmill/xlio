"""Workbooks saved by Excel for the tests: tables and filters, conditional formats, drawings, notes.

    uv run python -m capture.fixtures [NAME ...]
    uv run python -m capture.fixtures --adopt FILE NAME [--replace-media PART ...]
    uv run python -m capture.fixtures --rework NAME [--replace-media PART ...]

XlsxWriter builds each workbook in Excel's container; Excel opens it with alerts off and saves a
copy (``save_copy.applescript``). The copy is stripped of what would name the machine or the
person (``x15ac:absPath``, ``cp:lastModifiedBy``), checked for home-directory paths, and written to
``tests/files/excel/NAME.xlsx``. The tests read what Excel wrote; nothing here is a reading rule.

``--adopt`` takes a workbook a person made in Excel by hand (what no script reaches: threaded
comments, icons, pictures in cells) from Excel's container folder. Beside the stripping above,
it makes the threaded-comment people anonymous (``Writer``, provider ``None``, as Excel writes a
person who is not signed in), drops ``dc:creator``, and replaces each ``--replace-media`` PNG
with a solid-color one of the same size, for an image that could identify someone.

Run it only while Excel is running and answers AppleScript, no other capture is running, and no
test or run is feeding crafted input to the reader (CAPTURING.md).
"""

from __future__ import annotations

import io
import re
import struct
import sys
import zipfile
import zlib
from collections.abc import Callable
from pathlib import Path

import xlsxwriter

from capture.excel import refuse_open_workbooks
from capture.paths import HOME_PATH
from capture.probe import CONTAINER, ROOT, save_copy

OUTPUT = ROOT / "tests" / "files" / "excel"


def png(width: int, height: int, rgb: tuple[int, int, int]) -> bytes:
    """A solid-color PNG, built here so no image file is needed."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    row = b"\x00" + bytes(rgb) * width
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(row * height))
        + chunk(b"IEND", b"")
    )


def tables(wb: xlsxwriter.Workbook) -> None:
    """A sheet autofilter with a criterion; tables with and without style, totals, buttons."""
    ws = wb.add_worksheet("Filters")
    rows = [["Region", "Units", "Price"], ["East", 10, 2.5], ["West", 20, 3.0], ["East", 5, 4.0]]
    for r, row in enumerate(rows):
        ws.write_row(r, 0, row)
    ws.autofilter("A1:C4")
    ws.filter_column("A", "x == East")
    ws.set_row(2, None, None, {"hidden": True})  # West, hidden by the filter
    data = [[1, 10], [2, 20], [3, 30]]
    ws.add_table(
        "E1:F5",
        {
            "name": "Sales",
            "data": data,
            "style": "Table Style Medium 2",
            "total_row": True,
            "columns": [{"header": "Id"}, {"header": "Amount", "total_function": "sum"}],
        },
    )
    ws.add_table(
        "H1:I4",
        {
            "name": "Plain",
            "data": data,
            "style": None,
            "autofilter": False,
            "banded_rows": False,
            "first_column": True,
        },
    )
    ws.add_table(
        "K1:L4",
        {"name": "NoHeader", "data": data, "header_row": False, "style": "Table Style Light 9"},
    )


def conditional(wb: xlsxwriter.Workbook) -> None:
    """One rule of each kind Excel 2007 had, data bars in Excel 2010's form, and stacking."""
    ws = wb.add_worksheet("Rules")
    for r in range(10):
        ws.write_row(r, 0, [r + 1, 10 - r, (r * 7) % 5, f"item {r % 3}", r * 1.5])
    red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True})
    green = wb.add_format({"bg_color": "#C6EFCE", "font_size": 14})
    fmt = ws.conditional_format
    fmt("A1:A10", {"type": "cell", "criteria": ">", "value": 5, "format": red})
    fmt("A1:A10", {"type": "cell", "criteria": "between", "minimum": 2, "maximum": 4,
                   "format": green, "stop_if_true": True})  # fmt: skip
    fmt("B1:B10", {"type": "3_color_scale"})
    fmt("C1:C10", {"type": "data_bar", "data_bar_2010": True, "bar_negative_color": "#FF0000"})
    fmt("E1:E10", {"type": "data_bar", "data_bar_2010": False})
    fmt("C1:C10", {"type": "icon_set", "icon_style": "3_traffic_lights"})
    fmt("D1:D10", {"type": "text", "criteria": "containing", "value": "item 1", "format": red})
    fmt("D1:D10", {"type": "duplicate", "format": green})
    fmt("E1:E10", {"type": "top", "value": 3, "format": red})
    fmt("E1:E10", {"type": "average", "criteria": "above", "format": green})
    fmt("A1:E10", {"type": "formula", "criteria": "=MOD(ROW(),4)=0", "format": green})
    fmt("F1:F10", {"type": "blanks", "format": red})


def drawings(wb: xlsxwriter.Workbook) -> None:
    """Pictures under each anchoring, a text box, a chart."""
    ws = wb.add_worksheet("Drawings")
    ws.write_column("A1", [3, 5, 2, 8])
    image = io.BytesIO(png(40, 20, (30, 90, 200)))
    for row, position in ((1, 1), (8, 2), (15, 3)):
        ws.insert_image(
            row, 2, f"blue{position}.png",
            {"image_data": image, "object_position": position, "x_offset": 5,
             "y_offset": 3, "description": f"anchor {position}"},
        )  # fmt: skip
    ws.insert_textbox("F2", "A text box", {"width": 160, "height": 60, "object_position": 1})
    chart = wb.add_chart({"type": "column"})
    chart.add_series({"values": "=Drawings!$A$1:$A$4"})
    ws.insert_chart("F8", chart, {"x_scale": 0.5, "y_scale": 0.5})


def notes(wb: xlsxwriter.Workbook) -> None:
    """Notes: one shown, one hidden, one with its own size and author."""
    ws = wb.add_worksheet("Notes")
    ws.write_column("A1", ["shown", "hidden", "sized"])
    ws.write_comment("B1", "A note that is shown", {"visible": True, "author": "Writer"})
    ws.write_comment("B2", "A hidden note", {"author": "Writer"})
    ws.write_comment("B3", "A wide note", {"x_scale": 2, "y_scale": 0.5, "color": "#C6EFCE"})


FIXTURES: dict[str, Callable[[xlsxwriter.Workbook], None]] = {
    "tables": tables,
    "conditional": conditional,
    "drawings": drawings,
    "notes": notes,
}


def strip(data: bytes) -> bytes:
    """Excel's copy without ``x15ac:absPath`` and ``cp:lastModifiedBy``; stops if a path is left."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            body = src.read(info)
            body = re.sub(rb"<x15ac:absPath\b[^>]*/>", b"", body)
            body = re.sub(rb"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", b"", body)
            body = re.sub(rb"<cp:lastModifiedBy\s*/>", b"", body)
            if HOME_PATH.search(body):
                raise SystemExit(f"{info.filename}: a home-directory path is left")
            dst.writestr(info, body, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


SVG = (
    b'<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96">'
    b'<rect x="8" y="8" width="80" height="80" fill="#3C8C5A"/>'
    b'<circle cx="48" cy="48" r="24" fill="#F0F0F0"/></svg>'
)
"""A picture drawn here, to stand in for any SVG a hand-made workbook holds (an icon)."""
DESCRIPTION = b"A square drawn for the tests"
DESCR = re.compile(rb'(<xdr:cNvPr\b[^>]*?)\bdescr="[^"]*"')
GRAPHIC = re.compile(rb'(<xdr:cNvPr\b[^>]*?)\bname="Graphic (\d+)"')


def anonymous(data: bytes, replace: list[str]) -> bytes:
    """People renamed, the creator dropped, pictures replaced with ones drawn here.

    Every SVG part becomes :data:`SVG` and every named PNG a solid one of the same size (a
    picture could identify someone or be another's artwork, such as an icon and its PNG
    fallback); the drawings' picture names and descriptions say what the picture now is.
    """
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            body = src.read(info)
            if info.filename.startswith("xl/persons/"):
                body = re.sub(rb'displayName="[^"]*"', b'displayName="Writer"', body)
                body = re.sub(rb'userId="[^"]*"', b'userId="Writer"', body)
                body = re.sub(rb'providerId="[^"]*"', b'providerId="None"', body)
            body = re.sub(rb"<dc:creator>[^<]*</dc:creator>", b"", body)
            if info.filename in replace:
                width, height = struct.unpack(">II", body[16:24])
                body = png(width, height, (120, 160, 90))
            if info.filename.endswith(".svg"):
                body = SVG
            if re.match(r"xl/drawings/drawing\d+\.xml$", info.filename):
                body = DESCR.sub(rb'\1descr="' + DESCRIPTION + b'"', body)
                body = GRAPHIC.sub(rb'\1name="Picture \2"', body)
            dst.writestr(info, body, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


def adopt(source: Path, name: str, replace: list[str]) -> None:
    if source.parent.resolve() != CONTAINER.resolve():
        raise SystemExit(f"only a workbook in {CONTAINER} is adopted")
    target = OUTPUT / f"{name}.xlsx"
    OUTPUT.mkdir(parents=True, exist_ok=True)
    target.write_bytes(strip(anonymous(source.read_bytes(), replace)))
    print(f"{source.name}: wrote {target.relative_to(ROOT)}")


def main(names: list[str]) -> None:
    refuse_open_workbooks()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    OUTPUT.mkdir(parents=True, exist_ok=True)
    for name in names or list(FIXTURES):
        src, copy = CONTAINER / f"xlio_fixture_{name}.xlsx", CONTAINER / f"xlio_{name}.xlsx"
        try:
            wb = xlsxwriter.Workbook(str(src))
            FIXTURES[name](wb)
            wb.close()
            copy.unlink(missing_ok=True)
            error = save_copy(src, copy)
            if error:
                print(f"{name}: Excel refused it ({error})")
                continue
            target = OUTPUT / f"{name}.xlsx"
            target.write_bytes(strip(copy.read_bytes()))
            print(f"{name}: wrote {target.relative_to(ROOT)}")
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)


def rework(name: str, replace: list[str]) -> None:
    """Apply the adopting steps again to a committed hand-made workbook, in place."""
    target = OUTPUT / f"{name}.xlsx"
    target.write_bytes(strip(anonymous(target.read_bytes(), replace)))
    print(f"reworked {target.relative_to(ROOT)}")


if __name__ == "__main__":
    if sys.argv[1:2] in (["--adopt"], ["--rework"]):
        command, *rest = sys.argv[1:]
        media = [p for p in rest if p != "--replace-media"]
        if command == "--adopt":
            file, name, *media = media
            adopt(Path(file), name, media)
        else:
            name, *media = media
            rework(name, media)
    else:
        main(sys.argv[1:])
