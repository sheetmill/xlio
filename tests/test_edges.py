"""Every case of the ``edges`` capture, read as Excel read it.

Huge column widths, ``baseColWidth``, font sizes past 409, ``start`` and ``end`` beside ``left``
and ``right``, and text written beside rich-text runs (``capture/edges_probe.py``). The cases are
rebuilt from the probe, so the test needs openpyxl.
"""

from __future__ import annotations

import csv
import importlib.util
import io
import re
from pathlib import Path

import pytest

import xlio

CAPTURE = Path(__file__).with_name("captures") / "edges.tsv"
pytestmark = pytest.mark.skipif(
    importlib.util.find_spec("openpyxl") is None, reason="needs openpyxl"
)


def captured() -> list[dict[str, str]]:
    lines = [line for line in CAPTURE.read_text(encoding="utf-8").splitlines() if line[:1] != "#"]
    return list(csv.DictReader(lines, delimiter="\t"))


def read_everything(data: bytes) -> xlio.Workbook:
    xlio.check(io.BytesIO(data))
    wb = xlio.open(io.BytesIO(data))
    for sheet in wb.sheets:
        list(sheet.cells(rich=True))
        sheet.columns  # noqa: B018 - the structural pass
    wb.styles  # noqa: B018
    return wb


def _side(borders: str, name: str) -> str | None:
    """The style of the side Excel wrote back on border 1, or ``None``."""
    found = re.findall(rf'<{name} style="(\w+)"', borders)
    return found[0] if found else None


def test_every_edges_case_reads_as_excel_read_it():
    from capture import edges_probe
    from capture.probe import build, control_workbook

    control = control_workbook(edges_probe.fill)
    cases = edges_probe.cases()
    rows = captured()
    assert [r["case"] for r in rows] == [c.name for c in cases]
    assert rows[0]["outcome"] == rows[-1]["outcome"] == "opened"  # between two controls
    for row, case in zip(rows, cases, strict=True):
        name, data = row["case"], build(control, case)
        if row["outcome"] == "refused":
            with pytest.raises(xlio.PackageError):
                read_everything(data)
            continue
        wb = read_everything(data)
        sheet = wb.sheets[0]
        if name.startswith("width "):
            width = sheet.columns[0].width
            assert (width, float(row["width A"])) in ((0.0, 0.0), (0.1, 0.09)), name
        elif name.startswith("baseColWidth "):
            assert sheet.format.base_col_width == float(row["width B"]), name
        elif name.startswith("font size "):
            kept = float(re.findall(r'<sz val="([\d.]+)"', row["font 0"])[0])
            size = wb.styles.fonts[0].size
            assert size is not None
            assert int(size * 20 + 1e-6) / 20 == kept, name  # twentieths, cut down
        elif name.startswith("border "):
            border = wb.styles.borders[1]
            sides = (border.left and border.left.style, border.right and border.right.style)
            borders = row["borders"].split("</border>")[1]
            assert sides == (_side(borders, "left"), _side(borders, "right")), name
        elif name.startswith("text "):
            (cell,) = (c for c in sheet.cells(rich=True) if (c.row, c.column) == (2, 1))
            assert cell.rich is not None
            written = re.findall(r"<r>(.*?)</r>", row["strings"])
            runs = [(re.findall(r"<t[^>]*>(.*?)</t>", r)[0], "<b/>" in r) for r in written]
            assert [(r.text, r.font is not None and r.font.bold) for r in cell.rich.runs] == runs
