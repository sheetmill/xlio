"""Sheet views, the active tab and sheet states read as Excel reads them.

Capture ``views`` (``capture/views_probe.py``): each case is the probe's two-sheet workbook with
one view, workbook or order change. Excel either refused it (error -50) or opened it and saved a
copy; the TSV holds the elements the copy has. Here a refused case must be a ``PackageError``,
and an opened one must read as the copy states it: the active cell, zoom, active tab and states.
"""

from __future__ import annotations

import io
import re
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.probe import build
from capture.views_probe import cases, control_workbook, fill

CAPTURE = Path(__file__).with_name("captures") / "views.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = [
    dict(zip(HEADER, line.split("\t"), strict=False))
    for line in LINES
    if line.strip() and not line.startswith(("#", "case\t"))
]
CONTROL = control_workbook(fill, sheets=2)
BUILT = {case.name: case for case in cases()}


def attribute(element: str, name: str) -> str | None:
    found = re.search(rf'\b{name}="([^"]*)"', element)
    return found.group(1) if found else None


def test_every_case_is_captured():
    assert sorted(r["case"] for r in ROWS) == sorted(BUILT)


@pytest.mark.parametrize("row", ROWS, ids=[r["case"] for r in ROWS])
def test_read_as_excel_reads_it(row):
    data = build(CONTROL, BUILT[row["case"]])
    if row["outcome"] == "refused":
        with pytest.raises(xlio.PackageError), xlio.open(io.BytesIO(data)) as wb:
            wb.sheets[0].unparsed  # noqa: B018 - the whole structure
        return
    with xlio.open(io.BytesIO(data)) as wb:
        sheet, other = wb.sheets
        (view,) = sheet.views
        assert str(view.active_cell) == (attribute(row["selection"], "activeCell") or "A1")
        assert view.zoom_scale == int(attribute(row["sheetView"], "zoomScale") or 100)
        assert wb.views[0].active_tab == int(row["activeTab"] if row["activeTab"] != "-" else 0)
        saved = row["sheets"].split("|")[1]
        assert other.state == (attribute(saved, "state") or "visible")


def test_a_cell_reference_reads_and_prints_in_a1_form():
    assert str(xlio.CellRef(3, 2)) == "B3"
    assert str(xlio.CellRef(1_048_576, 16_384)) == "XFD1048576"
    assert xlio.CellRef(3, 2) == (3, 2)
