"""Outline levels read as Excel reads them.

Capture ``outline`` (``capture/outline_probe.py``): each case writes one value into a row's or a
column's ``outlineLevel``, or into the sheet's ``outlineLevelRow`` or ``outlineLevelCol``. Excel
refused the file (error -50) or opened it and saved a copy. A row's or column's level must read
as the copy states it (none written: 0). A sheet's level Excel counts again from the rows and
columns, so its copy says nothing of the value read: there a refused case must be a
``PackageError`` (from reading the sheet's format and from ``check()``), and an opened one must
read as a level from 0 to 7.
"""

from __future__ import annotations

import io
import re
from pathlib import Path

import pytest

import xlio
from xlio._types import outline_level

pytest.importorskip("openpyxl")
from capture.outline_probe import cases, control_workbook, fill
from capture.probe import build

CAPTURE = Path(__file__).with_name("captures") / "outline.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = [
    dict(zip(HEADER, line.split("\t"), strict=False))
    for line in LINES
    if line.strip() and not line.startswith(("#", "case\t"))
]
CONTROL = control_workbook(fill)
BUILT = {case.name: case for case in cases()}


def level(element: str) -> int:
    found = re.search(r'\boutlineLevel="([^"]*)"', element)
    return int(found[1]) if found else 0


def test_every_case_is_captured():
    assert sorted(r["case"] for r in ROWS) == sorted(BUILT)


@pytest.mark.parametrize("row", ROWS, ids=[r["case"] for r in ROWS])
def test_read_as_excel_reads_it(row):
    data = build(CONTROL, BUILT[row["case"]])
    with xlio.open(io.BytesIO(data)) as wb:
        sheet = wb.sheets[0]
        place = row["case"].split(" ", 1)[0]
        if place == "row":
            assert row["outcome"] == "opened"
            assert sheet.rows[2].outline_level == level(row["row 2"])
        elif place == "col":
            assert row["outcome"] == "opened"
            (column,) = sheet.columns
            assert column.outline_level == level(row["cols"])
        elif row["outcome"] == "refused":
            with pytest.raises(xlio.PackageError, match="which Excel refuses"):
                sheet.format  # noqa: B018
        else:
            assert 0 <= sheet.format.outline_level_row <= 7
            assert 0 <= sheet.format.outline_level_col <= 7
    if row["outcome"] == "refused":
        with pytest.raises(xlio.PackageError):
            xlio.check(io.BytesIO(data))
    else:
        xlio.check(io.BytesIO(data))


def test_long_digit_strings_cost_nothing():
    assert outline_level("9" * 5000) == 7  # ...999 is 7 modulo 8
    assert outline_level("1" + "0" * 5000) == 0
