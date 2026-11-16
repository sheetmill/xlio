"""Number forms Excel never writes, and rows and columns off its grid, read as Excel reads them.

Capture ``value_forms`` (``capture/value_probe.py``): each case is the workbook the probe built,
with A1 rewritten; Excel either refused it (error -50) or opened it, and the TSV holds A1's
value from the copy it saved. Here a refused case must be a ``PackageError``, and an opened one
must read A1 as Excel did.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.value_probe import cases

CAPTURE = Path(__file__).with_name("captures") / "value_forms.tsv"
ROWS = [
    line.split("\t")
    for line in CAPTURE.read_text(encoding="utf-8").splitlines()
    if line.strip() and not line.startswith("#") and not line.startswith("case\t")
]
BUILT = cases()


def expected(text: str) -> object:
    kind, _, value = text.partition(":")
    if text == "absent":
        return None
    if kind == "e":
        return xlio.ErrorCode(value)
    return float(value)


def a1(data: bytes) -> object:
    with xlio.open(io.BytesIO(data)) as wb:
        cells = {(c.row, c.column): c.value for c in wb.sheets[0].cells()}
    return cells.get((1, 1))


def test_every_case_is_captured():
    assert sorted(r[0] for r in ROWS) == sorted(BUILT)


@pytest.mark.parametrize("row", ROWS, ids=[r[0] for r in ROWS])
def test_read_as_excel_reads_it(row):
    case, outcome = row[0], row[1]
    data = BUILT[case]
    if case == "row 99999999999":
        # Excel opens it and drops the cell; the reader refuses a row off the sheet.
        assert outcome == "opened"
        with pytest.raises(xlio.PackageError, match="off the sheet"):
            a1(data)
        return
    if outcome == "refused":
        with pytest.raises(xlio.PackageError):
            a1(data)
        return
    assert a1(data) == expected(row[3])
