"""Error values read as Excel reads them: plain codes, rich-value errors, refusals.

Capture ``errors`` (``capture/errors_probe.py``): each case is the probe's workbook with A1
rewritten. Excel either refused it (error -50) or opened it; for an opened one, the TSV holds the
error's text as Excel shows it (``VALUETOTEXT(A1)``). Here a refused case must be a
``PackageError``, and an opened one must read A1 as the error Excel shows, or as the number
Excel saves for it (``errorType`` 14).
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.errors_probe import cases, control_workbook, fill
from capture.probe import build

CAPTURE = Path(__file__).with_name("captures") / "errors.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = [
    dict(zip(HEADER, line.split("\t"), strict=False))
    for line in LINES
    if line.strip() and not line.startswith(("#", "case\t"))
]
CONTROL = control_workbook(fill)
BUILT = {case.name: case for case in cases()}


def a1(data: bytes) -> object:
    with xlio.open(io.BytesIO(data)) as wb:
        return next(c.value for c in wb.sheets[0].cells() if (c.row, c.column) == (1, 1))


def test_every_case_is_captured():
    assert sorted(r["case"] for r in ROWS) == sorted(BUILT)


@pytest.mark.parametrize("row", ROWS, ids=[r["case"] for r in ROWS])
def test_read_as_excel_reads_it(row):
    data = build(CONTROL, BUILT[row["case"]])
    if row["outcome"] == "refused":
        with pytest.raises(xlio.PackageError):
            a1(data)
        return
    shown = row["E1 =VALUETOTEXT(A1)"].partition(":")[2]
    value = a1(data)
    if row["case"].startswith(("formula", "cached")):
        # Excel computes the formula on open; the file's cached value is what xlio reads.
        assert value is None or isinstance(value, xlio.ErrorCode)
        return
    if row["case"] == "control":
        assert value == 1.0
        return
    saved = row["A1"]
    if saved.startswith("n:"):
        # errorType 14: Excel saves the cell as a number, and =A1 caches it; read so.
        assert value == float(saved[2:])
        return
    assert isinstance(value, xlio.ErrorCode)
    assert value == shown


def test_every_error_excel_shows_is_an_error_code():
    shown = {r["E1 =VALUETOTEXT(A1)"].partition(":")[2] for r in ROWS if r["outcome"] == "opened"}
    errors = {s for s in shown if s.startswith("#")}
    assert errors <= {e.value for e in xlio.ErrorCode}
    assert {e.value for e in xlio.ErrorCode} - errors == {"#EXTERNAL!", "#PYTHON!"}
