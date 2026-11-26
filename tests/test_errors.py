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


# Capture ``errors_stale``: a formula cell whose vm leads to an error rich value (#CALC!) while
# its <v> caches something else, planted in a file Excel saved (so it keeps cached values).
# Excel shows the rich value's error whatever the cell caches; without the vm, the cached value.
from capture.errors_probe import A1, STALE, STALE_FORMULA  # noqa: E402
from capture.probe import SHEET, Case  # noqa: E402

STALE_ROWS = [
    dict(zip(HEADER, line.split("\t"), strict=False))
    for line in (Path(__file__).with_name("captures") / "errors_stale.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if line.strip() and not line.startswith(("# ", "case\t"))  # a case may start with "#"
]


def test_every_stale_case_is_captured():
    names = {r["case"] for r in STALE_ROWS}
    assert set(STALE) | {f"{n}, B1:E1 uncached" for n in STALE} <= names
    assert all(r["outcome"] == "opened" for r in STALE_ROWS)


@pytest.mark.parametrize("name", list(STALE))
def test_an_error_rich_value_wins_over_the_cached_value(name):
    shown = next(r for r in STALE_ROWS if r["case"] == f"{name}, B1:E1 uncached")
    kind, value = STALE[name]
    vm = "" if name.endswith("(control)") else ' vm="1"'
    t = "" if kind == "n" else f' t="{kind}"'
    rich = BUILT["rich 13"]  # the parts of a #CALC! rich value (errorType 13)
    cell = f'<c r="A1"{t}{vm}><f>{STALE_FORMULA}</f><v>{value}</v></c>'
    case = Case(name, ((SHEET, A1, cell), *rich.edits[1:]), rich.parts)
    assert a1(build(CONTROL, case)) == shown["E1 =VALUETOTEXT(A1)"].removeprefix("str:")


# Capture ``errors_shown`` (``capture/errors_shown.py``): what A1:E1 show while the workbook is
# open, what Excel saves in A1, and what its saved copy shows. For errorType 14 the cell shows
# #CALC! while Excel saves the number 0; xlio reads the value the file stores (0), as the copy
# then shows it, and the cases whose error Excel keeps read as that error.
SHOWN_CAPTURE = Path(__file__).with_name("captures") / "errors_shown.tsv"
SHOWN_LINES = SHOWN_CAPTURE.read_text(encoding="utf-8").splitlines()
SHOWN_HEADER = next(line.split("\t") for line in SHOWN_LINES if line.startswith("case\t"))
SHOWN_ROWS = {
    row["case"]: row
    for row in (
        dict(zip(SHOWN_HEADER, line.split("\t"), strict=True))
        for line in SHOWN_LINES
        if line.strip() and not line.startswith(("# ", "case\t"))
    )
}


def test_what_excel_shows_and_saves_for_each_rich_error():
    assert set(SHOWN_ROWS) == {"control", "rich 8", "rich 13", "rich 14"}
    for name, row in SHOWN_ROWS.items():
        value = a1(build(CONTROL, BUILT[name]))
        if name == "rich 14":
            assert (row["A1 shown"], row["copy A1"], row["reopened A1 shown"]) == (
                "#CALC!",
                "n:0",
                "0",
            )
            assert value == 0.0  # the stored value; the cell shows #CALC! until Excel saves it
        elif name == "control":
            assert value == float(row["A1 shown"]) == 1.0
        else:
            assert value == row["A1 shown"] == row["reopened A1 shown"]
