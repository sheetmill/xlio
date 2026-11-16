"""What Excel repairs, read as the repaired workbook has it; refused by ``check()``.

Capture ``repair`` (``capture/repair_probe.py``): each case was opened by a person, Excel's repair
dialog answered Yes, and the workbook saved; the TSV holds its anchors and conditional formats
and the repair report's records. Here a case Excel repaired must be refused by ``check()``
(Excel refuses it with alerts off), one it opened as it was must pass, and the drawing's first
anchor and the conditional formats must read as the repaired copy has them.
"""

from __future__ import annotations

import io
import re
from pathlib import Path

import pytest

import xlio

pytest.importorskip("xlsxwriter")
from capture.probe import build
from capture.repair_probe import CASES, control

CAPTURE = Path(__file__).with_name("captures") / "repair.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = [
    dict(zip(HEADER, line.split("\t"), strict=False))
    for line in LINES
    if line.strip() and not line.startswith(("#", "case\t"))
]
CONTROL = control()
# Excel removes a conditional format whose formula does not parse; xlio does not parse formulas.
NOT_REFUSED = {"cf formula that does not parse"}


def first_anchor(anchors: str) -> tuple[int, ...]:
    numbers = re.findall(r"<xdr:(?:col|colOff|row|rowOff)>(\d+)<", anchors.split("|")[0])
    return tuple(int(n) for n in numbers)


def test_every_case_is_captured():
    assert sorted(r["case"] for r in ROWS) == sorted(CASES)


@pytest.mark.parametrize("row", ROWS, ids=[r["case"] for r in ROWS])
def test_read_as_excel_repairs_it(row):
    data = build(CONTROL, CASES[row["case"]])
    if row["report"] and row["case"] not in NOT_REFUSED:
        with pytest.raises(xlio.PackageError):
            xlio.check(io.BytesIO(data))
    else:
        xlio.check(io.BytesIO(data))
    sheet = xlio.open(io.BytesIO(data)).sheets[0]
    assert sheet.drawing is not None
    anchor = sheet.drawing.objects[0].anchor
    assert anchor.from_ is not None
    assert anchor.to is not None
    corners = [(m.column, m.column_offset, m.row, m.row_offset) for m in (anchor.from_, anchor.to)]
    assert (*corners[0], *corners[1]) == first_anchor(row["anchors"])
    if row["case"] in NOT_REFUSED:
        return
    kept = [] if row["rules"] == "-" else re.findall(r"<cfRule\b[^>]*>", row["rules"])
    rules = [r for cf in sheet.conditional_formats for r in cf.rules]
    if row["case"] == "cf unknown type":
        assert [r.type for r in rules] == [None]  # read, and refused by check()
        return
    assert len(rules) == len(kept)
    for rule, saved in zip(rules, kept, strict=True):
        assert (rule.dxf is not None) == ("dxfId=" in saved)
