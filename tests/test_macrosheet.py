"""A macro sheet's cells read as Excel keeps them.

Capture ``macrosheet`` (``capture/macrosheet_probe.py``): a workbook with an Excel 4.0 macro
sheet was opened as a person opens it, its macro notice answered, and saved as a copy. Here the
cells the reader gives must be the ones Excel's copy keeps, with their formulas, for a macro
sheet and an international one in Excel's own form (root ``xm:macrosheet``).
"""

from __future__ import annotations

import io
import re
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.macrosheet_probe import CASES, fill
from capture.probe import build, control_workbook

CAPTURE = Path(__file__).with_name("captures") / "macrosheet.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = {
    row["case"]: row
    for row in (
        dict(zip(HEADER, line.split("\t"), strict=False))
        for line in LINES
        if line.strip() and not line.startswith(("#", "case\t"))
    )
}
EXCEL_FORM = ["xlsm, xlMacrosheet", "xlsm, xlIntlMacrosheet"]


def test_every_case_is_captured():
    assert sorted(ROWS) == sorted(c.name for _, c in CASES)


@pytest.mark.parametrize("name", EXCEL_FORM)
def test_cells_as_excel_keeps_them(name):
    row = ROWS[name]
    assert row["outcome"] == "opened"
    assert row["root"].startswith("<xm:macrosheet ")
    kept = re.findall(r'<c r="([A-Z]+\d+)"[^>]*>(?:<f>(.*?)</f>)?', row["sheetData"])
    case = next(c for _, c in CASES if c.name == name)
    with xlio.open(io.BytesIO(build(control_workbook(fill), case))) as wb:
        sheet = wb["Macro1"]
        assert sheet.kind == "macrosheet"
        cells = [
            (f"A{c.row}", c.formula.text.replace("&", "&amp;") if c.formula else "")
            for c in sheet.cells()
        ]
        assert sheet.columns[0].width == 20.0
    assert cells == kept
