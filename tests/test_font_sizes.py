"""Cell font sizes read as Excel for Windows gives them while the workbook is open.

The Windows capture ``tests/captures/windows/font_sizes.tsv`` (``capture/font_job.py``): each
workbook is ``edges``' with its font's size rewritten; the job read A1's ``Font.Size`` while it
was open. Excel keeps a size in twentieths of a point, cut down; a size from 409.6 is refused.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.font_job import files

CAPTURE = Path(__file__).with_name("captures") / "windows" / "font_sizes.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = {
    row["case"]: row
    for row in (
        dict(zip(HEADER, line.split("\t"), strict=True))
        for line in LINES
        if line.strip() and not line.startswith(("#", "case\t"))
    )
}
FILES = files()


def test_every_case_is_captured():
    assert set(ROWS) == set(FILES)


@pytest.mark.parametrize("case", sorted(FILES), ids=sorted(FILES))
def test_a_cell_font_size_reads_as_excel_holds_it(case):
    row, data = ROWS[case], FILES[case][1]
    if row["outcome"] == "refused":
        with pytest.raises(xlio.PackageError, match="font size"), xlio.open(io.BytesIO(data)) as wb:
            _ = wb.styles  # the styles part is read when asked for
        return
    with xlio.open(io.BytesIO(data)) as wb:
        cell = next(c for c in wb.sheets[0].cells() if (c.row, c.column) == (1, 1))
        size = wb.styles.fonts[wb.styles.cell_formats[cell.style].font_id].size
    assert size == float(row["A1 Font.Size"])
