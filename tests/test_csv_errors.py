"""Error codes written as CSV fields, read as Excel shows them while the workbook is open.

Captures ``csv_errors`` (the Mac's) and ``windows/csv_errors`` (Excel for Windows, by its job;
``capture/csv_errors.py``): each field of two files (quoted and bare) read while open, its text,
``ERROR.TYPE`` and ``ISERROR``, and the cell of the copy Excel saved. A field Excel holds as an
error (``ISERROR`` TRUE) must read as that error, the code it shows; any other as the text.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from capture.csv_errors import FIELDS, FILES

import xlio
from tests.test_text import SAVED_AS

CAPTURES = Path(__file__).with_name("captures")
PLATFORMS = {"mac": CAPTURES / "csv_errors.tsv", "windows": CAPTURES / "windows/csv_errors.tsv"}


def rows(path: Path) -> list[dict[str, str]]:
    lines = path.read_text(encoding="utf-8").splitlines()
    header = next(line.split("\t") for line in lines if line.startswith("file\t"))
    return [
        dict(zip(header, line.split("\t"), strict=True))
        for line in lines
        if line.strip()
        and not line.startswith(("# ", "#\t", "file\t"))
        and not line.startswith("#   ")
    ]


@pytest.mark.parametrize("platform", list(PLATFORMS))
def test_each_field_reads_as_the_open_cell_shows_it(platform):
    captured = rows(PLATFORMS[platform])
    assert [(r["file"], r["field"]) for r in captured] == [(f, x) for f in FILES for x in FIELDS]
    for case, (_, data) in FILES.items():
        with xlio.read_csv(io.BytesIO(data), platform=platform) as wb:
            cells = {c.row: c for c in wb.sheets[0].cells()}
        for r, row in enumerate((x for x in captured if x["file"] == case), start=1):
            cell = cells[r]
            assert row["outcome"] == "opened"
            if row["ISERROR"] == "TRUE":
                assert (cell.type, cell.value) == ("e", row["shown"]), row
                assert isinstance(cell.value, xlio.ErrorCode)
                assert row["ERROR.TYPE"].isdigit()
                saved = SAVED_AS.get(row["shown"], row["shown"])
                assert row["copy"] == f"e:{saved}", row
            else:
                assert (cell.type, cell.value) == ("s", row["field"]), row
