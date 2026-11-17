"""The shortcuts of the cell stream give what the full readings give.

Reading cells counts the cell formats without reading the styles, and reads a style index
through a cache of the texts met: both must agree with the full readings, on every committed
workbook and on crafted parts.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

import xlio
from xlio._cells import _STYLES, _index, read_cells
from xlio._styles import count_cell_formats

FILES = Path(__file__).with_name("files")
WORKBOOKS = sorted(FILES.rglob("*.xls[xm]"))
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


@pytest.mark.parametrize("path", WORKBOOKS, ids=[p.name for p in WORKBOOKS])
def test_the_count_is_the_number_of_cell_formats(path):
    with xlio.open(path) as wb:
        count = wb._cell_format_count
        assert count == len(wb.styles.cell_formats)


@pytest.mark.parametrize(
    ("body", "expected"),
    [
        ("<cellXfs><xf/><xf/></cellXfs>", 2),
        ("<cellXfs><xf/><extLst/><xf><alignment/></xf></cellXfs><cellXfs><xf/></cellXfs>", 2),
        ("<cellStyleXfs><xf/></cellStyleXfs>", 0),
        ("", 0),
    ],
)
def test_the_count_takes_the_first_cell_xfs_and_its_xf_children(body, expected):
    prefixed = body.replace("<", "<x:").replace("<x:/", "</x:")  # a prefix, as some writers use
    part = f'<x:styleSheet xmlns:x="{MAIN}">{prefixed}</x:styleSheet>'
    assert count_cell_formats(io.BytesIO(part.encode()), "styles.xml") == expected


@pytest.mark.parametrize(
    "s", ["0", "7", "007", "+1", " 3", "3 ", "1e3", "-1", "", "٣", "9" * 25, "abc"]
)
def test_a_style_index_reads_as_its_full_reading(s):
    _STYLES.clear()
    for _ in range(2):  # read, then from the cache
        row = f'<row r="1"><c r="A1" s="{s}"><v>1</v></c></row>'
        xml = f'<worksheet xmlns="{MAIN}"><sheetData>{row}</sheetData></worksheet>'
        (cell,) = read_cells(io.BytesIO(xml.encode()), [])
        assert cell.style == (_index(s, None) or 0)
