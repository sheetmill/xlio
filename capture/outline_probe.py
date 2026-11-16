"""What Excel reads from outline levels it does not write: past 7, negative, not integers.

    uv run python -m capture.outline_probe

The control (openpyxl's): one sheet with A1:A3 = 1, 2, 3. Each case writes one value into one
place: ``outlineLevel`` on row 2, ``outlineLevel`` on a ``<col>`` for column B, or the sheet's
``outlineLevelRow`` or ``outlineLevelCol`` (``sheetFormatPr``). The values: 7 (the deepest level
Excel writes, the control), 8, 9, 15, 255, 256, -1, 1.5, ``abc``, empty, `` 2``, ``02``, ``+2``.

Each row gives, from Excel's copy: row 2's element, the ``cols`` element and ``sheetFormatPr``.
Output: ``tests/captures/outline.tsv``.
"""

from __future__ import annotations

from capture.probe import SHEET, Case, control_workbook, run

ROW = '<row r="2">'
FORMAT = '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />'
DATA = "<sheetData>"
READS = (
    ("row 2", SHEET, r'<row r="2"[^>]*>'),
    ("cols", SHEET, r"<cols>.*?</cols>"),
    ("sheetFormatPr", SHEET, r"<sheetFormatPr\b[^>]*>"),
)
LABELS = [label for label, _, _ in READS]
VALUES = ["7", "8", "9", "15", "255", "256", "-1", "1.5", "abc", "", " 2", "02", "+2"]


def fill(wb: object) -> None:
    ws = wb.active  # type: ignore[attr-defined]
    for row in (1, 2, 3):
        ws.cell(row, 1, row)


def cases() -> list[Case]:
    out = []
    for value in VALUES:
        name = f"{value!r}" + (" (control)" if value == "7" else "")
        col = (
            f'<cols><col min="2" max="2" width="9" customWidth="1" outlineLevel="{value}"/></cols>'
        )
        fmt = FORMAT.replace(" />", "")
        out += [
            Case(
                f"row {name}",
                edits=((SHEET, ROW, f'<row r="2" outlineLevel="{value}">'),),
                reads=READS,
            ),
            Case(f"col {name}", edits=((SHEET, DATA, col + DATA),), reads=READS),
            Case(
                f"outlineLevelRow {name}",
                edits=((SHEET, FORMAT, f'{fmt} outlineLevelRow="{value}" />'),),
                reads=READS,
            ),
            Case(
                f"outlineLevelCol {name}",
                edits=((SHEET, FORMAT, f'{fmt} outlineLevelCol="{value}" />'),),
                reads=READS,
            ),
        ]
    return out


if __name__ == "__main__":
    run(
        "outline",
        control_workbook(fill),
        cases(),
        LABELS,
        "AppleScript open, save as .xlsx, read the copy's XML",
        "one outline level written into a row, a column or sheetFormatPr; what Excel saves",
    )
