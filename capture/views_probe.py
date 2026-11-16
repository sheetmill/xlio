"""What Excel does with odd sheet views, an odd active tab, and head records after the cells.

    uv run python -m capture.views_probe

The control (openpyxl's): two sheets; on the first, A1 = 1, C3 = 3 and the selection's
``activeCell`` C3. Each case rewrites one thing in the first sheet's or the workbook's XML:

- ``activeCell``: a cell Excel writes (``C3``) beside forms it does not (lowercase, ``$C$3``, a
  range, a list, off the grid, empty, not a reference), and a selection for a pane the view does
  not have;
- ``sheetView``: a zoom outside 10 to 400 or not a number; ``workbookView``: an ``activeTab`` past
  the sheets or negative; a sheet ``state`` outside the three;
- the head out of order: ``sheetViews``, ``cols`` or ``sheetPr`` written after ``sheetData``.

Each row gives, from Excel's copy: the first sheet's ``sheetView`` and ``selection`` elements,
its ``cols``, the ``workbookView``'s ``activeTab`` and each sheet's ``state``. Output:
``tests/captures/views.tsv``.
"""

from __future__ import annotations

from capture.probe import SHEET, Case, control_workbook, run

WORKBOOK = "xl/workbook.xml"
SELECTION = '<selection activeCell="C3" sqref="A1" />'
VIEW = '<sheetView workbookViewId="0">'
VIEWS = f"<sheetViews>{VIEW}{SELECTION}</sheetView></sheetViews>"
DATA_END = "</sheetData>"
READS = (
    ("sheetView", SHEET, r"<sheetView\b[^>]*>"),
    ("selection", SHEET, r"<selection\b[^>]*>"),
    ("cols", SHEET, r"<cols>.*?</cols>"),
    ("activeTab", WORKBOOK, r'<workbookView\b[^>]*?activeTab="([^"]*)"'),
    ("sheets", WORKBOOK, r"<sheet\b[^>]*>"),
)
LABELS = [label for label, _, _ in READS]

ACTIVE_CELLS = {
    "C3 (control)": "C3",
    "lowercase": "c3",
    "absolute": "$C$3",
    "range": "C3:D4",
    "list": "C3 E5",
    "R1C1": "R3C3",
    "row 0": "A0",
    "row past the grid": "A1048577",
    "column past the grid": "XFE1",
    "empty": "",
    "not a reference": "ZZZ",
    "sheet-qualified": "Sheet!C3",
    "space before": " C3",
}
ZOOMS = ["5", "9", "401", "1000", "0", "-50", "abc", "100.5"]


def fill(wb: object) -> None:
    ws = wb.active  # type: ignore[attr-defined]
    ws["A1"] = 1
    ws["C3"] = 3
    ws.sheet_view.selection[0].activeCell = "C3"


def selection(attrs: str) -> tuple[str, str, str]:
    return (SHEET, SELECTION, f"<selection {attrs} />")


def cases() -> list[Case]:
    out = [Case("control", reads=READS)]
    # sqref B2:D4 holds C3: an activeCell Excel takes stays C3; one it drops gives way to
    # sqref's first cell (B2), or to A1.
    for name, text in ACTIVE_CELLS.items():
        out.append(
            Case(
                f"activeCell {name}",
                (selection(f'activeCell="{text}" sqref="B2:D4"'),),
                reads=READS,
            )
        )
    out.append(Case("activeCell absent, sqref B2:D4", (selection('sqref="B2:D4"'),), reads=READS))
    out.append(
        Case(
            "activeCell E5 outside sqref B2:D4",
            (selection('activeCell="E5" sqref="B2:D4"'),),
            reads=READS,
        )
    )
    out.append(Case("activeCell C3, sqref absent", (selection('activeCell="C3"'),), reads=READS))
    out.append(
        Case(
            "activeCell C3, sqref list",
            (selection('activeCell="C3" sqref="A1 C3:D4"'),),
            reads=READS,
        )
    )
    out.append(
        Case(
            "selection for an absent pane",
            (selection('pane="bottomRight" activeCell="D4" sqref="D4"'),),
            reads=READS,
        )
    )
    out.append(
        Case(
            "selection for an absent pane topRight",
            (selection('pane="topRight" activeCell="D4" sqref="D4"'),),
            reads=READS,
        )
    )
    for zoom in ZOOMS:
        out.append(
            Case(
                f"zoomScale {zoom}",
                ((SHEET, VIEW, f'<sheetView workbookViewId="0" zoomScale="{zoom}">'),),
                reads=READS,
            )
        )
    for tab in ["1", "2", "5", "-1", "x"]:
        out.append(
            Case(
                f"activeTab {tab}",
                ((WORKBOOK, 'activeTab="0"', f'activeTab="{tab}"'),),
                reads=READS,
            )
        )
    for state in ["hidden", "Hidden", "bogus", ""]:
        out.append(
            Case(
                f"Sheet2 state {state!r}",
                (
                    (
                        WORKBOOK,
                        'name="Sheet2" sheetId="2" state="visible"',
                        f'name="Sheet2" sheetId="2" state="{state}"',
                    ),
                ),
                reads=READS,
            )
        )
    moved_views = (
        f'<sheetViews>{VIEW}<selection activeCell="C3" sqref="C3" /></sheetView></sheetViews>'
    )
    out.append(
        Case(
            "sheetViews after sheetData",
            ((SHEET, VIEWS, ""), (SHEET, DATA_END, DATA_END + moved_views)),
            reads=READS,
        )
    )
    cols = '<cols><col min="2" max="2" width="30" customWidth="1" /></cols>'
    out.append(
        Case(
            "cols before sheetData (control)",
            ((SHEET, "<sheetData>", cols + "<sheetData>"),),
            reads=READS,
        )
    )
    out.append(Case("cols after sheetData", ((SHEET, DATA_END, DATA_END + cols),), reads=READS))
    out.append(
        Case(
            "sheetPr after sheetData",
            (
                (
                    SHEET,
                    '<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr />'
                    "</sheetPr>",
                    "",
                ),
                (SHEET, DATA_END, DATA_END + '<sheetPr><tabColor rgb="FFFF0000" /></sheetPr>'),
            ),
            reads=(*READS, ("tabColor", SHEET, r"<tabColor\b[^>]*>")),
        )
    )
    return out


def main() -> None:
    run(
        "views",
        control_workbook(fill, sheets=2),
        cases(),
        [*LABELS, "extra"],
        "openpyxl's two-sheet workbook with one view, workbook or order change in the XML, "
        "opened with alerts off and saved as a copy",
        "elements as Excel's copy writes them, matches joined with |",
    )


if __name__ == "__main__":
    main()
