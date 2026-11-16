"""What Excel takes from attribute values it never writes, one value per workbook.

    uv run python -m capture.attributes_probe

The control (openpyxl's, its custom palette removed): A1 = 1; B1 = 2 rotated 45 degrees (cell
format 1); row 1 30 pt high; column B 20 wide; A2 an inline rich string whose first run is
14 pt; A3 under a conditional format whose differential format is 12 pt. Each case rewrites one
attribute in the sheet, styles or workbook XML to a value Excel does not write: flags that are
not booleans, numbers that do not parse, indexes past their tables, colors and sizes out of
range, page setup values, an absent ``pageMargins``, print names.

Each row gives, from Excel's copy, the elements the cases touch, as Excel writes them back
(what Excel took from the value, or ``-`` when it dropped the element). Output:
``tests/captures/attributes.tsv``.
"""

from __future__ import annotations

from capture.probe import SHEET, Case, build, control_workbook, part_text, run

STYLES = "xl/styles.xml"
WORKBOOK = "xl/workbook.xml"
ROW1 = '<row r="1" ht="30" customHeight="1">'
COL = '<col width="20" customWidth="1" min="2" max="2" />'
A1 = '<c r="A1" t="n">'
FONT_COLOR = '<color theme="1" />'
FONT_SIZE = '<font><name val="Calibri" /><family val="2" /><color theme="1" /><sz val="11" />'
ROTATION = '<alignment textRotation="45" />'
RUN_SIZE = '<sz val="14" />'
DXF_SIZE = '<dxf><font><sz val="12" /></font></dxf>'
FORMAT = '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />'
MARGINS = '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" />'
NAMES = "<definedNames />"
READS = (
    ("row 1", SHEET, r'<row r="1"[^>]*>'),
    ("cols", SHEET, r"<cols>.*?</cols>"),
    ("A1", SHEET, r'<c r="A1"[^>]*>'),
    ("font 0", STYLES, r"<fonts\b[^>]*>\s*<font>.*?</font>"),
    ("alignment", STYLES, r"<alignment\b[^>]*>"),
    ("cellXfs", STYLES, r"<cellXfs\b[^>]*>"),
    ("A2 runs", "xl/sharedStrings.xml", r"<sz\b[^>]*>"),
    ("A2 inline", SHEET, r'<c r="A2".*?</c>'),
    ("dxfs", STYLES, r"<dxfs\b.*?</dxfs>"),
    ("sheetFormatPr", SHEET, r"<sheetFormatPr\b[^>]*>"),
    ("pageMargins", SHEET, r"<pageMargins\b[^>]*>"),
    ("pageSetup", SHEET, r"<pageSetup\b[^>]*>"),
    ("definedNames", WORKBOOK, r"<definedNames>.*?</definedNames>"),
)

# name -> (part, old, new)
EDITS: dict[str, tuple[str, str, str]] = {}


def edit(name: str, part: str, old: str, new: str) -> None:
    EDITS[name] = (part, old, new)


for value in ["yes", "2", "TRUE", "on", ""]:
    edit(
        f"row hidden={value!r}",
        SHEET,
        ROW1,
        f'<row r="1" ht="30" customHeight="1" hidden="{value}">',
    )
for value in ["2", "yes", "0"]:
    edit(f"row customHeight={value!r}", SHEET, ROW1, f'<row r="1" ht="30" customHeight="{value}">')
for value in ["", "abc", "30pt", " 30", "0x1E"]:
    edit(f"row ht={value!r}", SHEET, ROW1, f'<row r="1" ht="{value}" customHeight="1">')
for value in ["abc", "", "20pt"]:
    edit(
        f"col width={value!r}",
        SHEET,
        COL,
        f'<col width="{value}" customWidth="1" min="2" max="2" />',
    )
edit(
    "col hidden='yes'",
    SHEET,
    COL,
    '<col width="20" customWidth="1" hidden="yes" min="2" max="2" />',
)
edit("col min='x'", SHEET, COL, '<col width="20" customWidth="1" min="x" max="2" />')
edit("col min > max", SHEET, COL, '<col width="20" customWidth="1" min="3" max="2" />')
for value in ["99", "2", "-1", "x", "1.0"]:
    edit(f"cell s={value!r}", SHEET, A1, f'<c r="A1" s="{value}" t="n">')
for attr in ["cm", "vm"]:
    for value in ["x", "-1", "5"]:
        edit(f"cell {attr}={value!r}", SHEET, A1, f'<c r="A1" {attr}="{value}" t="n">')
for color in [
    'theme="1" tint="2"',
    'theme="1" tint="-1.5"',
    'theme="1" tint="x"',
    'rgb="FF00"',
    'rgb="GGGGGGGG"',
    'rgb="FF0000"',
    'rgb="AFF0000FF"',
    'theme="99"',
    'theme="-1"',
    'indexed="64"',
    'indexed="65"',
    'indexed="200"',
    'auto="yes"',
]:
    edit(f"font color {color}", STYLES, FONT_COLOR, f"<color {color} />")
edit("theme color sysClr", STYLES, FONT_COLOR, '<color theme="0" />')  # the theme's lt1 or dk1
for value in ["181", "200", "254", "255", "256", "-1", "90.5", "x"]:
    edit(f"textRotation {value}", STYLES, ROTATION, f'<alignment textRotation="{value}" />')
for value in ["409", "409.5", "409.55", "410", "0", "-1", "-2", "1e400"]:
    edit(f"run sz {value}", SHEET, RUN_SIZE, f'<sz val="{value}" />')
    edit(f"dxf sz {value}", STYLES, DXF_SIZE, f'<dxf><font><sz val="{value}" /></font></dxf>')
edit("patternFill without patternType, fill 2", STYLES, '<fills count="2">', '<fills count="3">')
for value in ["255", "256", "299"]:
    edit(
        f"defaultColWidth {value}",
        SHEET,
        FORMAT,
        f'<sheetFormatPr baseColWidth="8" defaultRowHeight="15" defaultColWidth="{value}" />',
    )
edit("pageMargins absent", SHEET, MARGINS, "")
for value in [
    'copies="x"',
    'copies="0"',
    'copies="-1"',
    'scale="5"',
    'scale="401"',
    'orientation="sideways"',
    'paperSize="999"',
]:
    edit(f"pageSetup {value}", SHEET, MARGINS, MARGINS + f"<pageSetup {value} />")
edit(
    "Print_Area global",
    WORKBOOK,
    NAMES,
    '<definedNames><definedName name="_xlnm.Print_Area">Sheet!$A$1:$B$3</definedName>'
    "</definedNames>",
)
edit(
    "Print_Titles cells",
    WORKBOOK,
    NAMES,
    '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet!$A$1:$B$2'
    "</definedName></definedNames>",
)
edit(
    "Print_Titles rows (control)",
    WORKBOOK,
    NAMES,
    '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet!$1:$2'
    "</definedName></definedNames>",
)


def fill(wb: object) -> None:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Font

    ws = wb.active  # type: ignore[attr-defined]
    ws["A1"] = 1
    ws["B1"] = 2
    ws["B1"].alignment = Alignment(textRotation=45)
    ws["A2"] = CellRichText([TextBlock(InlineFont(sz=14), "big"), " small"])
    ws["A3"] = 1
    ws.conditional_formatting.add(
        "A3", CellIsRule(operator="greaterThan", formula=["0"], font=Font(size=12))
    )
    ws.column_dimensions["B"].width = 20
    ws.row_dimensions[1].height = 30


def control() -> bytes:
    """openpyxl's workbook without the custom palette it always writes."""
    data = control_workbook(fill)
    styles = part_text(data, STYLES)
    i, j = styles.index("<colors>"), styles.index("</colors>") + len("</colors>")
    return build(data, Case("palette removed", ((STYLES, styles[i:j], ""),)))


def cases() -> list[Case]:
    out = [Case("control", reads=READS)]
    for name, (part, old, new) in EDITS.items():
        edits = [(part, old, new)]
        if name.startswith("patternFill"):
            edits.append(
                (
                    STYLES,
                    '<fill><patternFill patternType="gray125" /></fill></fills>',
                    '<fill><patternFill patternType="gray125" /></fill><fill><patternFill>'
                    '<fgColor rgb="FFFF0000" /></patternFill></fill></fills>',
                )
            )
            edits.append(
                (
                    STYLES,
                    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0"',
                    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="2"',
                )
            )
        out.append(Case(name, tuple(edits), reads=READS))
    return out


def main() -> None:
    run(
        "attributes",
        control(),
        cases(),
        [label for label, _, _ in READS],
        "openpyxl's workbook (palette removed) with one attribute rewritten in the XML, opened "
        "with alerts off and saved as a copy",
        "elements as Excel's copy writes them, matches joined with |",
    )


if __name__ == "__main__":
    main()
