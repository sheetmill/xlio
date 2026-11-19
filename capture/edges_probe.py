"""The edges of five reading rules: huge column widths, ``baseColWidth``, font sizes past 409,
``start`` and ``end`` beside ``left`` and ``right``, and text written beside rich-text runs.

    uv run python -m capture.edges_probe build   # check that every case builds
    uv run python -m capture.edges_probe run     # open each in Excel, write the TSV

The control (openpyxl's): A1 = 1 with a thin left border (border 1), column A 20 wide, A2 an
inline string of one bold run "HH". Each case rewrites one thing: column A's width (where one
point turns to 0 between 1e300 and 1e308, and large negative widths); ``baseColWidth`` from 9
to 255; the Normal font's size from 409 to 410; border 1 as ``start``/``end`` beside
``left``/``right``, in both orders; A2 as a ``<t>`` before, after or around the run. A control
opens first and last: an outcome counts only between two that open.

``edges_open.applescript`` opens each one, reads the widths of columns A and B and A1's font
size, and saves a copy; from the copy, each row gives the elements the cases touch as Excel
writes them back. Output: ``tests/captures/edges.tsv``.
"""

from __future__ import annotations

import datetime
import re
import subprocess
import sys
import zipfile

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import (
    CAPTURES,
    CONTAINER,
    HERE,
    ROOT,
    SHEET,
    TIME_LIMIT,
    Case,
    build,
    control_workbook,
    excel_version,
    read,
)

STYLES = "xl/styles.xml"
STRINGS = "xl/sharedStrings.xml"
WIDTH = '<col width="20" customWidth="1" min="1" max="1" />'
FORMAT = '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />'
SIZE = '<sz val="11" />'
BORDER = '<border><left style="thin" /></border>'
RUN = '<is><r><rPr><b val="1" /></rPr><t>HH</t></r></is>'
BOLD = '<r><rPr><b val="1" /></rPr><t>HH</t></r>'
READS = (
    ("cols", SHEET, r"<cols>.*?</cols>"),
    ("sheetFormatPr", SHEET, r"<sheetFormatPr\b[^>]*>"),
    ("font 0", STYLES, r"<fonts\b[^>]*>\s*<font>.*?</font>"),
    ("borders", STYLES, r"<borders\b.*?</borders>"),
    ("A1", SHEET, r'<c r="A1"[^>]*>'),
    ("strings", STRINGS, r"<si>.*?</si>"),
)
LIVE = ("width A", "width B", "A1 font size")

WIDTHS = (
    "1e30", "1e100", "1e300", "1e301", "1e302", "1e303", "1e304", "1e305", "7e305", "8e305",
    "1e306", "1e307", "2e307", "3e307", "4e307", "1e308",
    "-1e9", "-1e30", "-1e100", "-1e300", "-1e308",
)  # fmt: skip
BASES = ("9", "10", "12", "20", "50", "100", "200", "254", "255")
SIZES = ("409", "409.5", "409.51", "409.54", "409.55", "409.56", "409.6", "409.9", "410")
BORDERS = {
    "start": '<border><start style="thick" /></border>',
    "end": '<border><end style="thick" /></border>',
    "left, start": '<border><left style="thin" /><start style="thick" /></border>',
    "start, left": '<border><start style="thick" /><left style="thin" /></border>',
    "right, end": '<border><right style="dashed" /><end style="thick" /></border>',
    "end, right": '<border><end style="thick" /><right style="dashed" /></border>',
}
TEXTS = {
    "t, r": f'<is><t xml:space="preserve">aa </t>{BOLD}</is>',
    "r, t": f'<is>{BOLD}<t xml:space="preserve"> aa</t></is>',
    "t, r, t": f'<is><t xml:space="preserve">aa </t>{BOLD}<t xml:space="preserve"> bb</t></is>',
}


def fill(wb: object) -> None:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import Border, Side

    ws = wb.active  # type: ignore[attr-defined]
    ws["A1"] = 1
    ws["A1"].border = Border(left=Side(style="thin"))
    ws["A2"] = CellRichText([TextBlock(InlineFont(b=True), "HH")])
    ws.column_dimensions["A"].width = 20


def cases() -> list[Case]:
    out = [Case("control", reads=READS)]
    for value in WIDTHS:
        new = WIDTH.replace('width="20"', f'width="{value}"')
        out.append(Case(f"width {value}", ((SHEET, WIDTH, new),), reads=READS))
    for value in BASES:
        new = FORMAT.replace('baseColWidth="8"', f'baseColWidth="{value}"')
        out.append(Case(f"baseColWidth {value}", ((SHEET, FORMAT, new),), reads=READS))
    for value in SIZES:
        edits = ((STYLES, SIZE, f'<sz val="{value}" />'),)
        out.append(Case(f"font size {value}", edits, reads=READS))
    for name, border in BORDERS.items():
        out.append(Case(f"border {name}", ((STYLES, BORDER, border),), reads=READS))
    for name, text in TEXTS.items():
        out.append(Case(f"text {name}", ((SHEET, RUN, text),), reads=READS))
    out.append(Case("control", reads=READS))
    return out


def _open(src: str, copy: str) -> tuple[str, list[str]]:
    """('' and the live reads) or (the error, [])."""
    result = subprocess.run(
        ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(HERE / "edges_open.applescript"),
         src, copy],
        capture_output=True, text=True, check=False,
    )  # fmt: skip
    if result.returncode == 0:
        return "", result.stdout.strip().split("\t")
    subprocess.run(
        ["gtimeout", "-k", "5", "20", "osascript", str(HERE / "restore_alerts.applescript")],
        check=False,
    )
    found = re.search(r"\((-?\d+)\)", result.stderr)
    return (found.group(1) if found else result.stderr.strip()[:80]), []


def run() -> None:
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src, copy = CONTAINER / "xlio_edges_src.xlsx", CONTAINER / "xlio_edges_copy.xlsx"
    control = control_workbook(fill)
    labels = [*LIVE, *(label for label, _, _ in READS)]
    rows = []
    try:
        for case in cases():
            src.write_bytes(build(control, case))
            copy.unlink(missing_ok=True)
            error, live = _open(str(src), str(copy))
            if error:
                rows.append("\t".join([case.name, "refused", error] + [""] * len(labels)))
            else:
                with zipfile.ZipFile(copy) as z:
                    values = [read(z, part, pattern) for _, part, pattern in case.reads]
                rows.append("\t".join([case.name, "opened", "", *live, *values]))
            print(rows[-1], flush=True)
    finally:
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
    if not (rows[0].split("\t")[1] == rows[-1].split("\t")[1] == "opened"):
        raise SystemExit("a control did not open: the run is not evidence; nothing written")
    header = [
        f"# suite: edges (capture); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: openpyxl's workbook with one thing "
        "rewritten in the XML, opened with alerts off, read (column widths in characters, A1's "
        "font size) and saved as a copy",
        "# elements as Excel's copy writes them, matches joined with |; refused = error -50 "
        "(Excel would repair the file)",
        "\t".join(["case", "outcome", "error", *labels]),
    ]
    out = CAPTURES / "edges.tsv"
    write_capture(out, header + rows)
    print(f"wrote {out.relative_to(ROOT)}: {len(rows)} cases")


def main() -> None:
    control = control_workbook(fill)
    for case in cases():
        build(control, case)  # every edit's anchor occurs once
    if sys.argv[1:] == ["run"]:
        run()
    else:
        print(f"{len(cases())} cases build")


if __name__ == "__main__":
    main()
