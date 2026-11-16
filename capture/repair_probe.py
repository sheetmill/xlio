"""What Excel keeps when it repairs a workbook: reversed anchors, invalid conditional formats.

    uv run python -m capture.repair_probe open NAME     # write NAME and open it as a person does
    uv run python -m capture.repair_probe save NAME     # save the open (repaired) workbook, read it
    uv run python -m capture.repair_probe write         # write tests/captures/repair.tsv

These need you at Excel: each case raises (or may raise) Excel's repair dialog. ``open`` writes
the case into Excel's container and opens it through Launch Services, then exits, so nothing is
sent to Excel while a dialog may be open. Answer the dialog with "Yes", copy the repair
report's text and close the report; once the workbook is open, run ``save``, which
saves it as a copy through AppleScript (``save_open.applescript``) under whatever name Excel
gave it (on 16.115 a repaired workbook kept its own name), reads what the copy keeps and deletes
the files. The repair report's text, which no script can read, is given with ``--report TEXT``
(the recovery log Excel shows); only its records are kept, paths removed. ``write`` gathers the
cases read so far into the TSV.

The control (XlsxWriter's): A1:A5 = 1..5; a picture anchored over C2 (two-cell, oneCell edit)
and a text box over F2:G4; a ``cellIs`` rule on A1:A5 with a differential fill. Each case
changes one thing:

- ``anchor ...``: the picture's ``from`` after its ``to`` on both axes, on rows only, on columns
  only; ``from`` equal to ``to``; within one cell, the ``from`` offset past the ``to`` offset;
- ``cf ...``: a formula that does not parse, an ``sqref`` that is not a range, a rule ``type``
  outside the schema, a ``dxfId`` past the table.

Each row gives: whether a repair was needed (the workbook's name in Excel), the drawing's
anchors as the copy writes them (or ``absent``), the conditional formats as the copy writes them,
and the report text.
"""

from __future__ import annotations

import datetime
import io
import json
import re
import subprocess
import sys
import zipfile

import xlsxwriter

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.fixtures import png
from capture.probe import CAPTURES, CONTAINER, HERE, TIME_LIMIT, Case, build, read

DRAWING = "xl/drawings/drawing1.xml"
SHEET = "xl/worksheets/sheet1.xml"
PICTURE_ANCHOR = (
    "<xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>381000"
    "</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
)
RULE = '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>2</formula>'
SQREF = '<conditionalFormatting sqref="A1:A5">'
STATE = HERE / "build" / "repair_cases.json"  # the cases read so far (capture/build: not committed)


def anchor(fc: int, fco: int, fr: int, fro: int, tc: int, tco: int, tr: int, tro: int) -> str:
    return (
        f"<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco}</xdr:colOff><xdr:row>{fr}</xdr:row>"
        f"<xdr:rowOff>{fro}</xdr:rowOff></xdr:from><xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>"
        f"{tco}</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>"
    )


CASES = {
    "control": Case("control"),
    "anchor reversed both": Case("", ((DRAWING, PICTURE_ANCHOR, anchor(4, 0, 6, 0, 2, 0, 1, 0)),)),
    "anchor reversed rows": Case("", ((DRAWING, PICTURE_ANCHOR, anchor(2, 0, 6, 0, 4, 0, 1, 0)),)),
    "anchor reversed columns": Case(
        "", ((DRAWING, PICTURE_ANCHOR, anchor(4, 0, 1, 0, 2, 0, 6, 0)),)
    ),
    "anchor from equal to": Case(
        "", ((DRAWING, PICTURE_ANCHOR, anchor(2, 0, 1, 0, 2, 0, 1, 0)),)
    ),
    "anchor reversed offset in one cell": Case(
        "", ((DRAWING, PICTURE_ANCHOR, anchor(2, 300000, 1, 100000, 2, 100000, 1, 50000)),)
    ),
    "cf formula that does not parse": Case(
        "", ((SHEET, "<formula>2</formula>", "<formula>2+</formula>"),)
    ),
    "cf sqref not a range": Case(
        "", ((SHEET, SQREF, '<conditionalFormatting sqref="A1:">'),)
    ),
    "cf unknown type": Case(
        "", ((SHEET, RULE, RULE.replace('type="cellIs"', 'type="bogus"')),)
    ),
    "cf dxfId past the table": Case(
        "", ((SHEET, RULE, RULE.replace('dxfId="0"', 'dxfId="99"')),)
    ),
}  # fmt: skip


def control() -> bytes:
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("Data")
    ws.write_column("A1", [1, 2, 3, 4, 5])
    red = wb.add_format({"bg_color": "#FFC7CE"})
    ws.insert_image("C2", "blue.png", {"image_data": io.BytesIO(png(40, 20, (30, 90, 200)))})
    ws.insert_textbox("F2", "box", {"width": 120, "height": 40})
    ws.conditional_format("A1:A5", {"type": "cell", "criteria": ">", "value": 2, "format": red})
    wb.close()
    return out.getvalue()


def file_name(name: str) -> str:
    return "xlio_repair_" + re.sub(r"\W+", "_", name).strip("_") + ".xlsx"


def open_case(name: str) -> None:
    refuse_open_workbooks()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / file_name(name)
    src.write_bytes(build(control(), CASES[name]))
    subprocess.run(["open", "-a", "Microsoft Excel", str(src)], check=True, timeout=30)
    print(f"{name}: opened {src.name}; answer Excel's dialog, copy the report, then run save")


def records(report: str) -> str:
    """The repair report's records (what was repaired or removed), without paths or names."""
    found = re.findall(
        r"<(repairedRecord|removedRecord|removedPart|repairedPart|removedFeature)>(.*?)</\1>",
        report,
        re.S,
    )
    text = " | ".join(f"{kind}: {body.strip()}" for kind, body in found) if found else report
    return re.sub(r"/Users/[^'\s<]*", "<path>", text).replace("\t", " ")


def save_case(name: str, report: str) -> None:
    src = CONTAINER / file_name(name)
    copy = CONTAINER / ("copy_" + file_name(name))
    names = subprocess.run(
        ["gtimeout", "-k", "5", "20", "osascript", "-e",
         'tell application "Microsoft Excel" to get name of workbooks'],
        capture_output=True, text=True, check=True,
    ).stdout.strip()  # fmt: skip
    open_names = [n.strip() for n in names.split(",") if n.strip()]
    if len(open_names) != 1:
        raise SystemExit(f"expected one open workbook, Excel has: {open_names}")
    (workbook,) = open_names
    try:
        result = subprocess.run(
            ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(HERE / "save_open.applescript"),
             workbook, str(copy)],
            capture_output=True, text=True, check=False,
        )  # fmt: skip
        if result.returncode != 0:
            raise SystemExit(f"could not save {workbook}: {result.stderr.strip()[:120]}")
        with zipfile.ZipFile(copy) as z:
            drawing = next(
                (n for n in z.namelist() if re.match(r"xl/drawings/drawing\d+\.xml$", n)), None
            )
            anchors = (
                read(z, drawing, r"<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?"
                     r"</xdr:(?:to|ext)>") if drawing else "absent"
            )  # fmt: skip
            pictures = len(re.findall(r"<xdr:pic>", z.read(drawing).decode())) if drawing else 0
            rules = read(z, SHEET, r"<conditionalFormatting\b.*?</conditionalFormatting>")
            dxfs = read(z, "xl/styles.xml", r'<dxfs count="(\d+)"')
    finally:
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    state[name] = [workbook, str(pictures), anchors, rules, dxfs, records(report)]
    STATE.write_text(json.dumps(state, indent=1))
    print("\t".join([name, *state[name]]))


def write() -> None:
    state = json.loads(STATE.read_text())
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True, text=True, check=True, timeout=30,
    ).stdout.strip()  # fmt: skip
    header = [
        f"# suite: repair (capture.repair_probe); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: opened through Launch Services, Excel's "
        "repair dialog answered Yes by a person, the report's text copied by the person, then "
        "the workbook saved as a copy through AppleScript",
        "# workbook: the name Excel gave it (a repaired one says so); anchors and rules as the "
        "copy writes them",
        "case\tworkbook\tpictures\tanchors\trules\tdxfs\treport",
    ]
    rows = ["\t".join([name, *state[name]]) for name in CASES if name in state]
    (CAPTURES / "repair.tsv").write_text("\n".join(header + rows) + "\n", encoding="utf-8")
    STATE.unlink()
    print(f"wrote tests/captures/repair.tsv: {len(rows)} cases")


if __name__ == "__main__":
    command, *rest = sys.argv[1:]
    if command == "open":
        open_case(" ".join(rest))
    elif command == "save":
        report = ""
        if "--report" in rest:
            i = rest.index("--report")
            rest, report = rest[:i], " ".join(rest[i + 1 :])
        save_case(" ".join(rest), report)
    elif command == "write":
        write()
    else:
        raise SystemExit(__doc__)
