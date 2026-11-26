"""Readings the earlier captures took from Excel's saved copy, read while the workbook is open.

    uv run python -m capture.live_reads           # tests/captures/live_reads.tsv
    uv run python -m capture.live_reads current   # what xlio reads for each case

The ``edges``, ``attributes``, ``outline`` and ``value_forms`` captures read some rules from the
copy Excel saves, which a saver may have changed. Here the same cases (rebuilt by those probes)
are opened and read through AppleScript while the workbook is open (``live_reads.applescript``),
each read given as ``class:value`` so that a whole number AppleScript gives for a size is told
apart from Excel's own rounding:

- font sizes in twentieths of a point (``edges``' cell font, ``attributes``' rich-text run), with
  12.5 as a control of AppleScript's rounding, and a second read of the cell's size that is
  not rounded to a whole number (Excel's dictionary types ``font size`` as an integer):
  ``GET.CELL(19)`` through ``evaluate`` and ``run XLM Macro`` (not through a defined name),
  and row 1's fitted height;
- run and differential (conditional format) font sizes out of range (``attributes``);
- outline levels past 7 (``outline``);
- rows and columns with attribute values Excel never writes (``attributes``);
- a subnormal ``<v>`` (``value_forms``' A1 = 111, B1 = ``A1*2``), with C1 = ``A1>0`` and D1 =
  ``A1*1E+300`` computed on opening.

Runs where the other probes run, and in a macOS VM (no ``gtimeout``: the time limit is
``subprocess``'s own, and alerts are restored when it kills osascript).
"""

from __future__ import annotations

import datetime
import io
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from capture import attributes_probe, edges_probe, outline_probe, value_probe
from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import (
    CAPTURES,
    CONTAINER,
    HERE,
    SHEET,
    TIME_LIMIT,
    Case,
    build,
    control_workbook,
    excel_version,
)

SUITE = "live_reads"
SHEET_REF = "of active sheet"
COLUMNS = ["case", "outcome", "error", "read", "value"]


@dataclass(frozen=True)
class Live:
    """A workbook and the expressions read from it while it is open."""

    name: str
    data: bytes
    reads: tuple[tuple[str, str], ...]  # (label, AppleScript expression)


def _r(ref: str) -> str:
    return f'range "{ref}" {SHEET_REF}'


CELL_SIZE = (("A1 font size", f"font size of font object of {_r('A1')}"),)
CELL_SIZE_XLM = (
    ("A1 GET.CELL(19) by evaluate", 'evaluate name "GET.CELL(19,Sheet!$A$1)"'),
    ("A1 GET.CELL(19) by run XLM Macro", 'run XLM Macro "GET.CELL(19,!R1C1)"'),
    ("row 1 height", f"row height of {_r('A1')}"),
)
_FONT = "font size of font object of"
RUN_SIZE = (
    ("A2 character 1 font size", f"{_FONT} character 1 of {_r('A2')}"),
    ("A2 character 5 font size", f"{_FONT} character 5 of {_r('A2')}"),
)
DXF_SIZE = (("A3 condition 1 font size", f"{_FONT} format condition 1 of {_r('A3')}"),)
ROW_COL = (
    ("row 1 height", f"row height of {_r('A1')}"),
    ("row 1 hidden", f"hidden of entire row of {_r('A1')}"),
    ("column B width", f"column width of {_r('B1')}"),
    ("column B hidden", f"hidden of entire column of {_r('B1')}"),
)
OUTLINE = (
    ("row 2 outline level", f"outline level of entire row of {_r('A2')}"),
    ("column B outline level", f"outline level of entire column of {_r('B1')}"),
)
VALUE = tuple(
    (f"{ref} {what}", f"{what} of {_r(ref)}")
    for ref in ("A1", "B1", "C1", "D1")
    for what in ("value", "string value")
)

CELL_SIZES = ("12", "12.3", "12.34", "12.35", "12.36", "12.4", "12.5", "12.6", "13", "409",
              "409.5", "409.54", "409.55", "409.56")  # fmt: skip
RUN_CASES = ("control", "run sz 409.5", "run sz 409.55", "run sz 410", "run sz 0", "run sz -1",
             "run sz -2", "run sz 1e400")  # fmt: skip
DXF_CASES = ("control", "dxf sz 409.55", "dxf sz -1", "dxf sz -2")
ROW_COL_CASES = (
    "control", "row hidden='yes'", "row hidden='2'", "row hidden='TRUE'", "row hidden='on'",
    "row hidden=''", "row customHeight='2'", "row customHeight='yes'", "row customHeight='0'",
    "row ht=''", "row ht='abc'", "row ht='30pt'", "row ht=' 30'", "row ht='0x1E'",
    "col width='abc'", "col width=''", "col width='20pt'", "col hidden='yes'",
)  # fmt: skip
OUTLINE_VALUES = ("7", "8", "9", "15", "255", "256")
SUBNORMALS = ("111", "4.9e-324", "1e-310", "2.2250738585072014e-308")


def _edges_cases() -> list[Live]:
    control = control_workbook(edges_probe.fill)
    out = []
    for size in CELL_SIZES:
        sized = (edges_probe.STYLES, edges_probe.SIZE, f'<sz val="{size}" />')
        out.append(Live(f"cell font size {size}", build(control, Case("", (sized,))),
                        CELL_SIZE + CELL_SIZE_XLM))  # fmt: skip
    return out


def _attributes_cases() -> list[Live]:
    control = attributes_probe.control()
    by_name = {case.name: case for case in attributes_probe.cases()}
    out = []
    for name in dict.fromkeys(RUN_CASES + DXF_CASES + ROW_COL_CASES):
        reads = (
            (RUN_SIZE if name in RUN_CASES else ())
            + (DXF_SIZE if name in DXF_CASES else ())
            + (ROW_COL if name in ROW_COL_CASES else ())
        )
        out.append(Live(f"attributes {name}", build(control, by_name[name]), reads))
    zero_row = (SHEET, attributes_probe.ROW1, '<row r="1" ht="0" customHeight="1">')
    zero_col = (
        SHEET,
        attributes_probe.COL,
        attributes_probe.COL.replace('width="20"', 'width="0"'),
    )
    for name, edit in (("row ht='0'", zero_row), ("col width='0'", zero_col)):
        out.append(Live(f"attributes {name}", build(control, Case(name, (edit,))), ROW_COL))
    return out


def _outline_cases() -> list[Live]:
    control = control_workbook(outline_probe.fill)
    by_name = {case.name: case for case in outline_probe.cases()}
    out = []
    for value in OUTLINE_VALUES:
        label = f"{value!r}" + (" (control)" if value == "7" else "")
        for kind in ("row", "col"):
            case = by_name[f"{kind} {label}"]
            out.append(Live(f"outline {kind} {value}", build(control, case), OUTLINE))
    return out


def _value_cases() -> list[Live]:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 111
    ws["B1"] = "=A1*2"
    ws["C1"] = "=A1>0"
    ws["D1"] = "=A1*1E+300"
    buf = io.BytesIO()
    wb.save(buf)
    base = buf.getvalue()
    return [
        Live(f"value {v}", value_probe.rewrite(base, value_probe.with_value(v)), VALUE)
        for v in SUBNORMALS
    ]


def cases() -> list[Live]:
    return _edges_cases() + _attributes_cases() + _outline_cases() + _value_cases()


def _open(src: Path, reads: tuple[tuple[str, str], ...]) -> tuple[list[str], str]:
    """Each read's answer, or [] and the error Excel gave."""
    args = ["osascript", str(HERE / "live_reads.applescript"), str(src)]
    args += [expr for _, expr in reads]
    try:  # the time limit kills osascript, which waiting on Excel ignores a terminate signal
        result = subprocess.run(
            args, capture_output=True, text=True, check=False, timeout=int(TIME_LIMIT)
        )
    except subprocess.TimeoutExpired:
        subprocess.run(
            ["osascript", str(HERE / "restore_alerts.applescript")], check=False, timeout=20
        )
        return [], "time limit"
    if result.returncode != 0:
        found = re.search(r"\((-?\d+)\)", result.stderr)
        said = re.search(r"error: (.*)", result.stderr)  # Excel's words, not the script's path
        message = " ".join((said.group(1) if said else "").split())[:160]
        return [], f"{found.group(1)} {message}" if found else message
    answers = result.stdout.rstrip("\n").split("\t")
    if len(answers) != len(reads):
        raise SystemExit(f"unexpected output from live_reads.applescript: {result.stdout!r}")
    return [" ".join(a.split())[:200] for a in answers], ""


def main() -> Path:
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlio_live_reads.xlsx"
    lines: list[str] = []
    try:
        for live in cases():
            src.write_bytes(live.data)
            answers, error = _open(src, live.reads)
            if error:
                lines.append("\t".join([live.name, "refused", error, "", ""]))
            else:
                for (label, _), answer in zip(live.reads, answers, strict=True):
                    lines.append("\t".join([live.name, "opened", "", label, answer]))
            print("\n".join(lines[-max(1, len(live.reads)) :]), flush=True)
    finally:
        src.unlink(missing_ok=True)
    head = [
        f"# suite: {SUITE} (capture/live_reads.py); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: cases of the edges, attributes, outline "
        "and value_forms probes opened with alerts off (live_reads.applescript) and read "
        "through AppleScript while open, each answer as class:value; nothing saved",
        "\t".join(COLUMNS),
    ]
    out = CAPTURES / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    return out


def current() -> str:
    """What xlio reads today for what each case reads live."""
    import xlio

    out = []
    for live in cases():
        try:
            with xlio.open(io.BytesIO(live.data)) as wb:
                sheet = wb.sheets[0]
                cells = {f"{xlio.CellRef(c.row, c.column)}": c for c in sheet.cells(rich=True)}
                fonts, xfs = wb.styles.fonts, wb.styles.cell_formats
                got = []
                labels = {label for label, _ in live.reads}
                if "A1 font size" in labels:
                    got.append(f"A1 font size={fonts[xfs[cells['A1'].style].font_id].size}")
                if "A2 run 1 font size" in labels:
                    runs = cells["A2"].rich
                    got.append(f"A2 runs={[r.font.size if r.font else None for r in runs.runs]}"
                               if runs else "A2 runs=None")  # fmt: skip
                if "A3 condition 1 font size" in labels:
                    dxfs = wb.styles.differential_formats
                    got.append(f"dxf 0 font size={dxfs[0].font.size if dxfs[0].font else None}")
                cols = {c.min: c for c in sheet.columns}
                if "row 1 height" in labels:
                    r1, b = sheet.rows.get(1), cols.get(2)
                    got.append(
                        f"row 1 height={r1.height if r1 else None} custom="
                        f"{r1.custom_height if r1 else None} hidden={r1.hidden if r1 else None}"
                    )
                    got.append(
                        f"col B width={b.width if b else None} hidden={b.hidden if b else None}"
                    )
                if "row 2 outline level" in labels:
                    r2, b = sheet.rows.get(2), cols.get(2)
                    got.append(f"row 2 level={r2.outline_level if r2 else 0}")
                    got.append(f"col B level={b.outline_level if b else 0}")
                if "A1 value" in labels:
                    got.append(f"A1={cells['A1'].value!r}")
            out.append(f"{live.name}\t" + "; ".join(got))
        except Exception as exc:  # a refusal is a reading too
            out.append(f"{live.name}\t{type(exc).__name__}: {exc}"[:300])
    return "\n".join(out)


if __name__ == "__main__":
    print(current() if sys.argv[1:] == ["current"] else main())
