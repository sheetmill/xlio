"""What Excel makes of CSV files at the edges: two-digit years, sheet names, and its limits.

    uv run python -m capture.text_extra_probe                 # text_extra, unattended
    uv run python -m capture.text_extra_probe --limits-open   # text_limits: opens, then exits
    uv run python -m capture.text_extra_probe --limits-save ANSWER [REPORT]

``text_extra`` (opened through AppleScript, as ``text`` is, and saved as a workbook copy):

- ``years``: dates with two-digit years (``1/2/00`` to ``1/2/99`` around 29/30 and 49/50, and
  ``2-Jan-29``, ``2-Jan-30``): where Excel's century changes;
- ``name <case>``: a one-cell file whose name gives the sheet its name: past 31 characters, with
  each character Excel refuses in a sheet name that a file name can hold (``[ ] * ? \\ :``),
  and ``History``, a name Excel reserves.

Each row is a cell of the copy's first sheet, as in ``text``, with the sheet's name.

``text_limits``: a field of 32,767 characters, one of 32,768 and one of 40,000; a file of
1,048,577 rows; and a row of 16,385 fields. Excel may tell a person that a file was not loaded
completely, so these files are opened through Launch Services (``--limits-open``) and the script
exits, sending Excel nothing while a dialog may be open. Answer each dialog (OK), then run
``--limits-save`` with ``none`` (no dialog) or ``ok`` (dialogs answered) and, after it, what
the dialogs said, which goes into the TSV's header. Each row gives the sheet's last row and
column, its number of cells, and the lengths of the texts in its first row.

Run it only while Excel is running and answers AppleScript, no other capture is running, and no
test or run is feeding crafted input to the reader (CAPTURING.md). Files live in Excel's
container and are deleted afterwards.
"""

from __future__ import annotations

import datetime
import subprocess
import sys
import zipfile
from pathlib import Path

import xlio
from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import CAPTURES, CONTAINER, HERE, TIME_LIMIT, excel_version, save_copy
from capture.text_probe import cells, csv_text, sheet_name

YEARS = ["1/2/00", "1/2/29", "1/2/30", "1/2/49", "1/2/50", "1/2/99", "2-Jan-29", "2-Jan-30",
         "1/2/2029", "1/2/1930"]  # fmt: skip
NAMES = {
    "31 characters": "n" * 31,
    "32 characters": "n" * 32,
    "40 characters": "abcdefghij" * 4,
    "bracket": "a[b]c",
    "star": "a*b",
    "question mark": "a?b",
    "backslash": "a\\b",
    "colon": "a:b",
    "History": "History",
    "apostrophe first": "'ab",
}
LIMITS = {
    "field 32767": ("x" * 32767 + "\n").encode(),
    "field 32768": ("x" * 32768 + "\n").encode(),
    "field 40000": ("x" * 40000 + "\n").encode(),
    "rows 1048577": b"".join(b"%d\n" % n for n in range(1, 1_048_578)),
    "columns 16385": (",".join(str(n) for n in range(1, 16386)) + "\n").encode(),
}


def extra_files() -> dict[str, tuple[str, bytes]]:
    """Case -> (file name, bytes)."""
    out = {"years": ("xlio_years.csv", csv_text(YEARS).encode())}
    for case, stem in NAMES.items():
        out[f"name {case}"] = (f"{stem}.csv", b"1\n")
    return out


def extra() -> None:
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    rows = []
    for case, (name, data) in extra_files().items():
        src, copy = CONTAINER / name, CONTAINER / "xlio_extra_copy.xlsx"
        try:
            src.write_bytes(data)
            copy.unlink(missing_ok=True)
            error = save_copy(src, copy)
            if error:
                rows.append(f"{case}\trefused\t{error}\t\t")
            else:
                with zipfile.ZipFile(copy) as z:
                    title = sheet_name(z)
                    rows += [f"{case}\topened\t{title}\t{line}" for line in cells(z)]
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print(rows[-1], flush=True)
    header = [
        f"# suite: text_extra (capture.text_extra_probe); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: a .csv file opened with alerts off "
        "(File > Open's defaults) and saved as a workbook copy",
        "# one line per cell of the copy's first sheet: type:value as the XML has it, and the "
        "cell's number format; the sheet column is the copy's sheet name",
        "case\toutcome\tsheet\tcell\tvalue\tformat",
    ]
    write_capture(CAPTURES / "text_extra.tsv", header + rows)


def limit_source(i: int) -> Path:
    return CONTAINER / f"xlio_limit_{i}.csv"


def limits_open() -> None:
    """Write the limit cases and open each as a person does; no more."""
    refuse_open_workbooks()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    for i, (case, data) in enumerate(LIMITS.items()):
        src = limit_source(i)
        src.write_bytes(data)
        subprocess.run(["open", "-a", "Microsoft Excel", str(src)], check=True, timeout=60)
        print(f"{case}: opened {src.name}; answer any dialog in Excel", flush=True)


def summary(copy: Path) -> list[str]:
    """The first sheet's last row and column, its cells, and its first row's text lengths."""
    last_row = last_col = count = 0
    lengths = []
    with xlio.open(copy) as wb:
        for cell in wb.sheets[0].cells():
            count += 1
            last_row, last_col = max(last_row, cell.row), max(last_col, cell.column)
            if cell.row == 1 and isinstance(cell.value, str) and len(lengths) < 3:
                lengths.append(str(len(cell.value)))
    return [str(last_row), str(last_col), str(count), " ".join(lengths) or "-"]


def limits_save(answer: str, report: str) -> None:
    version = excel_version()
    rows = []
    for i, case in enumerate(LIMITS):
        src, copy = limit_source(i), CONTAINER / f"xlio_limit_{i}_copy.xlsx"
        try:
            copy.unlink(missing_ok=True)
            result = subprocess.run(
                ["gtimeout", "-k", "5", TIME_LIMIT, "osascript",
                 str(HERE / "save_open.applescript"), src.name, str(copy)],
                capture_output=True, text=True, check=False,
            )  # fmt: skip
            if result.returncode != 0:
                error = result.stderr.strip().rsplit("error: ", 1)[-1].replace("\t", " ")[:120]
                rows.append(f"{case}\tnot saved\t{error}\t\t\t")
            else:
                rows.append("\t".join([case, "opened", "", *summary(copy)]))
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print(rows[-1], flush=True)
    header = [
        f"# suite: text_limits (capture.text_extra_probe); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: a .csv file opened through Launch "
        f"Services (open -a) as a person opens it, dialogs answered '{answer}' by the person "
        "(none: not shown), then saved as a workbook copy through AppleScript",
        "# the copy's first sheet: last row, last column, number of cells, lengths of the "
        "texts in row 1" + (f"; {report}" if report else ""),
        "case\toutcome\terror\tlast row\tlast column\tcells\tlengths",
    ]
    write_capture(CAPTURES / "text_limits.tsv", header + rows)


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        extra()
    elif args == ["--limits-open"]:
        limits_open()
    elif args[:1] == ["--limits-save"] and args[1:2] in (["none"], ["ok"]):
        limits_save(args[1], " ".join(args[2:]))
    else:
        raise SystemExit("no argument, --limits-open, or --limits-save none|ok [REPORT]")
