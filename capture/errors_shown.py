"""What Excel shows in a cell whose error rich value it does not keep (``errorType`` 14).

    uv run python -m capture.errors_shown     # tests/captures/errors_shown.tsv

The ``errors`` suite (``capture/errors_probe.py``) shows that Excel saves a cell written as
``#VALUE!`` with a rich value of ``errorType`` 14 as the number 0, while at load-time
calculation ``ERROR.TYPE`` of it gives 15 and ``VALUETOTEXT`` gives ``#CALC!``. A saved copy
cannot say what the cell itself shows. This suite reads it while the workbook is open: the
string value of A1:E1 (the text Excel draws) and the class of each value, through AppleScript.

Cases, on the ``errors`` suite's control (A1 = 1, B1 = ``=A1``, C1 = ``=ERROR.TYPE(A1)``, D1 =
``=ISERROR(A1)``, E1 = ``=VALUETOTEXT(A1)``, ``fullCalcOnLoad``): the control; ``rich 8``
(``#SPILL!``) and ``rich 13`` (``#CALC!``) as controls; ``rich 14``. Each is opened, read and
saved as a copy (its A1 as the ``errors`` suite reads it); the copy is opened again and read:
what a file Excel saved shows.

Runs where the other probes run, and in a macOS VM (no ``gtimeout``: the time limit is
``subprocess``'s own, and alerts are restored when it kills osascript).
"""

from __future__ import annotations

import datetime
import re
import subprocess
import zipfile
from pathlib import Path

from capture.edition import edition_line
from capture.errors_probe import cases as error_cases
from capture.errors_probe import cell, fill
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import (
    CAPTURES,
    CONTAINER,
    HERE,
    TIME_LIMIT,
    build,
    control_workbook,
    excel_version,
)

SUITE = "errors_shown"
CASES = ("control", "rich 8", "rich 13", "rich 14")
REFS = ("A1", "B1", "C1", "D1", "E1")
COLUMNS = [
    "case", "outcome", "error",
    *(f"{r} shown" for r in REFS), *(f"{r} class" for r in REFS),
    "copy A1",
    *(f"reopened {r} shown" for r in REFS),
]  # fmt: skip


def _shown(src: Path, copy: Path | None) -> tuple[list[str], str]:
    """A1:E1's string values, then their value classes; or [] and the error Excel gave."""
    args = ["osascript", str(HERE / "errors_shown.applescript"), str(src)]
    if copy is not None:
        args.append(str(copy))
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
    fields = result.stdout.rstrip("\n").split("\t")
    if len(fields) != 2 * len(REFS):  # a string value and a class for each cell
        raise SystemExit(f"unexpected output from errors_shown.applescript: {result.stdout!r}")
    return fields[0::2] + fields[1::2], ""


def main() -> Path:
    refuse_open_workbooks()
    version = excel_version()
    control = control_workbook(fill)
    chosen = [c for c in error_cases() if c.name in CASES]
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlio_errors_shown_src.xlsx"
    copy = CONTAINER / "xlio_errors_shown_copy.xlsx"
    lines: list[str] = []
    try:
        for case in chosen:
            src.write_bytes(build(control, case))
            copy.unlink(missing_ok=True)
            shown, error = _shown(src, copy)
            if error:
                lines.append("\t".join([case.name, "refused", error] + [""] * (len(COLUMNS) - 3)))
            else:
                with zipfile.ZipFile(copy) as z:
                    saved = cell(z, "A1")
                again, error = _shown(copy, None)
                again = again[: len(REFS)] if not error else [f"error {error}"] * len(REFS)
                lines.append("\t".join([case.name, "opened", "", *shown, saved, *again]))
            print(lines[-1], flush=True)
    finally:
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
    head = [
        f"# suite: {SUITE} (capture/errors_shown.py); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: capture/errors_probe.py's cases opened with "
        "alerts off (errors_shown.applescript): A1:E1's string values and value classes read "
        "while open, saved as a copy (its A1 as type:value), the copy opened again and read",
        "\t".join(COLUMNS),
    ]
    out = CAPTURES / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    return out


if __name__ == "__main__":
    print(main())
