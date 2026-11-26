"""What Excel shows when it opens its own saved copy of ``csv_errors``' error codes.

    uv run python -m capture.errors_saved           # tests/captures/errors_saved.tsv
    uv run python -m capture.errors_saved current   # what xlio reads for each saved form

Excel saves a newer error as ``#VALUE!`` with a rich value of its ``errorType`` (capture
``csv_errors``); ``#PYTHON!`` it saves as ``errorType`` 18 with ``subType`` 1 and three more keys,
where the ``errors`` capture's ``errorType`` 18 (``subType`` 0) shows ``#UNKNOWN!``. Here Excel
opens ``csv_errors``' quoted file and saves it as an .xlsx copy, then opens that copy and reads
each cell while it is open (``csv_errors.applescript`` both times): the text it shows, the class
of its value, ``ERROR.TYPE`` and ``ISERROR``. Each row gives the field, the saved copy's cell and
the error its rich value holds, and what the reopened copy shows.

Runs where the other probes run, and in a macOS VM.
"""

from __future__ import annotations

import datetime
import sys
from pathlib import Path

from capture.csv_errors import FIELDS, FILES, N, _open_on_the_mac, copy_cells
from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import CAPTURES, CONTAINER, excel_version

SUITE = "errors_saved"
COLUMNS = ["field", "outcome", "error", "copy", "copy rich value", "shown", "value",
           "ERROR.TYPE", "ISERROR"]  # fmt: skip


def main() -> Path:
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    name, data = FILES["quoted"]
    src = CONTAINER / name
    saved = CONTAINER / "xlio_errors_saved.xlsx"
    again = CONTAINER / "xlio_errors_saved_again.xlsx"
    lines: list[str] = []
    try:
        src.write_bytes(data)
        for path in (saved, again):
            path.unlink(missing_ok=True)
        _, error = _open_on_the_mac(src, saved)
        if error:
            raise SystemExit(f"Excel did not save the copy: {error}")
        cells = copy_cells(saved)
        rows, error = _open_on_the_mac(saved, again)
        for k, field in enumerate(FIELDS):
            if error:
                lines.append("\t".join([field, "refused", error, *cells[k], "", "", "", ""]))
            else:
                lines.append("\t".join([field, "opened", "", *cells[k], *rows[k]]))
    finally:
        for path in (src, saved, again):
            path.unlink(missing_ok=True)
    print("\n".join(lines), flush=True)
    head = [
        f"# suite: {SUITE} (capture/errors_saved.py); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: csv_errors' quoted file opened with alerts "
        "off and saved as an .xlsx copy (csv_errors.applescript); that copy opened, column A "
        f"widened to 30, A1:A{N} read while open (string value, the class of the value), "
        "=ERROR.TYPE and =ISERROR written beside each cell and read",
        "\t".join(COLUMNS),
    ]
    out = CAPTURES / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    return out


def current() -> str:
    """What xlio reads for each saved form csv_errors recorded (its copy's rich values)."""
    from xlio._book import _RICH_ERRORS

    out = []
    capture = (CAPTURES / "csv_errors.tsv").read_text(encoding="utf-8").splitlines()
    header = next(line.split("\t") for line in capture if line.startswith("file\t"))
    for line in capture:
        row = dict(zip(header, line.split("\t"), strict=False))
        if row.get("file") != "quoted":
            continue
        rich = row["copy rich value"]
        kind = next((p.split("=")[1] for p in rich.split() if p.startswith("errorType=")), None)
        got = _RICH_ERRORS.get(kind, "#VALUE!") if kind else row["copy"].partition(":")[2]
        out.append(f"{row['field']}\t{row['copy']} {rich}\t-> {got}")
    return "\n".join(out)


if __name__ == "__main__":
    print(current() if sys.argv[1:] == ["current"] else main())
