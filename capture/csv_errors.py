"""What Excel shows for an error code written as a CSV field, read while the workbook is open.

    uv run python -m capture.csv_errors write JOB_ID OUT_DIR     # the job folder, for Windows
    uv run python -m capture.csv_errors read JOB_DIR RESULT_DIR  # the Windows capture
    uv run python -m capture.csv_errors mac                      # the Mac twin
    uv run python -m capture.csv_errors current                  # what xlio reads for each row

The ``text_forms`` and ``csv_open`` captures read the newer errors (``#SPILL!``, ``#CALC!`` ...)
and ``#GETTING_DATA`` from the copy Excel saves, whose ``<v>`` holds ``#VALUE!`` or ``#N/A``; a
copy cannot say what the open cell showed. Here each field is read while the workbook is open,
column A widened to 30 characters first so that no text is cut to ``#``:
the text the cell shows, its value (on Windows ``Value2``, typed; on the Mac the class of the
value AppleScript gives), and ``ERROR.TYPE`` and ``ISERROR`` of it, written beside the cells
after the copy is saved. The saved copy's cell, with the error its rich value stands for, is
kept for the record.

Two files hold the same fields, one per line: quoted (as ``text_forms`` writes them) and bare.

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/3adb865/docs/capture-jobs.md
"""

from __future__ import annotations

import datetime
import hashlib
import io
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

from capture.csv_job import REQUIRES
from capture.edition import edition_line
from capture.errors_probe import rich_error
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import CAPTURES, CONTAINER, HERE, TIME_LIMIT, excel_version
from capture.text_probe import csv_text

SUITE = "csv_errors"
WINDOWS = CAPTURES / "windows"
FIELDS = [
    "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!",
    "#PYTHON!", "#TIMEOUT!", "#GETTING_DATA", "#EXTERNAL!", "#calc!", "#getting_data",
    "#N/A", "#DIV/0!",
]  # fmt: skip
FILES = {
    "quoted": ("xlio_errors_quoted.csv", csv_text(FIELDS).encode()),
    "bare": ("xlio_errors_bare.csv", "".join(f + "\n" for f in FIELDS).encode()),
}
COLUMNS = ["file", "field", "outcome", "error", "shown", "value", "ERROR.TYPE", "ISERROR",
           "copy", "copy rich value"]  # fmt: skip
CONTROL = "control.csv"
N = len(FIELDS)


def _saved(k: int) -> str:
    return f"case{k:02d}.saved.xlsx"


def manifest(job_id: str) -> dict[str, object]:
    formulas = [[f"=ERROR.TYPE(A{r})", f"=ISERROR(A{r})"] for r in range(1, N + 1)]
    workbooks: list[dict[str, object]] = []
    for k, (name, _) in enumerate(FILES.values()):
        steps: list[dict[str, object]] = [
            {"op": "open"},
            {"op": "set_size", "columns": "A:A", "width": 30},
            {"op": "read", "range": f"A1:A{N}", "props": ["text", "value2"]},
            {"op": "save_as", "name": _saved(k), "format": "xlsx"},
            {"op": "write", "range": f"B1:C{N}", "formulas": formulas},
            {"op": "calculate", "kind": "full"},
            {"op": "read", "range": f"B1:C{N}", "props": ["text"]},
        ]
        workbooks.append({"file": name, "steps": steps})
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "xlio",
        "purpose": "csv_errors: open .csv files of error codes, read the cells while open, "
        "save an .xlsx copy, read ERROR.TYPE and ISERROR of each",
        "requires": REQUIRES,
        "limits": {"job_seconds": 1800, "step_seconds": 300},
        "on_refused_open": "record",
        "control": CONTROL,
        "audit": False,
        "workbooks": workbooks,
    }


def write(job_id: str, out: Path) -> Path:
    folder = out / job_id
    (folder / "files").mkdir(parents=True)
    (folder / "job.json").write_text(json.dumps(manifest(job_id), indent=2) + "\n")
    (folder / "files" / CONTROL).write_bytes(b"1\n")
    for name, data in FILES.values():
        (folder / "files" / name).write_bytes(data)
    return folder


def copy_cells(copy: Path) -> list[tuple[str, str]]:
    """Each row's A cell in the saved copy, as type:value, and the error its rich value holds."""
    out = []
    with zipfile.ZipFile(copy) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
        for r in range(1, N + 1):
            found = re.search(rf'<c r="A{r}"([^>]*?)(?:/>|>(.*?)</c>)', sheet, re.S)
            if found is None:
                out.append(("absent", "-"))
                continue
            attrs, body = found.group(1), found.group(2) or ""
            t = re.search(r'\bt="([^"]*)"', attrs)
            vm = re.search(r'\bvm="(\d+)"', attrs)
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            out.append(
                (f"{t.group(1) if t else 'n'}:{v.group(1) if v else ''}",
                 rich_error(z, vm.group(1) if vm else None))
            )  # fmt: skip
    return out


def _typed(value: dict[str, object]) -> str:
    """A ``Value2`` read by the job, as ``type:value``."""
    kind = value.get("t")
    if kind == "error":
        return f"error:{value.get('v')}"
    if kind == "number":
        return f"number:{value.get('repr')}"
    return f"{kind}:{value.get('v', '')}"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(job: Path, result: Path) -> Path:
    """``tests/captures/windows/csv_errors.tsv`` and its manifest, from a job and its result."""
    outcome = json.loads((result / "result.json").read_text(encoding="utf-8"))
    if outcome.get("status") != "ok":
        sys.exit(f"the job did not end ok: {outcome.get('status')!r}")
    header = outcome["header"]
    if header["manifest_sha256"] != _sha256(job / "job.json"):
        sys.exit("the result's manifest_sha256 is not this job's manifest")
    for name, entry in header["inputs"].items():
        if entry["sha256"] != _sha256(job / "files" / name):
            sys.exit(f"the result's input {name} is not this job's")
    lines: list[str] = []
    for k, (case, wb) in enumerate(zip(FILES, outcome["workbooks"], strict=True)):
        steps = wb["steps"]
        if steps[0]["status"] != "ok":
            error = str(steps[0].get("result") or steps[0]["status"])[:200]
            lines += ["\t".join([case, f, "refused", error, *[""] * 6]) for f in FIELDS]
            continue
        cells, after = steps[2]["result"], steps[6]["result"]
        saved = copy_cells(result / "artifacts" / _saved(k))
        for r, field in enumerate(FIELDS):
            lines.append("\t".join([
                case, field, "opened", "", str(cells["text"][r][0]),
                _typed(cells["value2"][r][0]), str(after["text"][r][0]),
                str(after["text"][r][1]), *saved[r],
            ]))  # fmt: skip
    text = json.dumps(header, indent=1, sort_keys=True, ensure_ascii=False)
    head = [
        f"# suite: {SUITE}",
        "# generator: capture/csv_errors.py",
        f"# method: the job beside this file ({SUITE}.job.json) opens each .csv file through the",
        "#   object model, widens column A to 30, reads A1:A15 (Text, Value2) while it is open,",
        "#   saves an .xlsx copy,",
        "#   writes =ERROR.TYPE and =ISERROR beside each cell and reads their Text; the copy's",
        "#   cells are read with the error their rich value holds",
        "# header (the job's result, copied whole):",
        *("#   " + line for line in text.splitlines()),
        "\t".join(COLUMNS),
    ]
    WINDOWS.mkdir(exist_ok=True)
    out = WINDOWS / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    (WINDOWS / f"{SUITE}.job.json").write_bytes((job / "job.json").read_bytes())
    return out


def _open_on_the_mac(src: Path, copy: Path) -> tuple[list[list[str]], str]:
    """Excel for Mac's reads of each row, or [] and the error Excel gave."""
    script = HERE / "csv_errors.applescript"
    try:  # the time limit kills osascript, which waiting on Excel ignores a terminate signal
        result = subprocess.run(
            ["osascript", str(script), str(src), str(copy), str(N)],
            capture_output=True, text=True, check=False, timeout=int(TIME_LIMIT),
        )  # fmt: skip
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
    rows = [line.split("\t") for line in result.stdout.strip("\n").split("\n")]
    if len(rows) != N or any(len(row) != 4 for row in rows):
        raise SystemExit(f"unexpected output from csv_errors.applescript: {result.stdout!r}")
    return rows, ""


def mac() -> Path:
    """The Mac twin: the same files opened in Excel for Mac, ``tests/captures/csv_errors.tsv``."""
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    lines: list[str] = []
    for case, (name, data) in FILES.items():
        src, copy = CONTAINER / name, CONTAINER / "xlio_csv_errors_copy.xlsx"
        try:
            src.write_bytes(data)
            copy.unlink(missing_ok=True)
            rows, error = _open_on_the_mac(src, copy)
            if error:
                lines += ["\t".join([case, f, "refused", error, *[""] * 6]) for f in FIELDS]
            else:
                saved = copy_cells(copy)
                for field, row, cell in zip(FIELDS, rows, saved, strict=True):
                    lines.append("\t".join([case, field, "opened", "", *row, *cell]))
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print("\n".join(lines[-N:]), flush=True)
    head = [
        f"# suite: {SUITE} (capture/csv_errors.py mac); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: the job's files opened with alerts off "
        "(csv_errors.applescript); column A widened to 30, A1:A15 read while open (string "
        "value, the class of the value), saved as an .xlsx copy, =ERROR.TYPE and =ISERROR "
        "written beside each cell and read; the copy's cells read with the error their rich "
        "value holds",
        "\t".join(COLUMNS),
    ]
    out = CAPTURES / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    return out


def current() -> str:
    """What xlio reads for each row today, on both platforms (the same for these fields)."""
    import xlio

    out = []
    for case, (_, data) in FILES.items():
        for platform in ("mac", "windows"):
            with xlio.read_csv(io.BytesIO(data), platform=platform) as wb:
                cells = {c.row: c for c in wb.sheets[0].cells()}
            for r, field in enumerate(FIELDS, start=1):
                c = cells.get(r)
                got = f"{c.type}:{c.value}" if c is not None else "absent"
                out.append(f"{case}\t{platform}\t{field}\t{got}")
    return "\n".join(out)


def main(argv: list[str]) -> None:
    if argv[:1] == ["write"] and len(argv) == 3:
        print(write(argv[1], Path(argv[2])))
    elif argv[:1] == ["read"] and len(argv) == 3:
        print(read(Path(argv[1]), Path(argv[2])))
    elif argv == ["mac"]:
        print(mac())
    elif argv == ["current"]:
        print(current())
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
