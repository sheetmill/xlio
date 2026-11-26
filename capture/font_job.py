"""Cell font sizes in fractions of a point, read through Excel for Windows' object model.

    uv run python -m capture.font_job write JOB_ID OUT_DIR     # the job folder
    uv run python -m capture.font_job read JOB_DIR RESULT_DIR  # the Windows capture
    uv run python -m capture.font_job current                  # what xlio reads for each case

The ``edges`` capture's saved copies keep a cell font size cut to twentieths of a point (409.54
as 409.5); Excel for Mac's AppleScript gives a font size only as a whole number (capture
``live_reads``). This job opens ``edges``' workbook with its font's size rewritten (the sizes
``edges`` uses, and sizes between twentieths near 10 and 12) and reads A1's ``Font.Size`` while
the workbook is open (``read_format``), a number the object model gives with its fraction; the
saved copy's size is kept for the record.

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/e3a5207/docs/capture-jobs.md
"""

from __future__ import annotations

import hashlib
import io
import json
import re
import sys
import zipfile
from pathlib import Path

from capture import edges_probe
from capture.csv_job import REQUIRES
from capture.paths import write_capture
from capture.probe import CAPTURES, Case, build, control_workbook

SUITE = "font_sizes"
WINDOWS = CAPTURES / "windows"
SIZES = (
    "10", "10.03", "10.05", "10.07", "11.11", "12", "12.3", "12.34", "12.35", "12.36", "12.4",
    "12.5", "12.6", "13", "409", "409.5", "409.51", "409.54", "409.55", "409.56", "409.6",
    "409.9", "410",
)  # fmt: skip
CONTROL = "control.xlsx"
COLUMNS = ["case", "outcome", "error", "A1 Font.Size", "copy font 0 size"]


def _name(k: int) -> str:
    return f"size{k:02d}.xlsx"


def _saved(k: int) -> str:
    return f"size{k:02d}.saved.xlsx"


def files() -> dict[str, tuple[str, bytes]]:
    """Case -> (file name, bytes), in the job's order."""
    control = control_workbook(edges_probe.fill)
    out = {}
    for k, size in enumerate(SIZES):
        edit = (edges_probe.STYLES, edges_probe.SIZE, f'<sz val="{size}" />')
        out[f"font size {size}"] = (_name(k), build(control, Case("", (edit,))))
    return out


def manifest(job_id: str) -> dict[str, object]:
    workbooks = [
        {
            "file": name,
            "steps": [
                {"op": "open"},
                {"op": "read_format", "range": "A1", "props": ["font.size"]},
                {"op": "save_as", "name": _saved(k), "format": "xlsx"},
            ],
        }
        for k, (name, _) in enumerate(files().values())
    ]
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "xlio",
        "purpose": "font_sizes: open workbooks whose font has a fractional size, read A1's "
        "Font.Size while open, save an .xlsx copy",
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
    (folder / "files" / CONTROL).write_bytes(control_workbook(edges_probe.fill))
    for name, data in files().values():
        (folder / "files" / name).write_bytes(data)
    return folder


def _copy_size(copy: Path) -> str:
    with zipfile.ZipFile(copy) as z:
        styles = z.read("xl/styles.xml").decode()
    font = re.search(r"<font>.*?</font>", styles, re.S)
    size = re.search(r'<sz val="([^"]*)"', font.group(0)) if font else None
    return size.group(1) if size else "-"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(job: Path, result: Path) -> Path:
    """``tests/captures/windows/font_sizes.tsv`` and its manifest, from a job and its result."""
    outcome = json.loads((result / "result.json").read_text(encoding="utf-8"))
    if outcome.get("status") != "ok":
        sys.exit(f"the job did not end ok: {outcome.get('status')!r}")
    header = outcome["header"]
    if header["manifest_sha256"] != _sha256(job / "job.json"):
        sys.exit("the result's manifest_sha256 is not this job's manifest")
    for name, entry in header["inputs"].items():
        if entry["sha256"] != _sha256(job / "files" / name):
            sys.exit(f"the result's input {name} is not this job's")
    lines = []
    for k, (case, wb) in enumerate(zip(files(), outcome["workbooks"], strict=True)):
        steps = wb["steps"]
        if steps[0]["status"] != "ok":
            error = str(steps[0].get("result") or steps[0]["status"])[:200]
            lines.append("\t".join([case, "refused", error, "", ""]))
            continue
        size = steps[1]["result"]["cells"][0][0]["font.size"]
        copy = _copy_size(result / "artifacts" / _saved(k))
        lines.append("\t".join([case, "opened", "", repr(size), copy]))
    text = json.dumps(header, indent=1, sort_keys=True, ensure_ascii=False)
    head = [
        f"# suite: {SUITE}",
        "# generator: capture/font_job.py",
        f"# method: the job beside this file ({SUITE}.job.json) opens each workbook through the",
        "#   object model, reads A1's Font.Size while it is open (read_format) and saves an",
        "#   .xlsx copy, whose first font's size is read",
        "# header (the job's result, copied whole):",
        *("#   " + line for line in text.splitlines()),
        "\t".join(COLUMNS),
    ]
    WINDOWS.mkdir(exist_ok=True)
    out = WINDOWS / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    (WINDOWS / f"{SUITE}.job.json").write_bytes((job / "job.json").read_bytes())
    return out


def current() -> str:
    """What xlio reads for A1's font size in each case."""
    import xlio

    out = []
    for case, (_, data) in files().items():
        try:
            with xlio.open(io.BytesIO(data)) as wb:
                cell = next(c for c in wb.sheets[0].cells() if (c.row, c.column) == (1, 1))
                styles = wb.styles
                size = styles.fonts[styles.cell_formats[cell.style].font_id].size
            out.append(f"{case}\t{size!r}")
        except xlio.PackageError as exc:
            out.append(f"{case}\trefused: {exc}"[:160])
    return "\n".join(out)


def main(argv: list[str]) -> None:
    if argv[:1] == ["write"] and len(argv) == 3:
        print(write(argv[1], Path(argv[2])))
    elif argv[:1] == ["read"] and len(argv) == 3:
        print(read(Path(argv[1]), Path(argv[2])))
    elif argv == ["current"]:
        print(current())
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
