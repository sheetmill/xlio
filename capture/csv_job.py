"""How Excel opens CSV files, as a capture job (Windows) and as its Mac twin.

    uv run python -m capture.csv_job write JOB_ID OUT_DIR     # the job folder, for Windows
    uv run python -m capture.csv_job read JOB_DIR RESULT_DIR  # tests/captures/windows/csv_open.tsv
    uv run python -m capture.csv_job mac                      # tests/captures/csv_open.tsv

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/3adb865/docs/capture-jobs.md

The cases are ``.csv`` files the job opens, each through Excel's object model (``Workbooks.Open``,
alerts off), saving an ``.xlsx`` copy, read here as the Mac captures are read:
every byte from 0x80 to 0xFF in a file without a byte order mark; the same words with a UTF-8, a
UTF-16 LE and a UTF-16 BE byte order mark and without one; currency signs in Windows-1252 and in
UTF-8 with its mark; file names with ``[`` or ``]`` and the other names of ``text_extra``;
``text_extra``'s two-digit years and ``text_forms``' fields; ``text_limits``' five files, read as
a summary. ``mac`` opens the same files in Excel for Mac and writes the same columns, so the tests
pair the rows. Each row gives the copy's first font, the font of the workbook Excel made.
"""

from __future__ import annotations

import datetime
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.paths import write_capture
from capture.probe import CAPTURES, CONTAINER, HERE, TIME_LIMIT, excel_version
from capture.text_extra_probe import FORMS, LIMITS, NAMES, YEARS, summary
from capture.text_probe import WORDS, cells, csv_text, font, sheet_name

SUITE = "csv_open"
WINDOWS = CAPTURES / "windows"
REQUIRES: dict[str, object] = {
    "excel_build_min": {"windows": "16.0.20527.0"},
    "edition": "Microsoft 365",
}
COLUMNS = ["case", "outcome", "error", "sheet", "cell", "value", "format", "font"]
CONTROL = "control.csv"
_CONTROL_CHAR = re.compile("[\x00-\x08\x0b-\x1f\x7f-\x9f]")
_WINDOWS_NAMES = ("31 characters", "32 characters", "40 characters", "bracket", "History",
                  "apostrophe first")  # fmt: skip
"""``text_extra``'s names a Windows file can have (``* ? \\ :`` cannot be in one)."""


def cases() -> dict[str, tuple[str, bytes]]:
    """Case -> (file name, bytes), in the job's order."""
    out: dict[str, tuple[str, bytes]] = {}
    high = b"".join(bytes([b]) + b"\n" for b in range(0x80, 0x100))
    out["bytes 80-ff"] = ("bytes_80_ff.csv", high)
    words = csv_text(WORDS)
    out["words no mark"] = ("words_none.csv", words.encode("utf-8"))
    out["words utf-8 mark"] = ("words_utf8.csv", b"\xef\xbb\xbf" + words.encode("utf-8"))
    out["words utf-16-le mark"] = ("words_utf16le.csv", b"\xff\xfe" + words.encode("utf-16-le"))
    out["words utf-16-be mark"] = ("words_utf16be.csv", b"\xfe\xff" + words.encode("utf-16-be"))
    money = csv_text(["€5", "£5", "5€", "¥5", "$5"])
    out["currency windows-1252"] = ("currency_1252.csv", money.encode("cp1252"))
    out["currency utf-8 mark"] = ("currency_utf8.csv", b"\xef\xbb\xbf" + money.encode("utf-8"))
    for case in _WINDOWS_NAMES:
        out[f"name {case}"] = (f"{NAMES[case]}.csv", b"1\n")
    out["name bracket first"] = ("[ab].csv", b"1\n")
    out["name closing bracket"] = ("a]b.csv", b"1\n")
    out["years"] = ("xlio_years.csv", csv_text(YEARS).encode())
    out["forms"] = ("xlio_forms.csv", csv_text(FORMS).encode())
    for k, (case, data) in enumerate(LIMITS.items()):
        out[f"limits {case}"] = (f"xlio_limits_{k}.csv", data)
    return out


def _saved(k: int) -> str:
    return f"case{k:02d}.saved.xlsx"


def manifest(job_id: str) -> dict[str, object]:
    workbooks: list[dict[str, object]] = []
    for k, (name, _) in enumerate(cases().values()):
        steps: list[dict[str, object]] = [{"op": "open"}]
        if k == 0:
            steps.append({"op": "app", "get": ["standard_font", "standard_font_size"]})
        steps.append({"op": "save_as", "name": _saved(k), "format": "xlsx"})
        workbooks.append({"file": name, "steps": steps})
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "xlio",
        "purpose": "csv_open: open each .csv file through the object model, save an .xlsx copy",
        "requires": REQUIRES,
        "limits": {"job_seconds": 3600, "step_seconds": 600},
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
    for name, data in cases().values():
        (folder / "files" / name).write_bytes(data)
    return folder


def _shown(value: str) -> str:
    """A value with every control character written as ``\\uXXXX`` (captures hold none)."""
    return _CONTROL_CHAR.sub(lambda m: f"\\u{ord(m.group()):04x}", value)


def rows(case: str, copy: Path | None, error: str = "") -> list[str]:
    """The capture's rows for one case: refused, a summary (limits), or one row per cell."""
    if copy is None:
        return ["\t".join([case, "refused", error, "", "", "", "", ""])]
    with zipfile.ZipFile(copy) as z:
        title, first = sheet_name(z), font(z)
        if case.startswith("limits "):
            return ["\t".join([case, "opened", "", title, "summary", " ".join(summary(copy)), "",
                               first])]  # fmt: skip
        out = []
        for line in cells(z):
            ref, value, fmt = line.split("\t")
            out.append("\t".join([case, "opened", "", title, ref, _shown(value), fmt, first]))
        return out


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(job: Path, result: Path) -> Path:
    """``tests/captures/windows/csv_open.tsv`` and its manifest, from a job and its result."""
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
    for k, (case, steps) in enumerate(zip(cases(), outcome["workbooks"], strict=True)):
        opened = steps["steps"][0]
        if opened["status"] != "ok":
            lines += rows(case, None, str(opened.get("result") or opened["status"])[:200])
        else:
            lines += rows(case, result / "artifacts" / _saved(k))
    text = json.dumps(header, indent=1, sort_keys=True, ensure_ascii=False)
    head = [
        f"# suite: {SUITE}",
        "# generator: capture/csv_job.py",
        f"# method: the job beside this file ({SUITE}.job.json) opens each .csv file through the",
        "#   object model and saves an .xlsx copy; each copy is read as capture/text_probe.py",
        "#   reads the Mac's (control characters written \\uXXXX)",
        "# header (the job's result, copied whole):",
        *("#   " + line for line in text.splitlines()),
        "\t".join(COLUMNS),
    ]
    WINDOWS.mkdir(exist_ok=True)
    out = WINDOWS / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    (WINDOWS / f"{SUITE}.job.json").write_bytes((job / "job.json").read_bytes())
    return out


def _open_on_the_mac(src: Path, copy: Path) -> str:
    """Excel for Mac opens ``src`` and saves ``copy``; '' or the error Excel gave."""
    script = HERE / "csv_open.applescript"
    try:  # the time limit kills osascript, which waiting on Excel ignores a terminate signal
        result = subprocess.run(
            ["osascript", str(script), str(src), str(copy)],
            capture_output=True, text=True, check=False, timeout=int(TIME_LIMIT),
        )  # fmt: skip
    except subprocess.TimeoutExpired:
        subprocess.run(
            ["osascript", str(HERE / "restore_alerts.applescript")], check=False, timeout=20
        )
        return "time limit"
    if result.returncode == 0:
        return ""
    found = re.search(r"\((-?\d+)\)", result.stderr)
    said = re.search(r"error: (.*)", result.stderr)  # Excel's words, not the script's path
    message = " ".join((said.group(1) if said else "").split())[:160]
    return f"{found.group(1)} {message}" if found else message


def mac() -> Path:
    """The Mac twin: the same files opened in Excel for Mac, ``tests/captures/csv_open.tsv``."""
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    lines: list[str] = []
    for case, (name, data) in cases().items():
        src, copy = CONTAINER / name, CONTAINER / "xlio_csv_open_copy.xlsx"
        try:
            src.write_bytes(data)
            copy.unlink(missing_ok=True)
            error = _open_on_the_mac(src, copy)
            if error and ("[" in name or "]" in name):
                # Excel for Mac renames such a workbook (a[b]c.csv opens as a(b)c.csv), which
                # AppleScript cannot address: text_extra records the reading, made by hand.
                lines.append("\t".join([case, "not read", "renamed by Excel", "", "", "", "", ""]))
            else:
                lines += rows(case, None, error) if error else rows(case, copy)
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print(lines[-1][:120], flush=True)
    head = [
        f"# suite: {SUITE} (capture/csv_job.py mac); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: the job's files opened with alerts off "
        "(csv_open.applescript) and saved as an .xlsx copy, read as capture/text_probe.py reads "
        "it (control characters written \\uXXXX)",
        "\t".join(COLUMNS),
    ]
    out = CAPTURES / f"{SUITE}.tsv"
    write_capture(out, head + lines)
    return out


def main(argv: list[str]) -> None:
    if argv[:1] == ["write"] and len(argv) == 3:
        print(write(argv[1], Path(argv[2])))
    elif argv[:1] == ["read"] and len(argv) == 3:
        print(read(Path(argv[1]), Path(argv[2])))
    elif argv == ["mac"]:
        print(mac())
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
