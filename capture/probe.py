"""What the case-table probes share: a control workbook, part rewrites, Excel's copy, the TSV.

A probe is a table of :class:`Case` records. Each case is the control workbook with a few parts
rewritten (an exact text replaced, which must occur once) or added. Excel opens each with alerts
off and saves a copy (``save_copy.applescript``); a file it would offer to repair fails to open
with error -50 and is recorded as ``refused``. From the copy, each of the case's reads (a part
and a regular expression) gives what Excel wrote, or a probe's own reader does.

Run a probe only while Excel is running and answers AppleScript, no other capture is running, and no
test or run is feeding crafted input to the reader (CAPTURING.md).
Workbooks live in Excel's container and are deleted afterwards.
"""

from __future__ import annotations

import datetime
import io
import re
import subprocess
import zipfile
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlio"
CAPTURES = ROOT / "tests" / "captures"
SHEET = "xl/worksheets/sheet1.xml"
TIME_LIMIT = "60"
WIDTH = 400  # the longest text a read keeps in the TSV


@dataclass(frozen=True)
class Case:
    """One workbook: the control with ``edits`` made and ``parts`` added or replaced.

    Attributes:
        name: The case's name in the TSV.
        edits: (part, old, new): ``old`` must occur exactly once in the part.
        parts: Parts added or replaced whole (content types are the probe's to edit).
        reads: (label, part, pattern): every match of ``pattern`` in Excel's copy of the part,
            the first group if it has one, joined with ``|``; ``absent`` if the part is not
            there.
    """

    name: str
    edits: tuple[tuple[str, str, str], ...] = ()
    parts: dict[str, str] = field(default_factory=dict)
    reads: tuple[tuple[str, str, str], ...] = ()


def build(control: bytes, case: Case) -> bytes:
    """The control package with the case's edits made and parts added."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(control)) as src, zipfile.ZipFile(out, "w") as dst:
        names = set(src.namelist())
        edited = {p for p, _, _ in case.edits} | set(case.parts)
        for info in src.infolist():
            if info.filename not in edited:  # binary parts (media) are copied as they are
                dst.writestr(info, src.read(info), compress_type=zipfile.ZIP_DEFLATED)
                continue
            body = case.parts.get(info.filename) or src.read(info).decode()
            for part, old, new in case.edits:
                if part == info.filename:
                    if body.count(old) != 1:
                        raise ValueError(f"{case.name}: {old!r} occurs {body.count(old)}x")
                    body = body.replace(old, new)
            dst.writestr(info, body.encode(), compress_type=zipfile.ZIP_DEFLATED)
        for name, body in case.parts.items():
            if name not in names:
                dst.writestr(name, body.encode(), compress_type=zipfile.ZIP_DEFLATED)
        missing = {p for p, _, _ in case.edits} - names - set(case.parts)
        if missing:
            raise ValueError(f"{case.name}: no part {sorted(missing)}")
    return out.getvalue()


def save_copy(src: Path, copy: Path) -> str:
    """Excel opens ``src`` and saves ``copy``; '' or the error number Excel gave."""
    result = subprocess.run(
        ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(HERE / "save_copy.applescript"),
         str(src), str(copy)],
        capture_output=True, text=True, check=False,
    )  # fmt: skip
    if result.returncode == 0:
        return ""
    found = re.search(r"\((-?\d+)\)", result.stderr)
    return found.group(1) if found else result.stderr.strip()[:80]


def excel_version() -> str:
    return subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True, text=True, check=True, timeout=30,
    ).stdout.strip()  # fmt: skip


def read(copy: zipfile.ZipFile, part: str, pattern: str) -> str:
    """Every match of ``pattern`` in the part (its first group, if any), joined with ``|``."""
    if part not in copy.namelist():
        return "absent"
    text = copy.read(part).decode()
    found = [m.group(1) if m.groups() else m.group(0) for m in re.finditer(pattern, text, re.S)]
    joined = "|".join(" ".join(f.split()) for f in found)
    return joined[:WIDTH].replace("\t", " ") or "-"


def run(
    suite: str,
    control: bytes,
    cases: Iterable[Case],
    labels: list[str],
    method: str,
    notes: str,
    reader: Callable[[zipfile.ZipFile, Case], list[str]] | None = None,
) -> None:
    """Run every case through Excel and write ``tests/captures/<suite>.tsv``.

    ``labels`` name the read columns; ``reader``, if given, makes them instead of the reads.
    """
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src, copy = CONTAINER / f"xlio_{suite}_src.xlsx", CONTAINER / f"xlio_{suite}_copy.xlsx"
    rows = []
    try:
        for case in cases:
            src.write_bytes(build(control, case))
            copy.unlink(missing_ok=True)
            error = save_copy(src, copy)
            if error:
                rows.append("\t".join([case.name, "refused", error] + [""] * len(labels)))
            else:
                with zipfile.ZipFile(copy) as z:
                    if reader is not None:
                        values = reader(z, case)
                    else:
                        values = [read(z, part, pattern) for _, part, pattern in case.reads]
                values += [""] * (len(labels) - len(values))
                rows.append("\t".join([case.name, "opened", "", *values]))
            print(rows[-1], flush=True)
    finally:
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
    header = [
        f"# suite: {suite} (capture); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: {method}",
        f"# {notes}; refused = error -50 (Excel would repair the file)",
        "\t".join(["case", "outcome", "error", *labels]),
    ]
    out = CAPTURES / f"{suite}.tsv"
    out.write_text("\n".join(header + rows) + "\n", encoding="utf-8")
    print(f"wrote {out.relative_to(ROOT)}: {len(rows)} cases")


def control_workbook(fill: Callable[[object], None], sheets: int = 1) -> bytes:
    """openpyxl's workbook, filled by ``fill`` (given the workbook), as bytes."""
    import openpyxl

    wb = openpyxl.Workbook()
    for i in range(1, sheets):
        wb.create_sheet(f"Sheet{i + 1}")
    fill(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def part_text(data: bytes, part: str) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return z.read(part).decode()
