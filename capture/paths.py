"""What must not be in a committed workbook or capture: a path into someone's home directory.

Shared by the fixture tooling, which refuses to write such a workbook, the probes, which refuse
to write such a capture (:func:`write_capture`), and the tests, which check every committed
workbook and capture. A cached value such as ``CELL("filename")`` holds the saved file's path,
and an error from AppleScript names the script's, so every part and row is searched, in each
form a path takes in XML, and with the temporary folders macOS gives.
"""

from __future__ import annotations

import re
from pathlib import Path

HOME_PATH = re.compile(
    rb"/Users/|/home/|/private/var/|/var/folders/|[A-Za-z]:\\\\Users\\\\|[A-Za-z]:\\Users\\|file:///"
)


def write_capture(path: Path, lines: list[str]) -> None:
    """Write a capture's lines as a TSV, refusing any line that holds a path."""
    for number, line in enumerate(lines, 1):
        if HOME_PATH.search(line.encode("utf-8")):
            raise SystemExit(f"{path.name}, line {number}, holds a path: not written; {line!r}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
