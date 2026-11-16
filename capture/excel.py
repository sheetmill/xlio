"""Checks every capture makes before it asks Excel for anything (CAPTURING.md)."""

from __future__ import annotations

import subprocess


def refuse_open_workbooks() -> None:
    """Stop when Excel has any workbook open: a capture never runs beside someone's work."""
    result = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get count of workbooks'],
        capture_output=True,
        text=True,
        check=True,
        timeout=30,
    )
    if int(result.stdout.strip() or "0"):
        raise SystemExit("Excel has a workbook open: close it first, then run the capture")
