"""Committed captures hold no path into someone's home directory or a temporary folder."""

from __future__ import annotations

from pathlib import Path

import pytest
from capture.paths import HOME_PATH, write_capture

CAPTURES = sorted((Path(__file__).with_name("captures")).glob("*.tsv"))


def test_there_are_captures():
    assert len(CAPTURES) >= 15


@pytest.mark.parametrize("path", CAPTURES, ids=[p.name for p in CAPTURES])
def test_a_capture_holds_no_path(path):
    for number, line in enumerate(path.read_bytes().splitlines(), 1):
        assert HOME_PATH.search(line) is None, (path.name, number)


@pytest.mark.parametrize(
    "planted",
    [
        "/Users/x/Library/a.xlsx",
        "/private/var/folders/ab/T/a.xlsx",
        "/var/folders/ab/a",
        "C:\\Users\\x",
    ],
)
def test_a_probe_does_not_write_a_row_with_a_path(tmp_path, planted):
    out = tmp_path / "x.tsv"
    with pytest.raises(SystemExit, match="holds a path"):
        write_capture(out, ["# suite: x", "case\toutcome", f"a\trefused\t{planted}: not open"])
    assert not out.exists()
    write_capture(out, ["# suite: x", "case\toutcome", "a\topened"])
    assert out.read_text(encoding="utf-8").endswith("a\topened\n")
