"""Every text capture holds no control character other than tab and newline.

The check reads each file's raw text, gzipped files decompressed, and splits it on ``\\n``
only: ``str.splitlines`` would also split on, and so hide, ``\\r``, ``\\x0b``, ``\\x0c``,
``\\x1c`` to ``\\x1e`` and ``\\x85``.
"""

from __future__ import annotations

import gzip
import re
from pathlib import Path

import pytest

CAPTURES = Path(__file__).parent / "captures"
TEXT_SUFFIXES = (".tsv", ".json")
FILES = sorted(
    p
    for p in CAPTURES.rglob("*")
    if p.is_file() and p.name.removesuffix(".gz").endswith(TEXT_SUFFIXES)
)

# C0 controls except tab and newline, DEL, and the C1 controls.
CONTROL = re.compile("[\x00-\x08\x0b-\x1f\x7f-\x9f]")


def _text(path: Path) -> str:
    data = path.read_bytes()
    if path.suffix == ".gz":
        data = gzip.decompress(data)
    return data.decode("utf-8")


def control_lines(text: str) -> list[tuple[int, str, str]]:
    """Each control character of ``text``: its line number, the character, the line.

    Split on newline alone: ``str.splitlines`` also splits on carriage return, form feed, the
    file, group and record separators and others, which would never reach the pattern.
    """
    return [
        (number, repr(match.group()), line[:60])
        for number, line in enumerate(text.split("\n"), 1)
        for match in CONTROL.finditer(line)
    ]


def test_every_capture_is_checked() -> None:
    assert len(FILES) >= 20
    assert not [p for p in CAPTURES.rglob("*") if p.is_file() and p not in FILES]


@pytest.mark.parametrize("path", FILES, ids=lambda p: str(p.relative_to(CAPTURES)))
def test_a_text_capture_holds_no_control_character(path: Path) -> None:
    found = control_lines(_text(path))
    assert not found, found


@pytest.mark.parametrize(
    "char",
    ["\x00", "\x0b", "\x0c", "\r", "\x1c", "\x1d", "\x1e", "\x1f", "\x7f", "\x80", "\x85", "\x9f"],
)
def test_every_control_character_is_found(char: str) -> None:
    text = f"case\tvalue\nrow\t1{char}2\tdone\n"
    assert control_lines(text) == [(2, repr(char), f"row\t1{char}2\tdone")]


def test_tab_newline_and_other_text_pass() -> None:
    assert control_lines("a\tb\nc\u00a0d\u2028e\u2029f\u20ac\n") == []


def test_a_gzipped_capture_is_read_decompressed(tmp_path: Path) -> None:
    path = tmp_path / "suite.tsv.gz"
    path.write_bytes(gzip.compress("a\tb\x85c\n".encode()))
    assert control_lines(_text(path)) == [(1, repr("\x85"), "a\tb\x85c")]
