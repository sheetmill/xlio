"""Every capture the source cites is one a reader can open.

A comment that names a capture names a file in ``tests/captures/``; one kept by another project
("captured by xlrecalc", "xlrecalc's probe") comes with its link.
"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).parents[1]
SOURCE = sorted((ROOT / "src" / "xlio").glob("*.py"))
CITING = [
    *SOURCE,
    *sorted((ROOT / "capture").glob("*.py")),
    *(p for p in sorted((ROOT / "tests").glob("*.py")) if p.name != "test_citations.py"),
    ROOT / "CAPTURING.md",
]
POSSESSIVE = re.compile(r"\b(\w+)'s\s+(?:own\s+)?(?:probe|capture|captures|test|tests|runner)\b")
CAPTURES = {p.stem for p in (ROOT / "tests" / "captures").glob("*.tsv")}
NAMED = re.compile(r"\bcaptured?\s+`{0,2}([a-z][a-z_]*)`{0,2}")
WORDS = {"by", "for", "from", "with", "refers", "a", "the", "in", "as"}


def test_every_cited_capture_is_here():
    cited = {
        (path.name, name)
        for path in SOURCE
        for name in NAMED.findall(path.read_text(encoding="utf-8"))
        if name not in WORDS
    }
    assert cited
    assert {(f, n) for f, n in cited if n not in CAPTURES} == set()


def test_a_capture_kept_elsewhere_comes_with_its_link():
    for path in SOURCE:
        text = path.read_text(encoding="utf-8")
        for found in re.finditer(r"captured by (\w+)", text):
            after = text[found.end() : found.end() + 120]
            assert f"https://github.com/sheetmill/{found[1]}/" in after, (path.name, found[0])


def test_another_projects_probe_or_test_comes_with_its_link():
    for path in CITING:
        text = path.read_text(encoding="utf-8")
        for found in POSSESSIVE.finditer(text):
            if found[1] in ("Excel", "Ecma"):
                continue
            after = text[found.end() : found.end() + 160]
            assert f"https://github.com/sheetmill/{found[1]}/" in after, (path.name, found[0])
