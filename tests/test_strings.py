"""Escaped surrogates and empty text values read as Excel reads them.

Captures ``strings`` and ``strings_cached`` (``capture/strings_probe.py``): each case puts one
text in A1, as a shared string, an inline string or a formula's cached text, and Excel saved a
copy. The text Excel holds is the one its copy writes (a shared string, or A1's cached value),
and ``ISBLANK(A1)`` says when it holds none. Here the reader must give that text.
"""

from __future__ import annotations

import codecs
import io
import re
from pathlib import Path

import pytest

import xlio
from tests.test_read import package
from xlio._cells import unescape

CAPTURES = Path(__file__).with_name("captures")


def rows(suite: str) -> list[dict[str, str]]:
    lines = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
    header = next(line.split("\t") for line in lines if line.startswith("case\t"))
    return [
        dict(zip(header, line.split("\t"), strict=False))
        for line in lines
        if line.strip() and not line.startswith(("#", "case\t"))
    ]


STRINGS = rows("strings")
CACHED = rows("strings_cached")


def held(row: dict[str, str]) -> str | None:
    """The text Excel holds in A1, from what its copy writes; ``None`` for a blank cell."""
    if row["H1 ISBLANK(A1)"] == "b:1":
        return None
    a1 = row["A1"]
    if 't="s"' in a1:
        found = re.fullmatch(r"<si><t(?: [^>]*)?(?:/>|>(.*)</t>)</si>", row["sharedStrings"])
    else:
        found = re.search(r"<v(?:/>|>(.*)</v>)", a1)
    assert found is not None, row
    text = codecs.decode(found[1] or "", "unicode_escape")  # the probe's escapes undone
    return unescape(text)  # Excel writes a lone surrogate as an escape


def first_code(row: dict[str, str]) -> int | None:
    value = row["C1 _xlfn.UNICODE(A1)"]
    return int(value[2:]) if value.startswith("n:") else None


def test_the_captures_hold_every_case():
    from_strings = {r["case"].split(": ", 1)[-1] for r in STRINGS if ": " in r["case"]}
    from_cached = {r["case"].split(": ", 1)[-1] for r in CACHED}
    assert from_strings | {"zzz (control)"} == from_cached
    assert (len(STRINGS), len(CACHED)) == (16, 14)
    assert all(r["outcome"] == "opened" for r in STRINGS + CACHED)


@pytest.mark.parametrize("row", STRINGS, ids=[r["case"] for r in STRINGS])
def test_shared_and_inline_strings(row):
    pytest.importorskip("openpyxl")
    from capture.probe import build, control_workbook
    from capture.strings_probe import cases, fill

    case = {c.name: c for c in cases()}[row["case"]]
    data = build(control_workbook(fill), case)
    with xlio.open(io.BytesIO(data)) as wb:
        a1 = next(iter(wb.sheets[0].cells()), None)
    value = a1.value if a1 is not None and a1.type != "none" else None
    assert value == held(row)
    if value:
        assert ord(value[0]) == first_code(row)  # Excel's UNICODE(A1)


@pytest.mark.parametrize("row", CACHED, ids=[r["case"] for r in CACHED])
def test_a_formulas_cached_text(row):
    from capture.strings_probe import TEXTS

    name = row["case"].split(": ", 1)[1]
    text = "zzz" if name == "zzz (control)" else TEXTS[name]
    data = package(f'<row r="1"><c r="A1" t="str"><f>"q"</f><v>{text}</v></c></row>')
    with xlio.open(io.BytesIO(data)) as wb:
        (a1,) = wb.sheets[0].cells()
    assert a1.value == held(row)
    if row["case"].startswith("cached, B1:H1 uncached"):  # Excel computed them from A1
        assert ord(a1.value[0]) == first_code(row)
