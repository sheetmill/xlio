"""Attribute values Excel never writes, read as Excel reads them.

Capture ``attributes`` (``capture/attributes_probe.py``): each case is the probe's workbook with
one attribute rewritten. Excel either refused it (error -50) or opened it and saved a copy; the
TSV holds the elements the copy has. Here a refused case must be a ``PackageError``, and an
opened one must read as the copy states it, for the records xlio models.
"""

from __future__ import annotations

import io
import re
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.attributes_probe import cases, control
from capture.probe import build

CAPTURE = Path(__file__).with_name("captures") / "attributes.tsv"
LINES = CAPTURE.read_text(encoding="utf-8").splitlines()
HEADER = next(line.split("\t") for line in LINES if line.startswith("case\t"))
ROWS = [
    dict(zip(HEADER, line.split("\t"), strict=False))
    for line in LINES
    if line.strip() and not line.startswith(("#", "case\t"))
]
CONTROL = control()
BUILT = {case.name: case for case in cases()}
NOT_REFUSED: set[str] = set()  # cases Excel refuses that xlio reads (none)


def attribute(element: str, name: str) -> str | None:
    found = re.search(rf'\b{name}="([^"]*)"', element)
    return found.group(1) if found else None


def read_all(data: bytes) -> xlio.Workbook:
    wb = xlio.open(io.BytesIO(data))
    sheet = wb.sheets[0]
    sheet.unparsed  # noqa: B018 - the whole structure
    list(sheet.cells(rich=True))
    wb.styles  # noqa: B018
    return wb


def test_every_case_is_captured():
    assert sorted(r["case"] for r in ROWS) == sorted(BUILT)


@pytest.mark.parametrize("row", ROWS, ids=[r["case"] for r in ROWS])
def test_read_as_excel_reads_it(row):
    case = row["case"]
    data = build(CONTROL, BUILT[case])
    if row["outcome"] == "refused":
        if case in NOT_REFUSED:
            read_all(data)
            return
        with pytest.raises(xlio.PackageError):
            read_all(data)
        return
    wb = read_all(data)
    sheet = wb.sheets[0]
    styles = wb.styles
    if case.startswith("row hidden"):
        assert sheet.rows[1].hidden == (attribute(row["row 1"], "hidden") == "1")
    elif case.startswith("row customHeight"):
        assert sheet.rows[1].custom_height == (attribute(row["row 1"], "customHeight") == "1")
    elif case.startswith("row ht"):
        # Excel writes an ignored height as the default (15); an empty one is 0, hidden.
        saved = float(attribute(row["row 1"], "ht") or "nan")
        assert sheet.rows[1].height == (None if saved == 15 else saved)
    elif case.startswith("col"):
        (column,) = sheet.columns
        saved = row["cols"]
        assert column.width == float(attribute(saved, "width") or "nan")
        # Excel writes a zero width as hidden too; xlio keeps hidden as written.
        assert column.hidden is False
    elif case.startswith("cell"):
        (a1,) = (c for c in sheet.cells() if (c.row, c.column) == (1, 1))
        assert (a1.style, a1.cm, a1.vm) == (0, None, None)
    elif case.startswith("font color"):
        saved = re.search(r"<color\b[^>]*>", row["font 0"])
        color = styles.fonts[0].color
        got = (color.rgb, color.theme, color.indexed) if color else (None, None, None)
        if saved is None:
            assert got == (None, None, None)
        else:
            text = saved.group(0)
            expected = tuple(
                int(v) if v is not None and k != "rgb" else v
                for k, v in (("rgb", attribute(text, "rgb")), ("theme", attribute(text, "theme")),
                             ("indexed", attribute(text, "indexed")))
            )  # fmt: skip
            assert got == expected
            assert color is not None
            assert color.tint == pytest.approx(float(attribute(text, "tint") or 0), abs=1e-9)
    elif case.startswith("textRotation"):
        alignment = styles.cell_formats[1].alignment
        rotation = alignment.text_rotation if alignment is not None else 0
        assert rotation == int(attribute(row["alignment"], "textRotation") or 0)
    elif case.startswith("run sz"):
        (a2,) = (c for c in sheet.cells(rich=True) if (c.row, c.column) == (2, 1))
        assert a2.rich is not None
        first = re.search(r"<sz\b[^>]*>", row["A2 runs"])
        assert first is not None
        assert a2.rich.runs[0].font.size == float(attribute(first.group(0), "val") or "nan")
    elif case.startswith("dxf sz"):
        assert styles.differential_formats[0].font.size == float(
            attribute(row["dxfs"], "val") or "nan"
        )
    elif case.startswith("pageSetup"):
        saved = row["pageSetup"]
        setup = sheet.page_setup
        assert setup.copies == int(attribute(saved, "copies") or 1)
        assert setup.scale == int(attribute(saved, "scale") or 100)
        paper = attribute(saved, "paperSize")
        assert setup.paper_size == (int(paper) if paper is not None else None)


THEMES = [
    line.split("\t")[:2]
    for line in (CAPTURE.parent / "theme_bound.tsv").read_text(encoding="utf-8").splitlines()
    if line.startswith("theme ")
]


@pytest.mark.parametrize(("case", "outcome"), THEMES, ids=[t[0] for t in THEMES])
def test_the_theme_index_bound_is_excels(case, outcome):
    # Capture theme_bound: a font color's theme index of 11 opens, 12 and past are refused.
    from capture.theme_probe import case as theme_case

    data = build(CONTROL, theme_case(int(case.split()[1])))
    if outcome == "refused":
        with pytest.raises(xlio.PackageError, match="theme"):
            xlio.check(io.BytesIO(data))
        with pytest.raises(xlio.PackageError, match="theme"):
            xlio.open(io.BytesIO(data)).styles  # noqa: B018
    else:
        xlio.check(io.BytesIO(data))
