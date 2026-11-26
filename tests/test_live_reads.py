"""Readings checked against Excel while the workbook is open.

Capture ``live_reads`` (``capture/live_reads.py``): cases of the ``edges``, ``attributes``,
``outline`` and ``value_forms`` probes read through AppleScript while open, each answer as
``class:value``. Excel's dictionary types a font size as an integer and AppleScript gives it
cut to a whole number (12.5 reads 12), so a size is checked only as far as that goes; the
twentieths of a point are not settled by this capture.
"""

from __future__ import annotations

import io
import math
from pathlib import Path

import pytest

import xlio

pytest.importorskip("openpyxl")
from capture.live_reads import cases

CAPTURE = Path(__file__).with_name("captures") / "live_reads.tsv"
ANSWERS: dict[str, dict[str, str]] = {}
for line in CAPTURE.read_text(encoding="utf-8").splitlines():
    if line.startswith("#") or line.startswith("case\t") or not line.strip():
        continue
    case, outcome, _error, label, value = line.split("\t")
    assert outcome == "opened", line
    ANSWERS.setdefault(case, {})[label] = value
LIVE = {live.name: live for live in cases()}


def answer(case: str, label: str) -> str:
    return ANSWERS[case][label].partition(":")[2]


def test_every_case_is_captured():
    assert set(ANSWERS) == set(LIVE)


@pytest.mark.parametrize("name", sorted(LIVE), ids=sorted(LIVE))
def test_read_as_the_open_workbook_shows_it(name):
    got = ANSWERS[name]
    with xlio.open(io.BytesIO(LIVE[name].data)) as wb:
        sheet = wb.sheets[0]
        cells = {(c.row, c.column): c for c in sheet.cells(rich=True)}
        styles = wb.styles
        if "A1 font size" in got:
            size = styles.fonts[styles.cell_formats[cells[1, 1].style].font_id].size
            assert math.floor(size) == int(answer(name, "A1 font size"))
        if "A2 character 1 font size" in got:
            rich = cells[2, 1].rich
            assert rich is not None
            assert rich.runs[0].font is not None
            run = rich.runs[0].font.size
            assert run is not None
            assert math.floor(run) == int(answer(name, "A2 character 1 font size"))
        if "A3 condition 1 font size" in got:
            font = styles.differential_formats[0].font
            assert font is not None
            assert font.size is not None
            assert math.floor(font.size) == int(answer(name, "A3 condition 1 font size"))
        if "row 1 hidden" in got:
            row = sheet.rows.get(1)
            hidden = row.hidden if row else False
            assert hidden == (answer(name, "row 1 hidden") == "true")
            if row is not None and row.custom_height and row.height is not None:
                assert row.height == float(answer(name, "row 1 height"))
            (column,) = sheet.columns
            assert column.hidden == (answer(name, "column B hidden") == "true")
            if column.width == 0:
                assert float(answer(name, "column B width")) == 0
        if "row 2 outline level" in got:  # Excel's object model counts levels from 1
            row = sheet.rows.get(2)
            column = {c.min: c for c in sheet.columns}.get(2)
            assert (row.outline_level if row else 0) + 1 == int(answer(name, "row 2 outline level"))
            level = column.outline_level if column else 0
            assert level + 1 == int(answer(name, "column B outline level"))
        if "A1 value" in got:
            value = cells[1, 1].value
            # AppleScript gives a real to 13 significant digits
            assert value == pytest.approx(float(answer(name, "A1 value")), rel=1e-12, abs=0)
            assert (value > 0) == (answer(name, "C1 value") == "true")
