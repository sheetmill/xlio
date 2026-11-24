"""The ``csv_open`` capture (``capture/csv_job.py``): its Mac twin against the older captures."""

from __future__ import annotations

from pathlib import Path

import pytest

CAPTURES = Path(__file__).with_name("captures")


def rows(name: str) -> list[list[str]]:
    lines = (CAPTURES / name).read_text(encoding="utf-8").split("\n")
    return [line.split("\t") for line in lines if line and not line.startswith("#")][1:]


MAC = rows("csv_open.tsv")


def cells(case: str) -> list[list[str]]:
    """``cell, value, format`` of each row of a case."""
    return [r[4:7] for r in MAC if r[0] == case]


def test_every_case_is_captured():
    from capture.csv_job import cases

    assert {r[0] for r in MAC} == set(cases())


@pytest.mark.parametrize(
    ("case", "older"),
    [
        ("words no mark", "encoding utf-8"),
        ("words utf-8 mark", "encoding utf-8 with BOM"),
        ("words utf-16-le mark", "encoding utf-16-le with BOM"),
        ("words utf-16-be mark", "encoding utf-16-be with BOM"),
    ],
)
def test_the_encodings_read_as_in_the_text_capture(case, older):
    assert cells(case) == [r[-3:] for r in rows("text.tsv") if r[0] == older and r[-3] != "font"]


@pytest.mark.parametrize(
    "case",
    ["years", "name 31 characters", "name 32 characters", "name 40 characters",
     "name History", "name apostrophe first"],
)  # fmt: skip
def test_years_and_names_read_as_in_text_extra(case):
    # sheet, cell, value, format: text_extra has no error column
    older = [r[2:6] for r in rows("text_extra.tsv") if r[0] == case]
    assert [r[3:7] for r in MAC if r[0] == case] == older


def test_forms_read_as_in_text_forms():
    assert cells("forms") == [r[-3:] for r in rows("text_forms.tsv")]


def test_limits_read_as_in_text_limits():
    for older in rows("text_limits.tsv"):
        (row,) = [r for r in MAC if r[0] == f"limits {older[0]}"]
        assert row[5] == " ".join(older[3:7]), older[0]


def test_the_bracket_names_are_the_ones_not_read():
    # Excel for Mac renames a[b]c.csv to a(b)c.csv, which AppleScript cannot address;
    # text_extra holds that reading, made by hand.
    assert {r[0] for r in MAC if r[1] == "not read"} == {
        "name bracket", "name bracket first", "name closing bracket",
    }  # fmt: skip
