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


WINDOWS = rows("windows/csv_open.tsv")


def by_case(table: list[list[str]]) -> dict[str, list[list[str]]]:
    out: dict[str, list[list[str]]] = {}
    for r in table:
        out.setdefault(r[0], []).append(r[1:7])
    return out


def test_windows_reads_as_the_mac_but_for_the_code_page():
    # The cases whose bytes Windows-1252 and Mac Roman read differently differ, and the three
    # bracket names, which the Mac twin could not read; everything else is the same row.
    mac, win = by_case(MAC), by_case(WINDOWS)
    assert set(mac) == set(win)
    differ = {case for case in win if win[case] != mac[case]}
    assert differ == {
        "bytes 80-ff", "words no mark", "words utf-16-be mark", "currency windows-1252",
        "name bracket", "name bracket first", "name closing bracket",
    }  # fmt: skip


def test_windows_opens_a_bracketed_name_with_parentheses():
    names = {r[0]: r[3] for r in WINDOWS if r[0].startswith("name bracket") or "closing" in r[0]}
    assert names == {
        "name bracket": "a(b)c",
        "name bracket first": "(ab)",
        "name closing bracket": "a)b",
    }


def test_the_workbook_excel_makes_has_the_platform_standard_font():
    assert {r[7] for r in MAC} - {""} == {"Aptos Narrow 12"}
    assert {r[7] for r in WINDOWS} == {"Aptos Narrow 11"}
