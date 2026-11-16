"""Workbooks saved by Excel 16.115 (Microsoft 365), read whole.

``capture/fixtures.py`` built each with XlsxWriter and had Excel save it, or adopted one made by
hand in Excel (``handmade``); paths and people are stripped. What Excel wrote is what is read.
"""

from __future__ import annotations

from pathlib import Path

import pytest

import xlio

FILES = Path(__file__).with_name("files") / "excel"


def sheet(name: str) -> xlio.Sheet:
    return xlio.open(FILES / f"{name}.xlsx").sheets[0]


@pytest.mark.parametrize("name", ["tables", "conditional", "drawings", "notes", "handmade"])
def test_check_passes_and_nothing_is_left_unparsed(name):
    xlio.check(FILES / f"{name}.xlsx")
    assert sheet(name).unparsed == ()


def test_tables_and_the_sheet_filter():
    s = sheet("tables")
    sales, plain, no_header = s.tables
    assert (sales.name, sales.ref, sales.header_row_count, sales.totals_row_count) == (
        "Sales", "E1:F5", 1, 1,
    )  # fmt: skip
    assert sales.style is not None
    assert sales.style.name == "TableStyleMedium2"
    assert [c.name for c in sales.columns] == ["Id", "Amount"]
    assert plain.style is not None  # no style name, its options kept
    assert (plain.style.name, plain.style.show_first_column, plain.autofilter) == (None, True, None)
    assert (no_header.header_row_count, no_header.autofilter) == (0, None)
    assert s.autofilter is not None
    (column,) = s.autofilter.columns
    assert (s.autofilter.ref, column.column_id, column.values) == ("A1:C4", 0, ("East",))
    assert {r: f.hidden for r, f in s.rows.items()} == {3: True}


def test_conditional_formats_with_their_excel_2010_halves():
    s = sheet("conditional")
    kinds = {cf.sqref: [r.type for r in cf.rules] for cf in s.conditional_formats}
    assert kinds == {
        "A1:A10": ["cellIs", "cellIs"],
        "A1:E10": ["expression"],
        "B1:B10": ["colorScale"],
        "C1:C10": ["dataBar", "iconSet"],
        "D1:D10": ["containsText", "duplicateValues"],
        "E1:E10": ["dataBar", "top10", "aboveAverage"],
        "F1:F10": ["containsBlanks"],
    }
    priorities = sorted(r.priority for cf in s.conditional_formats for r in cf.rules)
    assert priorities == list(range(1, 13))
    bar = s.conditional_formats[3].rules[0]
    assert bar.data_bar is not None
    assert bar.extension is not None  # joined with its x14 rule


def test_drawings_in_z_order():
    s = sheet("drawings")
    assert s.drawing is not None
    objects = s.drawing.objects
    assert [o.kind for o in objects] == ["picture", "picture", "picture", "chart", "shape"]
    assert [o.anchor.edit_as for o in objects[:3]] == [None, "oneCell", "absolute"]
    assert {o.picture.embed_part for o in objects[:3] if o.picture} == {"xl/media/image1.png"}
    assert [o.description for o in objects[:3]] == ["anchor 1", "anchor 2", "anchor 3"]


def test_notes_and_their_vml():
    s = sheet("notes")
    assert [(n.ref, n.text.text) for n in s.comments] == [
        ("B1", "A note that is shown"), ("B2", "A hidden note"), ("B3", "A wide note"),
    ]  # fmt: skip
    assert s.legacy_drawing is not None
    shown = {(v.row, v.column): v.visible for v in s.legacy_drawing.shapes}
    assert shown == {(0, 1): True, (1, 1): False, (2, 1): False}


def test_threaded_comments_icons_and_a_picture_in_a_cell():
    s = sheet("handmade")
    first, reply = s.threaded_comments
    assert (first.ref, first.text, first.parent_id) == ("B2", "A comment", None)
    assert (reply.text, reply.parent_id) == ("A reply", first.id)
    assert first.person is not None
    assert first.person.display_name == "Writer"
    assert s.drawing is not None
    (icon,) = s.drawing.objects
    assert icon.picture is not None
    assert (icon.picture.embed_part, icon.picture.svg_part) == (
        "xl/media/image2.png", "xl/media/image3.svg",
    )  # fmt: skip
    (cell,) = s.cells()
    # A picture placed in E10: cached as #VALUE!, the picture in the rich value behind vm.
    assert (cell.row, cell.column, cell.value, cell.vm) == (10, 5, xlio.ErrorCode.VALUE, 1)
    picture = xlio.open(FILES / "handmade.xlsx").rich_values[1]
    assert isinstance(picture, xlio.ImageValue)
    assert (picture.type, picture.image_part, picture.calc_origin, picture.alt_text) == (
        "_localImage", "xl/media/image1.png", 5, None,
    )  # fmt: skip
