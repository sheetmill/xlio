"""Page breaks, page setup, print areas and titles; outline directions and panes."""

from __future__ import annotations

import io

import pytest

import xlio
from tests.test_read import MAIN, REL, package

SHEETS = (
    '<sheet name="First" sheetId="7" r:id="rId1"/>'
    '<sheet name="Data" sheetId="3" r:id="rId1" state="hidden"/>'
)


def book(names: str = "", sheet_body: str = "", head: str = "") -> xlio.Workbook:
    sheets = f"{SHEETS}</sheets><definedNames>{names}</definedNames><sheets>"
    xml = f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{head}<sheetData/>{sheet_body}</worksheet>'
    return xlio.open(io.BytesIO(package(sheets=sheets, sheet_xml=xml)))


def name(text: str, sheet: int | None, kind: str = "Print_Area") -> str:
    scope = f' localSheetId="{sheet}"' if sheet is not None else ""
    return f'<definedName name="_xlnm.{kind}"{scope}>{text}</definedName>'


def test_breaks_and_page_setup():
    sheet = book(
        sheet_body='<printOptions gridLines="1" horizontalCentered="1"/>'
        '<pageMargins left="0.5" right="0.5" top="1" bottom="1" header="0.3"/>'
        '<pageSetup paperSize="9" orientation="landscape" scale="80" fitToHeight="0" '
        'firstPageNumber="3" useFirstPageNumber="1" errors="dash" copies="x"/>'
        '<headerFooter differentFirst="1"><oddHeader>&amp;L&amp;D&amp;R&amp;P</oddHeader>'
        "<firstFooter>first</firstFooter></headerFooter>"
        '<rowBreaks count="2" manualBreakCount="1"><brk id="10" max="16383" man="1"/>'
        '<brk id="40" max="16383"/></rowBreaks>'
        '<colBreaks count="1"><brk id="6" max="1048575" man="1"/></colBreaks>'
    )["First"]
    assert sheet.row_breaks == (
        xlio.PageBreak(10, manual=True, max=16383),
        xlio.PageBreak(40, max=16383),
    )
    assert sheet.column_breaks == (xlio.PageBreak(6, manual=True, max=1048575),)
    setup = sheet.page_setup
    assert (setup.paper_size, setup.orientation, setup.scale, setup.fit_to_height) == (
        9, "landscape", 80, 0,
    )  # fmt: skip
    assert (setup.first_page_number, setup.use_first_page_number, setup.errors) == (3, True, "dash")
    assert setup.copies == 1  # not a number: the default (capture attributes)
    assert setup.margins == xlio.PageMargins(0.5, 0.5, 1.0, 1.0, 0.3, None)
    assert setup.options == xlio.PrintOptions(grid_lines=True, horizontal_centered=True)
    hf = setup.header_footer
    assert (hf.odd_header, hf.first_footer, hf.different_first) == ("&L&D&R&P", "first", True)
    assert sheet.unparsed == ()


def test_no_page_setup():
    sheet = book()["First"]
    assert (sheet.row_breaks, sheet.page_setup) == ((), xlio.PageSetup())
    assert (sheet.print_area, sheet.print_titles) == (None, None)


def test_print_areas_by_sheet_position():
    # localSheetId counts sheets in workbook order, hidden ones included, not sheetId.
    wb = book(
        name("First!$A$1:$C$5,First!$E:$F,First!$1:$2,First!B7", 0)
        + name("'Data'!$C$3:$E$8,'Data'!$H$12:$K$16", 1)
    )
    assert wb["First"].print_area.ranges == ("$A$1:$C$5", "$E:$F", "$1:$2", "B7")
    data = wb["Data"].print_area
    assert (data.ranges, data.problem) == (("$C$3:$E$8", "$H$12:$K$16"), None)


@pytest.mark.parametrize(
    ("text", "ranges", "problem"),
    [
        ("First!#REF!", (), "#REF!"),
        ("First!$A$1:$B$2,First!#REF!", ("$A$1:$B$2",), "#REF!"),
        ("OFFSET(First!$A$1,0,0,5,5)", (), "a formula"),
        ("MyRange", (), "a formula"),
        ("Data!$A$1:$B$2", (), "a range of another sheet (Data)"),
        ("'It''s, here'!$A$1,First!$B$2", ("$B$2",), "a range of another sheet (It's, here)"),
        ("first!$A$1", ("$A$1",), None),  # sheet names without case
    ],
)
def test_print_area_problems(text, ranges, problem):
    area = book(name(text, 0))["First"].print_area
    assert (area.ranges, area.problem, area.text) == (ranges, problem, text)


def test_a_global_print_area_is_not_taken():
    wb = book(name("First!$A$1:$B$2", None))
    assert wb["First"].print_area is None
    assert wb.defined_names[0].name == "_xlnm.Print_Area"  # still visible


def test_print_titles():
    wb = book(
        name("First!$1:$2,First!$A:$B", 0, "Print_Titles")
        + name("Data!$A$1:$B$2", 1, "Print_Titles")
    )
    titles = wb["First"].print_titles
    assert (titles.rows, titles.columns, titles.problem) == ("$1:$2", "$A:$B", None)
    odd = wb["Data"].print_titles
    assert (odd.rows, odd.columns) == (None, None)
    assert odd.problem == "a range that is neither whole rows nor whole columns"


def test_the_outline_summary_directions():
    sheet = book(head='<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>').sheets[0]
    assert (sheet.properties.summary_below, sheet.properties.summary_right) == (False, False)


def test_what_openpyxl_writes(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.worksheet.pagebreak import Break

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.print_area = "A1:D20"
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_margins.left = 0.25
    ws.row_breaks.append(Break(id=10))
    ws.row_dimensions.group(3, 6, outline_level=1, hidden=True)
    ws.freeze_panes = "C3"
    path = tmp_path / "o.xlsx"
    wb.save(path)
    with xlio.open(path) as book_:
        sheet = book_["Data"]
        assert sheet.print_area.ranges == ("$A$1:$D$20",)
        assert sheet.print_titles.rows == "$1:$2"
        assert sheet.page_setup.orientation == "landscape"
        assert sheet.page_setup.margins.left == 0.25
        assert [b.id for b in sheet.row_breaks] == [10]
        assert all(sheet.rows[r].outline_level == 1 and sheet.rows[r].hidden for r in range(3, 7))
        assert sheet.views[0].pane.top_left_cell == "C3"


def test_what_xlsxwriter_writes(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "x.xlsx"
    wb = xlsxwriter.Workbook(path)
    wb.add_worksheet("First")
    ws = wb.add_worksheet("Data")
    ws.print_area("B2:F30")
    ws.repeat_rows(0, 1)
    ws.repeat_columns("A:A")
    ws.set_h_pagebreaks([15])
    ws.set_v_pagebreaks([4])
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_header("&CPage &P of &N")
    ws.hide_zero()
    ws.freeze_panes(1, 0)
    ws.set_row(1, None, None, {"level": 2, "collapsed": True})
    ws.outline_settings(True, False, True, False)
    wb.close()
    with xlio.open(path) as book_:
        sheet = book_["Data"]
        assert (sheet.index, sheet.print_area.ranges) == (1, ("$B$2:$F$30",))
        titles = sheet.print_titles
        assert (titles.rows, titles.columns) == ("$1:$2", "$A:$A")
        assert sheet.row_breaks[0] == xlio.PageBreak(15, manual=True, max=16383)
        assert sheet.column_breaks[0].id == 4
        setup = sheet.page_setup
        assert (setup.paper_size, setup.fit_to_width, setup.fit_to_height) == (9, 1, 0)
        assert sheet.properties.fit_to_page
        assert setup.header_footer.odd_header == "&CPage &P of &N"
        view = sheet.views[0]
        assert (view.show_zeros, view.pane.state, view.pane.y_split) == (False, "frozen", 1.0)
        assert (sheet.rows[2].outline_level, sheet.rows[2].collapsed) == (2, True)
        assert sheet.properties.summary_below is False
        assert book_["First"].print_area is None


def test_a_sheet_without_a_part_still_counts_for_local_names_and_tabs():
    # localSheetId and activeTab count <sheets> as written; xlio lists only sheets with a part.
    sheets = (
        '<sheet name="Gone" sheetId="1" r:id=""/><sheet name="Data" sheetId="2" r:id="rId1"/>'
        "</sheets><definedNames>"
        f"{name('Gone!$A$1:$B$2', 0)}{name('Data!$C$3:$D$4', 1)}</definedNames>"
        '<bookViews><workbookView activeTab="1" firstSheet="1"/></bookViews><sheets>'
    )
    xml = f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData/></worksheet>'
    wb = xlio.open(io.BytesIO(package(sheets=sheets, sheet_xml=xml)))
    (data,) = wb.sheets
    assert (data.name, data.index, data.position) == ("Data", 0, 1)
    assert data.print_area is not None
    assert (data.print_area.ranges, data.print_area.problem) == (("$C$3:$D$4",), None)
    assert (wb.views[0].active_tab, wb.views[0].first_sheet) == (0, 0)
