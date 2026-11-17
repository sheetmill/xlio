"""A sheet's structure, the theme, the palette and the workbook's windows, as Excel reads them."""

from __future__ import annotations

import importlib.util
import io
import re
import zipfile
from pathlib import Path

import pytest

import xlio
from tests.test_read import MAIN, REL, package
from tests.watchdog import bounded

FILES = Path(__file__).with_name("files")
CAPTURE = Path(__file__).with_name("captures") / "package.tsv"


def sheet_of(xml_body: str, *, rels: str | None = None, **parts: str) -> xlio.Sheet:
    xml = f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{xml_body}</worksheet>'
    data = package(sheet_xml=xml, sheet_rels=rels, extra_parts=parts or None)
    return xlio.open(io.BytesIO(data))["Data"]


# The package capture: sizes Excel takes, and sizes it refuses --------------------------------


def capture_rows() -> list[list[str]]:
    return [
        line.split("\t")
        for line in CAPTURE.read_text(encoding="utf-8").splitlines()
        if not line.startswith(("#", "case\t"))
    ]


def shown_width(width: float) -> float:
    """Excel's ColumnWidth for a file width, at the captured Calibri 11 (6 points a digit,
    5 points of padding), for widths of one character or more."""
    return round((width * 6 - 5) / 6, 2)


def read_everything(data: bytes) -> xlio.Sheet:
    with xlio.open(io.BytesIO(data)) as wb:
        xlio.check(io.BytesIO(data))
        sheet = wb.sheets[0]
        list(sheet.cells())
        sheet.columns  # noqa: B018 - the structural pass
        return sheet


@pytest.mark.skipif(importlib.util.find_spec("openpyxl") is None, reason="needs openpyxl")
@bounded
def test_every_size_case_reads_as_excel_reads_it():
    from capture import package_probe

    cases = dict(package_probe.CASES)
    assert [r[0] for r in capture_rows()] == [name for name, _ in package_probe.CASES]
    base = package_probe.base()
    checked = 0
    for name, outcome, *sizes in capture_rows():
        if name.startswith("font_size"):
            continue  # the font, read with the styles
        data = cases[name](base)
        if outcome == "refused":
            with pytest.raises(xlio.PackageError):
                read_everything(data)
            checked += 1
            continue
        sheet = read_everything(data)
        _, _, width_a, height_1, width_b, _height_2, _font = sizes
        if name.startswith("width_"):
            (column,) = sheet.columns
            w = column.width
            assert w is not None
            if float(width_a) == 0:
                assert w == 0.0, name
            elif w < 1:
                assert (w, float(width_a)) == (0.1, 0.09), name  # one point
            else:
                assert shown_width(w) == float(width_a), (name, w)
        elif name.startswith("height_"):
            height = sheet.rows[1].height
            if height is None:
                assert float(height_1) == 15.0, name  # ignored: the row keeps its own height
            else:
                assert height == float(height_1), name
        elif name.startswith("default_width_"):
            if float(width_b) == 0:
                assert sheet.format.default_col_width == 0.0, name
            else:
                assert sheet.format.default_col_width is None, name
        elif name.startswith("custom_default_height_"):
            assert sheet.format.custom_height
            assert sheet.format.default_row_height in (float(height_1), 0.0), name
        checked += 1
    fonts = sum(1 for r in capture_rows() if r[0].startswith("font_size"))
    assert (checked, fonts) == (len(capture_rows()) - fonts, 10)


@pytest.mark.parametrize(
    ("attributes", "base", "width", "height"),
    [
        ('baseColWidth="12" defaultColWidth="20.5" defaultRowHeight="-18"', 12, 20.5, 18.0),
        ('baseColWidth="-1" defaultColWidth="abc" defaultRowHeight="15"', 8, None, 15.0),
        ('baseColWidth="99999999999" defaultRowHeight="1e308"', 8, None, 0.0),
    ],
)
def test_default_sizes(attributes, base, width, height):
    fmt = sheet_of(f"<sheetFormatPr {attributes}/><sheetData/>").format
    assert (fmt.base_col_width, fmt.default_col_width, fmt.default_row_height) == (
        base,
        width,
        height,
    )
    assert fmt.source is not None


# Columns, rows, merges -----------------------------------------------------------------------


def test_columns_rows_and_merges():
    sheet = sheet_of(
        '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"/>'
        '<cols><col min="1" max="3" width="20" style="4" customWidth="1" outlineLevel="1"/>'
        '<col min="4" max="4" width="-1" hidden="true"/></cols>'
        '<sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.2" xmlns:x14ac="urn:x"><c r="A1"/>'
        '</row><row r="2" ht="30" customHeight="1" s="3" customFormat="1" outlineLevel="2"/>'
        '<row hidden="1"/><row r="5" ht="1e4" collapsed="yes"/></sheetData>'
        '<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="C3:D4"/></mergeCells>'
    )
    first, second = sheet.columns
    assert (first.min, first.max, first.width, first.style, first.custom_width) == (
        1, 3, 20.0, 4, True,
    )  # fmt: skip
    assert first.outline_level == 1
    assert (second.width, second.raw_width, second.hidden) == (65531 / 6, "-1", True)
    assert sheet.unparsed == ()
    assert set(sheet.rows) == {2, 3, 5}  # row 1 says nothing beside its cells
    assert sheet.rows[2] == xlio.RowFormat(30.0, True, False, 3, True, 2, raw_height="30")
    assert sheet.rows[3].hidden
    assert (sheet.rows[5].height, sheet.rows[5].raw_height) == (169.5, "1e4")
    assert sheet.rows[5].collapsed is False  # not a boolean: the default, as for hidden
    assert sheet.merges == ("A1:B2", "C3:D4")
    assert sheet.format.outline_level_row == 2


ROWS_BODY = (
    '<sheetPr filterMode="1"/><sheetViews><sheetView workbookViewId="0"><selection '
    'activeCell="C3" sqref="C3"/></sheetView></sheetViews>'
    '<sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c></row>'
    '<row r="2" ht="30" customHeight="1"><c r="A2"><v>2</v></c></row>'
    '<row hidden="1"><c r="A3"><v>3</v></c></row></sheetData>'
    '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>'
)


def test_rows_are_gathered_by_a_whole_pass_of_cells():
    expected = sheet_of(ROWS_BODY).rows
    assert set(expected) == {2, 3}
    sheet = sheet_of(ROWS_BODY)
    assert [c.value for c in sheet.cells()] == [1.0, 2.0, 3.0]
    assert sheet.rows == expected
    assert "_structure" not in sheet.__dict__  # no second pass over the part
    assert sheet.merges == ("A1:B2",)  # the tail still reads, in a pass of its own
    assert sheet.rows == expected


def test_rows_are_not_taken_from_a_pass_of_cells_stopped_early():
    sheet = sheet_of(ROWS_BODY)
    cells = sheet.cells()
    next(cells)
    cells.close()
    assert set(sheet.rows) == {2, 3}
    assert "_structure" in sheet.__dict__


def test_the_head_is_read_without_the_cells_or_the_tail():
    sheet = sheet_of(ROWS_BODY)
    assert sheet.properties.filter_mode
    assert str(sheet.views[0].active_cell) == "C3"
    assert "_structure" not in sheet.__dict__
    assert sheet.merges == ("A1:B2",)
    assert str(sheet.views[0].active_cell) == "C3"


def test_a_head_record_after_the_cells_is_refused_by_the_structural_pass():
    # Excel refuses the file (capture views); the head alone, read up to the cells, cannot see it.
    sheet = sheet_of('<sheetData/><cols><col min="1" max="1" width="9"/></cols>')
    assert sheet.columns == ()
    with pytest.raises(xlio.PackageError, match="after the cells"):
        sheet.merges  # noqa: B018


@pytest.mark.parametrize("col", ['<col min="x" max="5"/>', '<col min="3" max="2"/>'])
def test_a_column_range_excel_refuses_is_refused(col):
    with pytest.raises(xlio.PackageError, match="column range"):
        sheet_of(f"<cols>{col}</cols><sheetData/>").columns  # noqa: B018


def test_a_row_off_the_sheet_is_refused_by_the_structural_pass_too():
    sheet = sheet_of('<sheetData><row r="1048577"/></sheetData>')
    with pytest.raises(xlio.PackageError, match="off the sheet"):
        sheet.rows  # noqa: B018


# Properties, views, panes ----------------------------------------------------------------------


def test_properties_and_views():
    sheet = sheet_of(
        '<sheetPr codeName="Sheet1" filterMode="1"><tabColor rgb="ff00b050"/>'
        '<outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>'
        '<dimension ref="A1:C9"/><sheetViews>'
        '<sheetView showGridLines="0" zoomScale="150" view="pageLayout" rightToLeft="1" '
        'defaultGridColor="0" colorId="10" workbookViewId="0">'
        '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>'
        '<selection pane="topRight" activeCell="B1" sqref="B1"/>'
        '<selection pane="bottomRight" activeCell="C5:D6" sqref="C5:D6"/>'
        "</sheetView></sheetViews><sheetData/>"
    )
    props = sheet.properties
    assert props.tab_color == xlio.Color(rgb="FF00B050")
    assert (props.summary_below, props.summary_right, props.fit_to_page) == (False, True, True)
    assert (props.filter_mode, props.code_name) == (True, "Sheet1")
    assert sheet.dimension == "A1:C9"
    (view,) = sheet.views
    assert (view.show_grid_lines, view.zoom_scale, view.view, view.right_to_left) == (
        False, 150, "pageLayout", True,
    )  # fmt: skip
    assert (view.default_grid_color, view.color_id) == (False, 10)
    assert view.pane == xlio.Pane(1.0, 2.0, "B3", "bottomRight", "frozen")
    assert view.active_cell == xlio.CellRef(5, 3)  # C5:D6 is not a cell: sqref's first
    assert sheet.unparsed == ()


def test_a_sheet_without_views_or_properties():
    sheet = sheet_of("<sheetData/>")
    assert (sheet.views, sheet.properties, sheet.format) == (
        (),
        xlio.SheetProperties(),
        xlio.SheetFormat(),
    )


def test_colors_as_written_and_typed():
    sheet = sheet_of('<sheetPr><tabColor theme="4" tint="-0.249977111117893"/></sheetPr>')
    assert sheet.properties.tab_color == xlio.Color(theme=4, tint=-0.249977111117893)
    odd = sheet_of('<sheetPr><tabColor rgb="red" indexed="-2" auto="1" tint="x"/></sheetPr>')
    color = odd.properties.tab_color
    # Text that is not hex is opaque black (capture attributes: "GGGGGGGG").
    assert (color.rgb, color.indexed, color.auto, color.tint) == ("FF000000", None, True, 0.0)
    assert color.source.get("rgb") == "red"  # the text as written


# Hyperlinks ----------------------------------------------------------------------------------


def test_hyperlinks_with_their_targets():
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'
        "</Relationships>"
    )
    sheet = sheet_of(
        '<sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Go"/>'
        '<hyperlink ref="B2" location="Data!C3" display="here"/>'
        '<hyperlink ref="C3" r:id="rId9"/></hyperlinks>',
        rels=rels,
    )
    a, b, c = sheet.hyperlinks
    assert (a.ref, a.target, a.tooltip, a.relationship_id) == (
        "A1", "https://example.com/", "Go", "rId1",
    )  # fmt: skip
    assert (b.target, b.location, b.display) == (None, "Data!C3", "here")
    assert (c.target, c.relationship_id) == (None, "rId9")  # a relationship that is missing


# What is not modeled ---------------------------------------------------------------------------


def test_every_child_not_modeled_is_listed_whole():
    sheet = sheet_of(
        '<sheetData/><sheetProtection sheet="1"/><conditionalFormatting sqref="A1">'
        '<cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule>'
        '</conditionalFormatting><phoneticPr fontId="1"/>'
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"/></extLst>'
    )
    assert [u.path for u in sheet.unparsed] == [
        "worksheet/sheetProtection",
        "worksheet/phoneticPr",
    ]
    assert sheet.unparsed[0].element.get("sheet") == "1"
    (ext,) = sheet.extensions  # modeled, and still listed whole
    assert ext.uri == xlio.ns.EXT_CONDITIONAL_FORMATTINGS
    (cf,) = sheet.conditional_formats
    assert cf.rules[0].formulas == ("TRUE",)


def test_the_workbook_lists_what_it_does_not_model():
    extra = (
        '<fileVersion appName="xl"/><workbookPr date1904="1" showObjects="placeholders" '
        'codeName="ThisWorkbook" defaultThemeVersion="166925"/>'
        '<bookViews><workbookView activeTab="2" firstSheet="1" showSheetTabs="0" '
        'tabRatio="750" xWindow="-20" windowWidth="28800"/></bookViews>'
        '<workbookProtection lockStructure="1"/>'
    )
    wb = xlio.open(io.BytesIO(package(workbook_extra=extra)))
    assert [u.path for u in wb.unparsed] == ["workbook/fileVersion", "workbook/workbookProtection"]
    props = wb.properties
    assert (props.date1904, props.show_objects, props.code_name) == (
        True,
        "placeholders",
        "ThisWorkbook",
    )
    assert props.default_theme_version == 166925
    (view,) = wb.views
    # activeTab 2 leads nowhere in a one-sheet workbook: the first sheet
    assert (view.active_tab, view.first_sheet, view.show_sheet_tabs, view.tab_ratio) == (
        0, 0, False, 750,  # activeTab and firstSheet both past the one sheet
    )  # fmt: skip
    assert (view.x_window, view.window_width, view.y_window) == (-20, 28800, None)


# Theme and palette ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", ["marker_typed.xlsx", "excel_formulas.xlsx"])
def test_the_theme_excel_saved(name):
    with zipfile.ZipFile(FILES / name) as z:
        xml = z.read("xl/theme/theme1.xml").decode()
    with xlio.open(FILES / name) as wb:
        theme = wb.theme
    assert theme is not None
    scheme = re.search(r"<a:clrScheme name=\"([^\"]*)\">(.*?)</a:clrScheme>", xml, re.S)
    assert theme.colors.name == scheme[1]
    for slot, body in re.findall(r"<a:(\w+)><a:(?:srgbClr|sysClr)([^>]*)/?>", scheme[2]):
        found = getattr(theme.colors, slot)
        written = re.search(r'(?:lastClr|val)="([0-9A-F]{6})"', body)
        assert found.rgb == written[1], slot
        if "sysClr" in xml.split(f"<a:{slot}>")[1][:20]:
            assert found.system is not None
    minor = re.search(r"<a:minorFont><a:latin typeface=\"([^\"]*)\"", xml)
    assert theme.minor_fonts.latin == minor[1]
    assert theme.minor_fonts.scripts.get("Jpan")
    assert theme.colors.by_index(0) == theme.colors.lt1  # Excel's order swaps the pairs
    assert theme.colors.by_index(1) == theme.colors.dk1
    assert theme.colors.by_index(4) == theme.colors.accent1
    assert theme.colors.by_index(12) is None


def test_no_theme_and_no_styles():
    wb = xlio.open(io.BytesIO(package()))
    assert wb.theme is None
    assert wb.styles == xlio.Styles()


def test_a_custom_palette():
    styles = (
        f'<styleSheet xmlns="{MAIN}"><colors><indexedColors><rgbColor rgb="ff000000"/>'
        '<rgbColor rgb="00FFFFFF"/><rgbColor rgb="nope"/></indexedColors>'
        '<mruColors><color rgb="FF123456"/></mruColors></colors></styleSheet>'
    )
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/></Relationships>'
    )
    data = package(extra_parts={"xl/styles.xml": styles, "xl/_rels/workbook.xml.rels": rels})
    wb = xlio.open(io.BytesIO(data))
    assert wb.styles.indexed_colors == ("FF000000", "00FFFFFF", None)
    assert wb.styles.mru_colors == (xlio.Color(rgb="FF123456"),)
    assert xlio.open(FILES / "marker_typed.xlsx").styles.indexed_colors is None


# Files from other writers --------------------------------------------------------------------


def test_what_openpyxl_writes(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "link"
    ws["A1"].hyperlink = "https://example.com/x"
    ws.column_dimensions["B"].width = 31.5
    ws.column_dimensions["C"].hidden = True
    ws.row_dimensions[3].height = 40
    ws.row_dimensions[4].hidden = True
    ws.merge_cells("D1:E2")
    ws.freeze_panes = "B2"
    ws.sheet_properties.tabColor = "1072BA"
    ws.sheet_view.zoomScale = 85
    wb.create_sheet("Hidden").sheet_state = "hidden"
    path = tmp_path / "o.xlsx"
    wb.save(path)
    with xlio.open(path) as book:
        sheet = book["Data"]
        widths = {(c.min, c.max): c.width for c in sheet.columns}
        assert widths[2, 2] == 31.5
        assert next(c for c in sheet.columns if c.min == 3).hidden
        assert (sheet.rows[3].height, sheet.rows[4].hidden) == (40.0, True)
        assert sheet.merges == ("D1:E2",)
        (view,) = sheet.views
        assert (view.zoom_scale, view.pane.state, view.pane.top_left_cell) == (85, "frozen", "B2")
        assert sheet.properties.tab_color.rgb == "001072BA"
        (link,) = sheet.hyperlinks
        assert (link.ref, link.target) == ("A1", "https://example.com/x")
        assert book["Hidden"].state == "hidden"
        assert book.theme is not None
        assert book.theme.colors.accent1.rgb is not None


def test_what_xlsxwriter_writes(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "x.xlsx"
    book = xlsxwriter.Workbook(path, {"date_1904": True})
    first = book.add_worksheet("First")
    second = book.add_worksheet("Second")
    second.hide_gridlines(2)
    second.set_zoom(75)
    second.right_to_left()
    second.set_tab_color("#FF9900")
    second.split_panes(15, 8.43)
    second.set_selection("C4")
    second.set_column("B:D", 12, None, {"level": 1, "hidden": True})
    second.set_row(5, 25)
    second.write_url("A1", "internal:First!A1", string="back")
    second.activate()
    first.write("A1", 1)
    book.close()
    with xlio.open(path) as wb:
        assert wb.properties.date1904
        assert wb.date1904
        assert wb.views[0].active_tab == 1
        sheet = wb["Second"]
        (view,) = sheet.views
        assert (view.show_grid_lines, view.zoom_scale, view.right_to_left) == (False, 75, True)
        assert view.tab_selected
        assert view.pane is not None
        assert view.pane.state == "split"
        assert sheet.properties.tab_color.rgb == "FFFF9900"
        (cols,) = sheet.columns
        assert (cols.min, cols.max, cols.width, cols.hidden, cols.outline_level) == (
            2, 4, 12.7109375, True, 1,
        )  # fmt: skip
        assert sheet.rows[6].height == float(sheet.rows[6].raw_height)  # XlsxWriter rounds
        (link,) = sheet.hyperlinks
        assert (link.location, link.target) == ("First!A1", None)
        assert wb.theme.minor_fonts.latin == "Calibri"


@pytest.mark.parametrize(
    "name", ["marker_typed.xlsx", "iterate_cached.xlsx", "excel_formulas.xlsx"]
)
def test_structures_excel_saved_agree_with_openpyxl(name):
    openpyxl = pytest.importorskip("openpyxl")
    ws = openpyxl.load_workbook(FILES / name).worksheets[0]
    with xlio.open(FILES / name) as wb:
        sheet = wb.sheets[0]
        assert sheet.views
        assert sheet.views[0].tab_selected
        assert str(sheet.views[0].active_cell) == (ws.sheet_view.selection[0].activeCell or "A1")
        for col in sheet.columns:
            theirs = ws.column_dimensions[openpyxl.utils.get_column_letter(col.min)].width
            assert col.width == theirs
        for number, row in sheet.rows.items():
            assert row.height == ws.row_dimensions[number].height
        assert sheet.format.default_row_height == ws.sheet_format.defaultRowHeight
        assert wb.views[0].active_tab == 0


@pytest.mark.parametrize(
    ("written", "shown", "text"),
    [
        (None, True, None),
        ("0", False, "0"),
        ("1", True, "1"),
        ("false", False, "false"),
        ("x", True, "x"),
    ],
)
def test_outline_symbols_with_their_text_as_written(written, shown, text):
    attribute = f' showOutlineSymbols="{written}"' if written is not None else ""
    sheet = sheet_of(
        f"<sheetPr><outlinePr{attribute}/></sheetPr>"
        f'<sheetViews><sheetView workbookViewId="0"{attribute}/></sheetViews><sheetData/>'
    )
    properties, view = sheet.properties, sheet.views[0]
    assert (properties.show_outline_symbols, properties.show_outline_symbols_text) == (shown, text)
    assert (view.show_outline_symbols, view.show_outline_symbols_text) == (shown, text)


def test_outline_symbols_without_outline_properties():
    properties = sheet_of("<sheetPr/><sheetData/>").properties
    assert (properties.show_outline_symbols, properties.show_outline_symbols_text) == (True, None)


EMPTY_ROWS = (
    '<sheetData><row r="1" ht="20" customHeight="1"><c r="A1"><v>1</v></c></row>'
    '<row r="3"/><row r="4" spans="1:2"></row><row r="5"><c r="A5"><v>5</v></c></row>'
    '<row r="6"><extLst/></row><row r="7" hidden="1"/></sheetData>'
)


def test_a_row_without_cells_or_formats_is_counted():
    expected = {
        1: xlio.RowFormat(height=20.0, custom_height=True, raw_height="20"),
        3: xlio.RowFormat(),
        4: xlio.RowFormat(),
        6: xlio.RowFormat(),
        7: xlio.RowFormat(hidden=True),
    }
    structural = sheet_of(EMPTY_ROWS)  # the rows from the structural pass
    assert dict(structural.rows) == expected
    streamed = sheet_of(EMPTY_ROWS)  # the rows gathered by a pass of cells
    assert [(c.row, c.value) for c in streamed.cells()] == [(1, 1.0), (5, 5.0)]
    assert dict(streamed.rows) == expected
    assert list(streamed.rows) == list(structural.rows) == [1, 3, 4, 6, 7]
