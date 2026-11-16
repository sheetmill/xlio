"""Tables, table styles and their presets, autofilters and sort state."""

from __future__ import annotations

import hashlib
import io
import re
from pathlib import Path

import pytest

import xlio
from tests.test_read import DOC, MAIN, PKG, REL, package

PRESETS = Path(xlio.__file__).with_name("data") / "presetTableStyles.xml"


# The presets ----------------------------------------------------------------------------------


def test_the_presets_are_ecmas_file_byte_for_byte():
    # The file as Ecma publishes it (NOTICE names the source and this digest).
    digest = hashlib.sha256(PRESETS.read_bytes()).hexdigest()
    assert digest == "74c6cf73d551bb2bd33a51690d4922095260b6647b2a38fcbc0d6da6c6275d7e"
    notice = (Path(__file__).parents[1] / "NOTICE").read_text(encoding="utf-8")
    assert digest in notice


def test_144_presets():
    names = xlio.presets.names()
    assert len(names) == 144
    tables = [n for n in names if xlio.presets.table_style(n).table]
    assert len(tables) == 60
    for prefix, count in (("Light", 21), ("Medium", 28), ("Dark", 11)):
        found = {n for n in tables if re.fullmatch(f"TableStyle{prefix}[0-9]+", n)}
        assert found == {f"TableStyle{prefix}{k}" for k in range(1, count + 1)}
    assert sum(1 for n in names if n.startswith("PivotStyle")) == 84
    for name in names:
        style = xlio.presets.table_style(name)
        assert style.builtin
        assert style.unparsed == ()
        assert all(e.dxf is not None for e in style.elements), name  # every id resolves
    assert xlio.presets.table_style("tablestylemedium2") is None  # exact names
    # Two inner names in the file differ from their presets' names.
    assert xlio.presets.table_style("TableStyleLight14").source.get("name") == "TableStyleLight14 "
    assert xlio.presets.table_style("TableStyleDark9").source.get("name") == "TableStyleDark9 2"


def test_a_preset_reads_as_excel_draws_it():
    # TableStyleMedium2, as Excel shows it: a header of accent 1 with
    # bold text in background 1 (white under the default theme), bands of accent 1 lightened.
    style = xlio.presets.table_style("TableStyleMedium2")
    header = style.element("headerRow").dxf
    assert (header.font.bold, header.font.color, header.fill.foreground) == (
        True, xlio.Color(theme=0), xlio.Color(theme=4),
    )  # fmt: skip
    band = style.element("firstRowStripe").dxf
    assert (band.fill.foreground.theme, band.fill.foreground.tint) == (
        4,
        pytest.approx(0.8, abs=1e-3),
    )
    assert style.element("blankRow") is None


def test_pivot_presets_count_their_formats_from_zero():
    style = xlio.presets.table_style("PivotStyleLight16")
    assert style.table is False
    assert all(e.dxf is not None for e in style.elements)
    assert min(e.dxf_id for e in style.elements) == 0


# A workbook's own table styles ----------------------------------------------------------------

WORKBOOK_RELS = (
    f'<Relationships xmlns="{PKG}">'
    f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
    f'<Relationship Id="rId3" Type="{DOC}/styles" Target="styles.xml"/></Relationships>'
)


def test_own_table_styles_come_first():
    styles = (
        f'<styleSheet xmlns="{MAIN}"><dxfs count="2"><dxf><font><b/></font></dxf>'
        '<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>'
        '<tableStyles count="2" defaultTableStyle="Mine" defaultPivotStyle="PivotStyleLight16">'
        '<tableStyle name="Mine" pivot="0" count="2">'
        '<tableStyleElement type="headerRow" dxfId="0"/>'
        '<tableStyleElement type="firstRowStripe" dxfId="1" size="2"/>'
        '<tableStyleElement type="madeUp" dxfId="0"/></tableStyle>'
        '<tableStyle name="TableStyleLight1" count="1"><tableStyleElement type="wholeTable" '
        'dxfId="9"/></tableStyle></tableStyles></styleSheet>'
    )
    parts = {"xl/styles.xml": styles, "xl/_rels/workbook.xml.rels": WORKBOOK_RELS}
    wb = xlio.open(io.BytesIO(package(extra_parts=parts)))
    own = wb.styles.table_styles
    assert (own.default_table_style, own.default_pivot_style) == ("Mine", "PivotStyleLight16")
    mine = wb.styles.table_style("Mine")
    assert (mine.pivot, mine.table, mine.builtin) == (False, True, False)
    assert mine.element("headerRow").dxf.font.bold
    stripe = mine.element("firstRowStripe")
    assert (stripe.size, stripe.dxf.fill.background.rgb) == (2, "FFFFFF00")
    assert [u.element.get("type") for u in mine.unparsed] == ["madeUp"]
    shadow = wb.styles.table_style("TableStyleLight1")  # the file's own, named like a preset
    assert (shadow.builtin, shadow.elements[0].dxf) == (False, None)  # dxfId 9 leads nowhere
    assert wb.styles.table_style("TableStyleDark1").builtin


# Autofilters and sort state -------------------------------------------------------------------


def sheet_of(body: str, rels: str | None = None, **parts: str) -> xlio.Sheet:
    xml = f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData/>{body}</worksheet>'
    data = package(sheet_xml=xml, sheet_rels=rels, extra_parts=parts or None)
    return xlio.open(io.BytesIO(data))["Data"]


FILTER = (
    '<autoFilter ref="A1:F20">'
    '<filterColumn colId="0"><filters blank="1"><filter val="North"/><filter val="South"/>'
    '<dateGroupItem year="2024" month="3" dateTimeGrouping="month"/></filters></filterColumn>'
    '<filterColumn colId="1" hiddenButton="1"><customFilters and="1">'
    '<customFilter operator="greaterThan" val="10"/><customFilter operator="lessThan" val="99"/>'
    "</customFilters></filterColumn>"
    '<filterColumn colId="2"><top10 top="0" percent="1" val="5" filterVal="1.5"/></filterColumn>'
    '<filterColumn colId="3"><dynamicFilter type="aboveAverage" val="12.5"/></filterColumn>'
    '<filterColumn colId="4" showButton="0"><colorFilter dxfId="3" cellColor="0"/></filterColumn>'
    '<filterColumn colId="5"><iconFilter iconSet="3Arrows" iconId="2"/></filterColumn>'
    '<sortState ref="A2:F20"><sortCondition descending="1" ref="B2:B20"/>'
    '<sortCondition sortBy="cellColor" ref="C2:C20" dxfId="1"/></sortState></autoFilter>'
)


def test_an_autofilter_with_every_kind_of_filter():
    f = sheet_of(FILTER).autofilter
    assert f.ref == "A1:F20"
    values, custom, top, dynamic, color, icon = f.columns
    assert (values.values, values.blank, values.date_groups) == (
        ("North", "South"), True, (xlio.DateGroup("month", 2024, 3),),
    )  # fmt: skip
    assert (custom.hidden_button, custom.custom_and) == (True, True)
    assert custom.custom == (
        xlio.CustomFilter("greaterThan", "10"),
        xlio.CustomFilter("lessThan", "99"),
    )
    assert top.top10 == xlio.Top10(top=False, percent=True, value=5.0, filter_value=1.5)
    assert dynamic.dynamic == xlio.DynamicFilter("aboveAverage", 12.5)
    assert (color.show_button, color.color) == (False, xlio.ColorFilter(3, cell_color=False))
    assert icon.icon == xlio.IconFilter("3Arrows", 2)
    assert f.sort_state.conditions == (
        xlio.SortCondition("B2:B20", descending=True),
        xlio.SortCondition("C2:C20", sort_by="cellColor", dxf_id=1),
    )


def test_a_sort_outside_a_filter():
    sheet = sheet_of('<sortState ref="A1:B9" columnSort="1" sortMethod="stroke">'
                     '<sortCondition ref="A1:A9" sortBy="icon" iconSet="3Flags" iconId="0"/>'
                     "</sortState>")  # fmt: skip
    state = sheet.sort_state
    assert (state.ref, state.column_sort, state.sort_method) == ("A1:B9", True, "stroke")
    assert state.conditions[0] == xlio.SortCondition(
        "A1:A9", sort_by="icon", icon_set="3Flags", icon_id=0
    )
    assert sheet.autofilter is None
    assert sheet.unparsed == ()


# Table parts ---------------------------------------------------------------------------------

TABLE = (
    f'<table xmlns="{MAIN}" id="3" name="Table1" displayName="Sales" ref="A1:C6" '
    'totalsRowCount="1" headerRowDxfId="4">'
    '<autoFilter ref="A1:C5"><filterColumn colId="1" hiddenButton="1"/></autoFilter>'
    '<sortState ref="A2:C5"><sortCondition ref="A2:A5"/></sortState>'
    '<tableColumns count="3"><tableColumn id="1" name="Region" totalsRowLabel="Total"/>'
    '<tableColumn id="2" name="Units" totalsRowFunction="sum" dataDxfId="2"/>'
    '<tableColumn id="3" name="Price" totalsRowFunction="custom">'
    "<calculatedColumnFormula>Sales[[#This Row],[Units]]*2</calculatedColumnFormula>"
    "<totalsRowFormula>SUBTOTAL(109,Sales[Price])</totalsRowFormula></tableColumn></tableColumns>"
    '<tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/><extLst/></table>'
)
TABLE_RELS = (
    f'<Relationships xmlns="{PKG}">'
    f'<Relationship Id="rId1" Type="{DOC}/table" Target="../tables/table1.xml"/>'
    f'<Relationship Id="rId2" Type="{DOC}/table" Target="../tables/missing.xml"/></Relationships>'
)


def test_a_table_part():
    sheet = sheet_of(
        '<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>',
        TABLE_RELS,
        **{"xl/tables/table1.xml": TABLE},
    )
    (table,) = sheet.tables  # rId2 leads nowhere
    assert (table.name, table.display_name, table.ref, table.id) == ("Table1", "Sales", "A1:C6", 3)
    assert (table.header_row_count, table.totals_row_count, table.header_row_dxf_id) == (1, 1, 4)
    region, units, price = table.columns
    assert (region.name, region.totals_row_label) == ("Region", "Total")
    assert (units.totals_row_function, units.data_dxf_id) == ("sum", 2)
    assert price.calculated_formula == "Sales[[#This Row],[Units]]*2"
    assert price.totals_row_formula == "SUBTOTAL(109,Sales[Price])"
    # Absent options are off, as Excel draws them.
    assert table.style == xlio.TableStyleInfo("TableStyleMedium9", show_row_stripes=True)
    assert table.autofilter.columns[0].hidden_button
    assert table.sort_state.conditions == (xlio.SortCondition("A2:A5"),)
    assert [u.path for u in table.unparsed] == ["table/extLst"]
    assert table.part == "xl/tables/table1.xml"
    assert sheet.unparsed == ()


# Files from other writers ---------------------------------------------------------------------


def test_what_openpyxl_writes(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [["Name", "Score"], ["a", 1], ["b", 2]]:
        ws.append(row)
    table = Table(displayName="Scores", ref="A1:B3")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(table)
    other = wb.create_sheet("Other")
    other.append(["k", "v"])
    other.auto_filter.ref = "A1:B5"
    other.auto_filter.add_filter_column(0, ["x", "y"])
    other.auto_filter.add_sort_condition("B2:B5", descending=True)
    path = tmp_path / "o.xlsx"
    wb.save(path)
    with xlio.open(path) as book:
        (t,) = book["Data"].tables
        assert (t.display_name, t.ref, [c.name for c in t.columns]) == (
            "Scores", "A1:B3", ["Name", "Score"],
        )  # fmt: skip
        assert t.style.name == "TableStyleLight9"
        assert book.styles.table_style(t.style.name).builtin
        f = book["Other"].autofilter
        assert (f.ref, f.columns[0].values) == ("A1:B5", ("x", "y"))
        assert f.sort_state.conditions[0].descending


def test_what_xlsxwriter_writes(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "x.xlsx"
    book = xlsxwriter.Workbook(path)
    ws = book.add_worksheet("Data")
    ws.add_table(
        "B2:D6",
        {"name": "Sales", "style": "Table Style Dark 3", "total_row": True, "banded_rows": False,
         "first_column": True,
         "columns": [{"header": "Item", "total_string": "Total"},
                     {"header": "Qty", "total_function": "sum"},
                     {"header": "Price", "formula": "=[@Qty]*2"}]},
    )  # fmt: skip
    other = book.add_worksheet("Other")
    other.autofilter("A1:C9")
    other.filter_column("A", "x > 2000 and x < 5000")
    other.filter_column_list("B", ["East", "West"])
    book.close()
    with xlio.open(path) as wb:
        (t,) = wb["Data"].tables
        assert (t.name, t.ref, t.totals_row_count) == ("Sales", "B2:D6", 1)
        assert t.style == xlio.TableStyleInfo("TableStyleDark3", show_first_column=True)
        item, qty, price = t.columns
        assert (item.totals_row_label, qty.totals_row_function) == ("Total", "sum")
        assert price.calculated_formula == "[[#This Row],Qty]*2"  # as XlsxWriter writes it
        f = wb["Other"].autofilter
        a, b = f.columns
        assert (a.custom_and, [c.operator for c in a.custom]) == (True, ["greaterThan", "lessThan"])
        assert b.values == ("East", "West")
