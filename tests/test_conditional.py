"""Conditional formats, with their Excel 2010 (``x14``) rules joined by ``x14:id``."""

from __future__ import annotations

import io

import pytest

import xlio
from tests.test_read import DOC, MAIN, PKG, REL, package

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
GUID = "{00000000-0000-0000-0000-000000000001}"
WORKBOOK_RELS = (
    f'<Relationships xmlns="{PKG}">'
    f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
    f'<Relationship Id="rId3" Type="{DOC}/styles" Target="styles.xml"/></Relationships>'
)
STYLES = (
    f'<styleSheet xmlns="{MAIN}"><dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font>'
    '<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>'
)


def sheet_of(body: str) -> xlio.Sheet:
    xml = (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}" xmlns:x14="{X14}" xmlns:xm="{XM}">'
        f"<sheetData/>{body}</worksheet>"
    )
    parts = {"xl/styles.xml": STYLES, "xl/_rels/workbook.xml.rels": WORKBOOK_RELS}
    return xlio.open(io.BytesIO(package(sheet_xml=xml, extra_parts=parts)))["Data"]


def extension(body: str, uri: str = xlio.ns.EXT_CONDITIONAL_FORMATTINGS) -> str:
    return (
        f'<extLst><ext uri="{uri}"><x14:conditionalFormattings>{body}'
        "</x14:conditionalFormattings></ext></extLst>"
    )


BASE = (
    '<conditionalFormatting sqref="A1:A10 C1:C10">'
    '<cfRule type="cellIs" dxfId="0" priority="2" operator="between" stopIfTrue="1">'
    "<formula>5</formula><formula>$B$1</formula></cfRule>"
    '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/>'
    '<color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">'
    f"<x14:id>{GUID}</x14:id></ext></extLst></cfRule></conditionalFormatting>"
    '<conditionalFormatting sqref="B1:B10">'
    '<cfRule type="containsText" dxfId="7" priority="3" operator="containsText" text="ab">'
    '<formula>NOT(ISERROR(SEARCH("ab",B1)))</formula></cfRule>'
    '<cfRule type="top10" priority="4" rank="5" percent="1" bottom="1"/>'
    '<cfRule type="aboveAverage" priority="5" aboveAverage="0" equalAverage="1"/>'
    '<cfRule type="timePeriod" priority="6" timePeriod="last7Days"/>'
    '<cfRule type="colorScale" priority="7"><colorScale><cfvo type="min"/>'
    '<cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/>'
    '<color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>'
    '<cfRule type="madeUp" priority="8"/></conditionalFormatting>'
)
EXTENDED = extension(
    f'<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{GUID}"><x14:dataBar '
    'minLength="0" maxLength="100" border="1" gradient="0" direction="rightToLeft" '
    'negativeBarBorderColorSameAsPositive="0" axisPosition="middle">'
    '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:fillColor rgb="FF638EC6"/>'
    '<x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/>'
    '<x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>'
    "</x14:cfRule><xm:sqref>A1:A10 C1:C10</xm:sqref></x14:conditionalFormatting>"
    '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="9" id="{X}">'
    '<x14:iconSet iconSet="3Stars" custom="1" showValue="0"><x14:cfvo type="percent">'
    '<xm:f>0</xm:f></x14:cfvo><x14:cfvo type="formula"><xm:f>Other!$A$1</xm:f></x14:cfvo>'
    '<x14:cfvo type="percent" gte="0"><xm:f>67</xm:f></x14:cfvo>'
    '<x14:cfIcon iconSet="3Stars" iconId="2"/><x14:cfIcon iconSet="NoIcons" iconId="0"/>'
    "</x14:iconSet></x14:cfRule>"
    '<x14:cfRule type="expression" priority="10" id="{Y}"><xm:f>Other!$A$1&gt;0</xm:f>'
    "<x14:dxf><font><b/></font></x14:dxf></x14:cfRule><xm:sqref>D1:D5</xm:sqref>"
    "</x14:conditionalFormatting>"
)


def test_base_rules_typed():
    formats = sheet_of(BASE).conditional_formats
    first, second = formats
    assert (first.sqref, first.extension, first.pivot) == ("A1:A10 C1:C10", False, False)
    between, bar = first.rules
    assert (between.type, between.priority, between.operator, between.stop_if_true) == (
        "cellIs", 2, "between", True,
    )  # fmt: skip
    assert between.formulas == ("5", "$B$1")
    assert between.dxf.fill.background.rgb == "FFFFC7CE"
    assert bar.data_bar.color == xlio.Color(rgb="FF638EC6")
    assert bar.data_bar.cfvos == (xlio.Cfvo("min"), xlio.Cfvo("max"))
    text, top, average, period, scale, made_up = second.rules
    assert (text.type, text.text, text.dxf_id, text.dxf) == ("containsText", "ab", 7, None)
    assert text.formulas == ('NOT(ISERROR(SEARCH("ab",B1)))',)  # the standard form, as written
    assert (top.rank, top.percent, top.bottom) == (5, True, True)
    assert (average.above_average, average.equal_average, average.std_dev) == (False, True, None)
    assert period.time_period == "last7Days"
    assert scale.color_scale.cfvos[1] == xlio.Cfvo("percentile", "50")
    assert [c.rgb for c in scale.color_scale.colors] == ["FFF8696B", "FFFFEB84", "FF63BE7B"]
    assert (made_up.type, made_up.source.get("type")) == (None, "madeUp")


def test_extension_rules_join_their_base_rules():
    sheet = sheet_of(BASE + EXTENDED)
    first, _, own = sheet.conditional_formats
    bar = first.rules[1]
    assert bar.id == GUID
    x14 = bar.x14.data_bar
    assert (x14.min_length, x14.max_length, x14.border, x14.gradient) == (0, 100, True, False)
    assert (x14.direction, x14.axis_position) == ("rightToLeft", "middle")
    assert not x14.negative_bar_color_same_as_positive
    assert not x14.negative_bar_border_color_same_as_positive
    assert x14.cfvos == (xlio.Cfvo("autoMin"), xlio.Cfvo("autoMax"))
    assert (x14.color.rgb, x14.negative_fill_color.rgb, x14.axis_color.rgb) == (
        "FF638EC6", "FFFF0000", "FF000000",
    )  # fmt: skip
    assert bar.x14.extension
    # Rules only in the extension stand on their own, after the base formats.
    assert (own.sqref, own.extension) == ("D1:D5", True)
    stars, expression = own.rules
    icons = stars.icon_set
    assert (icons.icon_set, icons.custom, icons.show_value) == ("3Stars", True, False)
    assert icons.cfvos == (
        xlio.Cfvo("percent", "0"),
        xlio.Cfvo("formula", "Other!$A$1"),
        xlio.Cfvo("percent", "67", gte=False),
    )
    assert icons.icons == (("3Stars", 2), ("NoIcons", 0))
    assert (expression.formulas, expression.dxf.font.bold, expression.priority) == (
        ("Other!$A$1>0",), True, 10,
    )  # fmt: skip
    assert sheet.unparsed == ()


def test_the_extension_uri_is_compared_exactly():
    # [MS-XLSX] prints 407D; Excel ignores the extension under it.
    printed = xlio.ns.EXT_CONDITIONAL_FORMATTINGS.replace("407d", "407D")
    sheet = sheet_of(BASE + EXTENDED.replace(xlio.ns.EXT_CONDITIONAL_FORMATTINGS, printed))
    first, _ = sheet.conditional_formats  # no extension-only format
    assert first.rules[1].x14 is None  # the id leads nowhere Excel reads
    (ext,) = sheet.extensions
    assert ext.uri == printed  # kept, unread


def test_no_conditional_formats_reads_no_styles():
    sheet = sheet_of("")
    assert sheet.conditional_formats == ()
    assert sheet.extensions == ()


# Files from other writers ---------------------------------------------------------------------


def test_what_openpyxl_writes(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
    from openpyxl.styles import Font, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    red = PatternFill(bgColor="FFC7CE")
    ws.conditional_formatting.add(
        "A1:A9", CellIsRule(operator="greaterThan", formula=["5"], fill=red)
    )
    ws.conditional_formatting.add(
        "B1:B9", FormulaRule(formula=["ISBLANK(B1)"], font=Font(bold=True))
    )
    ws.conditional_formatting.add(
        "C1:C9",
        ColorScaleRule(start_type="min", start_color="FF0000", end_type="max", end_color="00FF00"),
    )
    path = tmp_path / "o.xlsx"
    wb.save(path)
    with xlio.open(path) as book:
        a, b, c = book["Data"].conditional_formats
        (gt,) = a.rules
        assert (gt.type, gt.operator, gt.formulas) == ("cellIs", "greaterThan", ("5",))
        assert gt.dxf.fill.background.rgb == "00FFC7CE"
        (blank,) = b.rules
        assert (blank.type, blank.formulas, blank.dxf.font.bold) == (
            "expression",
            ("ISBLANK(B1)",),
            True,
        )
        (scale,) = c.rules
        assert [v.type for v in scale.color_scale.cfvos] == ["min", "max"]


def test_what_xlsxwriter_writes(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "x.xlsx"
    book = xlsxwriter.Workbook(path)
    ws = book.add_worksheet("Data")
    ws.conditional_format(
        "A1:A9",
        {"type": "data_bar", "bar_color": "#63C384", "bar_negative_color": "#FF0000",
         "bar_axis_position": "middle", "bar_direction": "right"},
    )  # fmt: skip
    ws.conditional_format("B1:B9", {"type": "text", "criteria": "begins with", "value": "ab",
                                    "format": book.add_format({"italic": True})})  # fmt: skip
    ws.conditional_format(
        "C1:C9", {"type": "icon_set", "icon_style": "4_arrows", "reverse_icons": True}
    )
    book.close()
    with xlio.open(path) as wb:
        sheet = wb["Data"]
        assert [e.uri for e in sheet.extensions] == [xlio.ns.EXT_CONDITIONAL_FORMATTINGS]
        bar_format, text_format, icon_format = sheet.conditional_formats  # all joined
        (bar,) = bar_format.rules
        assert bar.data_bar.color.rgb == "FF63C384"
        assert bar.x14 is not None
        assert (bar.x14.data_bar.axis_position, bar.x14.data_bar.direction) == (
            "middle",
            "rightToLeft",
        )
        assert bar.x14.data_bar.negative_fill_color.rgb == "FFFF0000"
        (begins,) = text_format.rules
        assert (begins.type, begins.text, begins.dxf.font.italic) == ("beginsWith", "ab", True)
        assert begins.formulas  # the standard formula, for the consumer to recognize
        (icons,) = icon_format.rules
        assert (icons.icon_set.icon_set, icons.icon_set.reverse) == ("4Arrows", True)
