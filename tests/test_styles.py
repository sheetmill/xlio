"""Styles with their indices resolved, font sizes as Excel reads them, and rich text."""

from __future__ import annotations

import importlib.util
import io
from pathlib import Path

import pytest

import xlio
from tests.test_read import DOC, MAIN, PKG, package

FILES = Path(__file__).with_name("files")
CAPTURE = Path(__file__).with_name("captures") / "package.tsv"
WORKBOOK_RELS = (
    f'<Relationships xmlns="{PKG}">'
    f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
    f'<Relationship Id="rId3" Type="{DOC}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId9" Type="{DOC}/sharedStrings" Target="sharedStrings.xml"/>'
    "</Relationships>"
)
needs_openpyxl = pytest.mark.skipif(
    importlib.util.find_spec("openpyxl") is None, reason="needs openpyxl"
)


def with_styles(body: str, rows: str = "", strings: list[str] | None = None) -> xlio.Workbook:
    parts = {
        "xl/styles.xml": f'<styleSheet xmlns="{MAIN}">{body}</styleSheet>',
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
    }
    return xlio.open(io.BytesIO(package(rows, strings=strings, extra_parts=parts)))


# Font sizes, as captured ----------------------------------------------------------------------


@needs_openpyxl
def test_font_sizes_as_excel_reads_them():
    from capture import package_probe

    cases = dict(package_probe.CASES)
    base = package_probe.base()
    rows = [
        line.split("\t")
        for line in CAPTURE.read_text(encoding="utf-8").splitlines()
        if line.startswith("font_size")
    ]
    assert len(rows) == 10
    for name, outcome, *sizes in rows:
        data = cases[name](base)
        wb = xlio.open(io.BytesIO(data))  # styles are read when asked for
        if outcome == "refused":
            with pytest.raises(xlio.PackageError, match="font size"):
                wb.styles  # noqa: B018
            with pytest.raises(xlio.PackageError, match="font size"):
                xlio.check(io.BytesIO(data))  # would Excel open it: no
        else:
            assert wb.styles.fonts[0].size == float(sizes[-1]), name
            xlio.check(io.BytesIO(data))


def test_check_reads_the_font_sizes_of_fonts_and_differential_formats():
    # A differential format's size of 0 is refused by Excel (capture attributes); -2 is 2.
    with pytest.raises(xlio.PackageError, match="differential font size of '0'"):
        xlio.check(
            io.BytesIO(
                package(extra_parts=_parts('<dxfs><dxf><font><sz val="0"/></font></dxf></dxfs>'))
            )
        )
    xlio.check(
        io.BytesIO(
            package(extra_parts=_parts('<dxfs><dxf><font><sz val="-2"/></font></dxf></dxfs>'))
        )
    )
    with pytest.raises(xlio.PackageError, match="font size of '410'"):
        xlio.check(
            io.BytesIO(package(extra_parts=_parts('<fonts><font><sz val="410"/></font></fonts>')))
        )


def _parts(body: str) -> dict[str, str]:
    return {
        "xl/styles.xml": f'<styleSheet xmlns="{MAIN}">{body}</styleSheet>',
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
    }


@pytest.mark.parametrize(("text", "size"), [("11", 11.0), ("-12.5", 12.5), (" 409 ", 409.0)])
def test_a_font_size_is_its_magnitude(text, size):
    wb = with_styles(f'<fonts><font><sz val="{text}"/></font></fonts>')
    assert (wb.styles.fonts[0].size, wb.styles.fonts[0].raw_size) == (size, text)


@pytest.mark.parametrize(
    ("text", "size"), [("409.5", 409.5), ("-409.5", 409.5), ("409.55", 409.55), ("409.59", 409.55)]
)
def test_a_font_size_past_409_opens(text, size):
    # Excel opens sizes below 409.6, kept in twentieths of a point, cut down; 409.6 is refused
    # (capture edges; tests/captures/windows/font_sizes.tsv).
    wb = with_styles(f'<fonts><font><sz val="{text}"/></font></fonts>')
    assert (wb.styles.fonts[0].size, wb.styles.fonts[0].raw_size) == (size, text)
    with pytest.raises(xlio.PackageError, match=r"font size of '409\.6'"):
        xlio.check(
            io.BytesIO(package(extra_parts=_parts('<fonts><font><sz val="409.6"/></font></fonts>')))
        )


def test_start_and_end_are_drawn_over_left_and_right():
    # Excel keeps start over a left written after it, and end over a right; alone, start and
    # end are the left and right sides (capture edges).
    wb = with_styles(
        "<borders>"
        '<border><start style="thick"/><end style="thick"/><left style="thin"/>'
        '<right style="dashed"/><top/><bottom/><diagonal/></border>'
        '<border><start style="double"/><end style="mediumDashed"/><top/><bottom/></border>'
        "</borders>"
    )
    both, alone = wb.styles.borders
    assert (both.left.style, both.right.style) == ("thick", "thick")
    assert (alone.left.style, alone.right.style) == ("double", "mediumDashed")


# The tables, crafted --------------------------------------------------------------------------

STYLES = (
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000%"/></numFmts>'
    '<fonts><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/>'
    '<scheme val="minor"/></font><font><b/><i val="0"/><u/><strike val="true"/>'
    '<vertAlign val="superscript"/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/>'
    "</font></fonts>"
    '<fills><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/>'
    '</fill><fill><patternFill patternType="solid"><fgColor rgb="FF00B0F0"/><bgColor indexed="64"/>'
    '</patternFill></fill><fill><gradientFill degree="90"><stop position="0"><color theme="0"/>'
    '</stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>'
    "<fill><patternFill/></fill></fills>"
    '<borders><border><left/><right/><top/><bottom/><diagonal/></border><border diagonalUp="1">'
    '<left style="thin"><color auto="1"/></left><right style="bogus"/><bottom style="double"/>'
    '<diagonal style="dashDot"><color indexed="10"/></diagonal></border></borders>'
    '<cellStyleXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" '
    'applyFont="0"><alignment horizontal="center" vertical="top" wrapText="1" '
    'textRotation="255" indent="2"/><protection locked="0" hidden="1"/></xf>'
    '<xf numFmtId="14" fontId="7" fillId="99" borderId="1" xfId="5"/></cellXfs>'
    '<cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/>'
    '<cellStyle name="Mine" xfId="3"/></cellStyles>'
    '<dxfs><dxf><font><b/><sz val="0.5"/><color rgb="FF9C0006"/></font>'
    '<numFmt numFmtId="165" formatCode="0.0"/><fill><patternFill><bgColor rgb="FFFFC7CE"/>'
    "</patternFill></fill></dxf></dxfs>"
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>'
    '<colors><indexedColors><rgbColor rgb="FF000000"/></indexedColors></colors>'
    "<extLst/>"
)


def test_every_table_with_its_indices_resolved():
    styles = with_styles(STYLES).styles
    assert styles.number_formats == {164: "0.000%"}
    plain, fancy = styles.fonts
    assert (plain.name, plain.size, plain.color, plain.family, plain.scheme) == (
        "Calibri", 11.0, xlio.Color(theme=1), 2, "minor",
    )  # fmt: skip
    assert (fancy.bold, fancy.italic, fancy.underline, fancy.strike, fancy.vert_align) == (
        True, False, "single", True, "superscript",
    )  # fmt: skip
    none, gray, solid, gradient, bare = styles.fills
    assert none == xlio.PatternFill("none")
    assert gray.pattern == "gray125"
    assert solid == xlio.PatternFill("solid", xlio.Color(rgb="FF00B0F0"), xlio.Color(indexed=64))
    assert gradient.degree == 90.0
    assert gradient.stops == (
        xlio.GradientStop(0.0, xlio.Color(theme=0)),
        xlio.GradientStop(1.0, xlio.Color(theme=4)),
    )
    assert bare.pattern is None  # no patternType written
    empty, lined = styles.borders
    assert empty.left == xlio.BorderSide()
    assert lined.left == xlio.BorderSide("thin", xlio.Color(auto=True))
    assert lined.right == xlio.BorderSide()  # a style that is not one: none
    assert (lined.top, lined.bottom.style, lined.diagonal.style) == (None, "double", "dashDot")
    assert (lined.diagonal_up, lined.diagonal_down) == (True, False)

    base, styled, broken = styles.cell_formats
    assert base.style is styles.cell_style_formats[0]
    assert (styled.number_format_id, styled.number_format) == (164, "0.000%")
    assert (styled.font, styled.fill, styled.border) == (fancy, gradient, lined)
    assert styled.alignment == xlio.Alignment(
        "center", "top", wrap_text=True, indent=2, text_rotation=255
    )
    assert styled.protection == xlio.Protection(locked=False, hidden=True)
    assert (styled.apply_number_format, styled.apply_font, styled.apply_fill) == (True, False, None)
    # Indices that lead nowhere are None; the numbers stay.
    assert (broken.font, broken.fill, broken.style) == (None, None, None)
    assert (broken.font_id, broken.fill_id, broken.xf_id, broken.number_format) == (7, 99, 5, None)
    assert styles.cell_format(1) is styled
    assert styles.cell_format(3) is None

    normal, mine = styles.cell_styles
    assert (normal.name, normal.builtin_id, normal.format) == ("Normal", 0, base.style)
    assert (mine.builtin_id, mine.format) == (None, None)
    (dxf,) = styles.differential_formats
    assert (dxf.font.bold, dxf.font.size, dxf.font.color) == (True, 0.5, xlio.Color("FF9C0006"))
    assert (dxf.number_format_id, dxf.number_format) == (165, "0.0")
    assert dxf.fill == xlio.PatternFill(None, None, xlio.Color(rgb="FFFFC7CE"))
    assert styles.indexed_colors == ("FF000000",)
    assert [u.path for u in styles.unparsed] == ["styleSheet/extLst"]
    assert styles.table_styles.default_table_style == "TableStyleMedium2"


def test_a_cells_style_leads_to_its_format():
    wb = with_styles(STYLES, '<row r="1"><c r="A1" s="1"><v>0.5</v></c></row>')
    (cell,) = wb["Data"].cells()
    assert wb.styles.cell_format(cell.style).number_format == "0.000%"


# Rich text ------------------------------------------------------------------------------------

STRINGS = [
    "<si><t>plain</t></si>",
    '<si><r><t xml:space="preserve">bo </t></r><r><rPr><b/><sz val="14"/><color theme="4"/>'
    '<rFont val="Arial"/><vertAlign val="subscript"/></rPr><t>ld_x000D_</t></r></si>',
    '<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh>'
    '<phoneticPr fontId="1" type="Hiragana" alignment="center"/></si>',
    "<si><t>  outer  </t></si>",
]


def test_rich_strings_and_their_plain_text_agree():
    wb = with_styles("<fonts><font/><font/></fonts>", strings=STRINGS)
    rich = wb.rich_strings
    assert (
        [r.text for r in rich] == list(wb.shared_strings) == ["plain", "bo ld\r", "東京", "outer"]
    )
    assert rich[0].runs == ()
    first, second = rich[1].runs
    assert (first.text, first.font) == ("bo ", None)
    font = second.font
    assert (second.text, font.name, font.bold, font.size, font.vert_align) == (
        "ld\r", "Arial", True, 14.0, "subscript",
    )  # fmt: skip
    assert font.color == xlio.Color(theme=4)
    assert rich[2].phonetic == (xlio.PhoneticRun("トウキョウ", 0, 2),)
    assert rich[2].phonetic_properties == xlio.PhoneticProperties(1, "Hiragana", "center")


@pytest.mark.parametrize(
    "border",
    ['<left style="thin"/><start style="thick"/>', '<right style="thin"/><end style="thick"/>'],
)
def test_a_left_before_start_is_refused(border):
    # Excel refuses a border that writes left before start, or right before end (capture edges).
    parts = _parts(f"<borders><border/><border>{border}</border></borders>")
    with pytest.raises(xlio.PackageError, match="before"):
        xlio.check(io.BytesIO(package(extra_parts=parts)))
    with pytest.raises(xlio.PackageError, match="before"):
        xlio.open(io.BytesIO(package(extra_parts=parts))).styles  # noqa: B018


def test_a_t_before_runs_is_a_run_of_its_own():
    # Excel keeps a <t> written before runs as a run without properties, first (capture edges).
    strings = ['<si><t xml:space="preserve">HH </t><r><rPr><b/></rPr><t>HH</t></r></si>']
    wb = with_styles("<fonts><font/></fonts>", strings=strings)
    (first,) = wb.rich_strings
    assert [(r.text, r.font is None) for r in first.runs] == [("HH ", True), ("HH", False)]
    assert first.runs[1].font is not None
    assert first.runs[1].font.bold
    assert [r.text for r in wb.rich_strings] == list(wb.shared_strings) == ["HH HH"]


@pytest.mark.parametrize("text", ["<r><t>a</t></r><t>b</t>", "<t>a</t><r><t>b</t></r><t>c</t>"])
def test_a_t_after_a_run_is_refused(text):
    # Excel refuses a string with a <t> after a run (capture edges), shared or inline.
    shared = package('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', strings=[f"<si>{text}</si>"])
    with pytest.raises(xlio.PackageError, match="after its runs"):
        list(xlio.open(io.BytesIO(shared))["Data"].cells())
    with pytest.raises(xlio.PackageError, match="after its runs"):
        xlio.check(io.BytesIO(shared))
    rows = f'<row r="1"><c r="A1" t="inlineStr"><is>{text}</is></c></row>'
    with pytest.raises(xlio.PackageError, match="after its runs"):
        list(xlio.open(io.BytesIO(package(rows)))["Data"].cells())


def test_cells_with_their_runs():
    rows = (
        '<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s" ph="1"><v>2</v></c>'
        '<c r="C1" t="inlineStr"><is><r><t>in</t></r><r><rPr><i/></rPr>'
        '<t xml:space="preserve"> line</t></r></is></c><c r="D1"><v>1</v></c></row>'
    )
    wb = with_styles("", rows, STRINGS)
    plain = list(wb["Data"].cells())
    assert all(c.rich is None for c in plain)
    a, b, c, d = wb["Data"].cells(rich=True)
    assert a.rich is wb.rich_strings[1]
    assert (b.phonetic, b.rich.phonetic[0].text) == (True, "トウキョウ")
    assert (c.value, c.rich.text, [r.text for r in c.rich.runs]) == ("in line",) * 2 + (
        ["in", " line"],
    )
    assert c.rich.runs[1].font.italic
    assert d.rich is None


def test_prefixed_inline_runs():
    sheet = (
        f'<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr">'
        '<x:is><x:r><x:rPr><x:b/><x:sz val="9"/></x:rPr><x:t>b</x:t></x:r></x:is></x:c></x:row>'
        "</x:sheetData></x:worksheet>"
    )
    wb = xlio.open(io.BytesIO(package(sheet_xml=sheet)))
    (cell,) = wb["Data"].cells(rich=True)
    (run,) = cell.rich.runs
    assert (run.text, run.font.bold, run.font.size) == ("b", True, 9.0)


# Files from Excel and other writers -----------------------------------------------------------


@needs_openpyxl
@pytest.mark.parametrize(
    "name", ["marker_typed.xlsx", "iterate_cached.xlsx", "excel_formulas.xlsx"]
)
def test_styles_excel_saved_agree_with_openpyxl(name):
    import openpyxl

    ws = openpyxl.load_workbook(FILES / name).worksheets[0]
    with xlio.open(FILES / name) as wb:
        styles = wb.styles
        assert styles.cell_styles[0].name == "Normal"
        for cell in wb.sheets[0].cells():
            fmt = styles.cell_format(cell.style)
            theirs = ws.cell(cell.row, cell.column)
            assert (fmt.font.name, fmt.font.size, fmt.font.bold) == (
                theirs.font.name, theirs.font.size, bool(theirs.font.bold),
            )  # fmt: skip
            assert (fmt.fill.pattern or "none") == (theirs.fill.patternType or "none")
            if fmt.number_format is not None:
                assert fmt.number_format == theirs.number_format


def test_what_xlsxwriter_writes(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "x.xlsx"
    book = xlsxwriter.Workbook(path)
    ws = book.add_worksheet("Data")
    fancy = book.add_format(
        {"bold": True, "italic": True, "underline": 2, "font_color": "#FF0000", "font_size": 14,
         "bg_color": "#FFFF00", "pattern": 1, "border": 5, "align": "center", "valign": "top",
         "text_wrap": True, "rotation": 45, "indent": 1, "num_format": "0.000%"}
    )  # fmt: skip
    ws.write("A1", 0.5, fancy)
    sup = book.add_format({"font_script": 1, "font_name": "Courier New"})
    ws.write_rich_string("A2", "E = mc", sup, "2")
    ws.write("A3", 1)
    ws.conditional_format(
        "A3", {"type": "cell", "criteria": ">", "value": 0, "format": book.add_format({"bold": 1})}
    )
    book.close()
    with xlio.open(path) as wb:
        cells = {(c.row, c.column): c for c in wb["Data"].cells(rich=True)}
        fmt = wb.styles.cell_format(cells[1, 1].style)
        font = fmt.font
        assert (font.bold, font.italic, font.underline, font.size) == (True, True, "double", 14.0)
        assert font.color.rgb == "FFFF0000"
        assert fmt.fill == xlio.PatternFill(
            "solid", xlio.Color(rgb="FFFFFF00"), xlio.Color(indexed=64)
        )
        assert fmt.border.left.style == "thick"
        assert fmt.alignment == xlio.Alignment(
            "center", "top", wrap_text=True, indent=1, text_rotation=45
        )
        assert fmt.number_format == "0.000%"
        rich = cells[2, 1].rich
        assert [r.text for r in rich.runs] == ["E = mc", "2"]
        assert (rich.runs[1].font.vert_align, rich.runs[1].font.name) == (
            "superscript",
            "Courier New",
        )
        (dxf,) = wb.styles.differential_formats
        assert dxf.font.bold


def test_what_openpyxl_writes(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Border, Font, GradientFill, NamedStyle, Protection, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "x"
    ws["A1"].fill = GradientFill(type="path", left=0.2, stop=("FF0000", "0000FF"))
    ws["A1"].border = Border(diagonal=Side(style="thin"), diagonalDown=True)
    ws["A1"].protection = Protection(locked=False)
    ws["A1"].font = Font(name="Georgia", size=-10)  # a negative size: Excel takes 10
    highlight = NamedStyle(name="highlight", font=Font(bold=True))
    wb.add_named_style(highlight)
    ws["A2"] = "y"
    ws["A2"].style = "highlight"
    path = tmp_path / "o.xlsx"
    wb.save(path)
    with xlio.open(path) as book:
        cells = {(c.row, c.column): c for c in book["Data"].cells()}
        a1 = book.styles.cell_format(cells[1, 1].style)
        assert (a1.fill.type, a1.fill.left) == ("path", 0.2)
        assert [s.color.rgb for s in a1.fill.stops] == ["00FF0000", "000000FF"]
        assert (a1.border.diagonal.style, a1.border.diagonal_down) == ("thin", True)
        assert a1.protection.locked is False
        assert (a1.font.name, a1.font.size, a1.font.raw_size) == ("Georgia", 10.0, "-10")
        a2 = book.styles.cell_format(cells[2, 1].style)
        assert a2.style.font.bold
        assert "highlight" in [s.name for s in book.styles.cell_styles]
