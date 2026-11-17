"""The public reading layer: ``open``, ``Workbook``, ``Sheet``, ``Parts``, ``check``, limits."""

from __future__ import annotations

import dataclasses
import io
import re
import zipfile
from pathlib import Path

import pytest

import xlio
from tests.test_read import DOC, MAIN, PKG, REL, package

FILES = Path(__file__).with_name("files")
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"


def rels(*entries: tuple[str, str, str]) -> str:
    body = "".join(
        f'<Relationship Id="{i}" Type="{t}" Target="{target}"/>' for i, t, target in entries
    )
    return f'<Relationships xmlns="{PKG}">{body}</Relationships>'


def opened(data: bytes, limits: xlio.ReadLimits | None = None) -> xlio.Workbook:
    return xlio.open(io.BytesIO(data), limits=limits)


# Opening ---------------------------------------------------------------------------------------


def test_open_a_path_and_a_stream(tmp_path):
    data = package('<row r="1"><c r="A1"><v>1</v></c></row>')
    path = tmp_path / "book.xlsx"
    path.write_bytes(data)
    with xlio.open(path) as wb:
        assert [c.value for c in wb["Data"].cells()] == [1.0]
    stream = io.BytesIO(data)
    with xlio.open(stream) as wb:
        assert len(wb) == 1
    assert not stream.closed  # a stream given to open is the caller's


def test_an_ole_compound_file_is_refused(tmp_path):
    head = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 600
    with pytest.raises(xlio.PackageError, match=r"\.xls workbook, or an xlsx file protected"):
        xlio.open(io.BytesIO(head))
    path = tmp_path / "old.xls"
    path.write_bytes(head)
    with pytest.raises(xlio.PackageError, match="OLE compound file"):
        xlio.open(path)


def test_a_binary_workbook_is_refused():
    data = package(extra_parts={"xl/workbook.bin": b"\0"})
    with pytest.raises(xlio.PackageError, match=r"\.xlsb"):
        opened(data)


def test_stated_sizes_are_checked_before_anything_is_read():
    data = package(extra_parts={"xl/media/big.bin": b"\0" * 5000})
    with pytest.raises(xlio.LimitError, match=r"big\.bin: 5000 bytes"):
        opened(data, xlio.ReadLimits(max_part_bytes=4000))
    with pytest.raises(xlio.LimitError, match="in all"):
        opened(data, xlio.ReadLimits(max_total_bytes=5000))


def test_a_part_read_again_counts_once_toward_the_total():
    data = package('<row r="1"><c r="A1"><v>1</v></c></row>')
    total = sum(i.file_size for i in zipfile.ZipFile(io.BytesIO(data)).infolist())
    wb = opened(data, xlio.ReadLimits(max_total_bytes=total))
    sheet = wb["Data"]
    for _ in range(3):  # the sheet's part, read again each time
        assert [c.value for c in sheet.cells()] == [1.0]
    sheet.merges  # noqa: B018 - and once more for the structure
    xlio.check(io.BytesIO(data), limits=xlio.ReadLimits(max_total_bytes=total))


# Sheets ----------------------------------------------------------------------------------------


SHEETS = (
    '<sheet name="Data" sheetId="1" r:id="rId1"/>'
    '<sheet name="Chart" sheetId="2" r:id="rId2" state="hidden"/>'
    '<sheet name="Dialog" sheetId="3" r:id="rId3" state="veryHidden"/>'
    '<sheet name="Macro" sheetId="4" r:id="rId4" state="odd"/>'
    '<sheet name="Other" sheetId="5" r:id="rId5"/>'
)
SHEET_RELS = rels(
    ("rId1", f"{DOC}/worksheet", "worksheets/sheet1.xml"),
    ("rId2", f"{DOC}/chartsheet", "chartsheets/sheet1.xml"),
    ("rId3", f"{DOC}/dialogsheet", "dialogsheets/sheet1.xml"),
    ("rId4", "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet", "m.xml"),
    ("rId5", "urn:example:sheet", "o.xml"),
)


def test_every_sheet_with_its_kind_and_state():
    wb = opened(package(sheets=SHEETS, extra_parts={WORKBOOK_RELS: SHEET_RELS}))
    assert [(s.name, s.index, s.kind, s.state) for s in wb.sheets] == [
        ("Data", 0, "worksheet", "visible"),
        ("Chart", 1, "chartsheet", "hidden"),
        ("Dialog", 2, "dialogsheet", "veryHidden"),
        ("Macro", 3, "macrosheet", "visible"),
        ("Other", 4, "other", "visible"),
    ]
    assert wb["chart"] is wb.sheets[1] is wb.sheet(1)  # names without case, as Excel
    assert list(wb) == list(wb.sheets)
    with pytest.raises(KeyError):
        wb["Nope"]
    with pytest.raises(IndexError):
        wb[9]
    with pytest.raises(ValueError, match="chartsheet, not a worksheet"):
        list(wb["Chart"].cells())


MACRO_RELS = rels(
    ("rId1", f"{DOC}/worksheet", "worksheets/sheet1.xml"),
    ("rId2", "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet", "m.xml"),
    ("rId3", "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet", "i.xml"),
)
MACRO_SHEETS = (
    '<sheet name="Data" sheetId="1" r:id="rId1"/>'
    '<sheet name="Macro1" sheetId="2" r:id="rId2"/>'
    '<sheet name="Intl" sheetId="3" r:id="rId3"/>'
)


def macro_sheet(rows: str) -> str:
    return (
        f'<xm:macrosheet xmlns="{MAIN}" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        f"<sheetData>{rows}</sheetData></xm:macrosheet>"
    )


def test_a_macro_sheet_has_cells_as_a_worksheet_does():
    rows = (
        '<row r="1"><c r="A1"><v>1</v></c></row>'
        '<row r="2"><c r="A2" t="str"><f>A1&amp;""</f><v>1</v></c></row>'
    )
    parts = {WORKBOOK_RELS: MACRO_RELS, "xl/m.xml": macro_sheet(rows), "xl/i.xml": macro_sheet("")}
    data = package(sheets=MACRO_SHEETS, extra_parts=parts)
    with opened(data) as wb:
        assert [s.kind for s in wb.sheets] == ["worksheet", "macrosheet", "macrosheet"]
        assert [
            (c.row, c.column, c.value, c.formula and c.formula.text) for c in wb["Macro1"].cells()
        ] == [
            (1, 1, 1.0, None),
            (2, 1, "1", 'A1&""'),
        ]
        assert list(wb["Intl"].cells()) == []
    xlio.check(io.BytesIO(data))
    off_grid = '<row r="1048577"><c r="A1048577"><v>1</v></c></row>'
    parts["xl/m.xml"] = macro_sheet(off_grid)
    with pytest.raises(xlio.PackageError, match="1048577"):  # check() reads macro sheets too
        xlio.check(io.BytesIO(package(sheets=MACRO_SHEETS, extra_parts=parts)))


@pytest.mark.parametrize(
    ("content_type", "kind"),
    [
        ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "workbook"),
        ("application/vnd.ms-excel.sheet.macroEnabled.main+xml", "macro-enabled"),
        (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "template",
        ),
        ("application/vnd.ms-excel.template.macroEnabled.main+xml", "macro-enabled template"),
        ("text/plain", "workbook"),
    ],
)
def test_the_kind_of_workbook_comes_from_its_content_type(content_type, kind):
    types = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f'<Override PartName="/XL/Workbook.xml" ContentType="{content_type}"/></Types>'
    )
    workbook_rels = rels(
        ("rId1", f"{DOC}/worksheet", "worksheets/sheet1.xml"),
        ("rId7", "http://schemas.microsoft.com/office/2006/relationships/vbaProject", "v.bin"),
    )
    wb = opened(package(extra_parts={"[Content_Types].xml": types, WORKBOOK_RELS: workbook_rels}))
    assert wb.kind == kind
    assert wb.has_vba
    assert not opened(package()).has_vba


def test_names_calc_and_dates():
    sheets = (
        '<sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames>'
        '<definedName name="Rate">Data!$A$1</definedName></definedNames>'
        '<calcPr calcId="191029" iterate="1"/><sheets>'
    )
    wb = opened(package(workbook_extra='<workbookPr date1904="true"/>', sheets=sheets))
    assert wb.date1904
    assert wb.defined_names == (xlio.DefinedName("Rate", "Data!$A$1", None, False),)
    assert (wb.calc.calc_id, wb.calc.iterate) == (191029, True)


# Cells -----------------------------------------------------------------------------------------


def test_cells_carry_their_style_and_metadata_indexes():
    rows = (
        '<row r="1"><c r="A1" s="3"><v>1</v></c><c r="B1" s="nan"><v>2</v></c>'
        '<c r="C1" cm="1" vm="1"><f>1</f><v>1</v></c><c r="D1" s="-1" cm="x"><v>4</v></c></row>'
    )
    got = [(c.style, c.cm, c.vm) for c in opened(rich_error_package(rows))["Data"].cells()]
    assert got == [(3, None, None), (0, None, None), (0, 1, 1), (0, None, None)]


def test_errors_are_error_codes():
    rows = (
        '<row r="1"><c r="A1" t="e"><v>#N/A</v></c><c r="B1" t="e"><v>#DIV/0!</v></c>'
        '<c r="C1" t="e"><v>#BUSY!</v></c><c r="D1"><v>1e400</v></c></row>'
    )
    values = [c.value for c in opened(package(rows))["Data"].cells()]
    assert values == [
        xlio.ErrorCode.NA,
        xlio.ErrorCode.DIV0,
        xlio.ErrorCode.BUSY,
        xlio.ErrorCode.NULL,
    ]
    assert all(type(v) is xlio.ErrorCode for v in values)
    assert values[0] == "#N/A"  # an ErrorCode is a str


@pytest.mark.parametrize("code", ["#UNHEARD!", "#SPILL!", "#field!"])
def test_a_plain_error_excel_refuses_is_refused(code):
    rows = f'<row r="1"><c r="A1" t="e"><v>{code}</v></c></row>'
    with pytest.raises(xlio.PackageError, match="an error Excel refuses"):
        list(opened(package(rows))["Data"].cells())


def test_a_subnormal_reads_as_zero():
    rows = '<row r="1"><c r="A1"><v>4.9e-324</v></c><c r="B1"><v>-2.2e-308</v></c></row>'
    assert [c.value for c in opened(package(rows))["Data"].cells()] == [0.0, 0.0]


def rich_error_package(rows: str) -> bytes:
    """A package with Excel's own metadata and rich-value parts (saved by Excel 16 for Mac)."""
    workbook_rels = rels(
        ("rId1", f"{DOC}/worksheet", "worksheets/sheet1.xml"),
        ("rId3", f"{DOC}/sheetMetadata", "metadata.xml"),
        ("rId4", "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
         "richData/rdrichvalue.xml"),
        ("rId5", "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
         "richData/rdrichvaluestructure.xml"),
    )  # fmt: skip
    parts = {
        WORKBOOK_RELS: workbook_rels,
        "xl/metadata.xml": (FILES / "metadata.xml").read_bytes(),
        "xl/richData/rdrichvalue.xml": (FILES / "rdrichvalue.xml").read_bytes(),
        "xl/richData/rdrichvaluestructure.xml": (FILES / "rdrichvaluestructure.xml").read_bytes(),
    }
    return package(rows, extra_parts=parts)


def test_a_rich_value_error_reads_as_excel_shows_it():
    # metadata.xml: value metadata 1 is rich value 0, an _error of errorType 13 (#CALC!).
    rows = (
        '<row r="1"><c r="A1" t="e" vm="1"><f>FILTER(B1:B2,B1:B2&gt;5)</f><v>#VALUE!</v></c>'
        '<c r="B1" t="e"><v>#VALUE!</v></c><c r="C1" cm="1"><f>1</f><v>1</v></c></row>'
    )
    wb = opened(rich_error_package(rows))
    assert wb.cell_metadata == xlio.CellMetadata(frozenset({1}), {1: xlio.ErrorCode.CALC})
    values = [c.value for c in wb["Data"].cells()]
    assert values == [xlio.ErrorCode.CALC, xlio.ErrorCode.VALUE, 1.0]
    (error,) = wb.rich_values.values()  # the same rich value, raw
    assert (type(error), error.type, error.values) == (
        xlio.RichValue, "_error", {"errorType": "13", "subType": "3"},
    )  # fmt: skip


def test_an_error_rich_value_wins_over_any_cached_value():
    # Excel shows the error a vm points to over a cached 5 or TRUE when it does not recalculate.
    rows = (
        '<row r="1"><c r="A1" t="e" vm="1"><v>#DIV/0!</v></c><c r="B1" vm="1"><v>5</v></c>'
        '<c r="C1" t="b" vm="1"><v>1</v></c><c r="D1" t="str" vm="1"><f>"x"</f><v>x</v></c>'
        '<c r="E1"><v>5</v></c></row>'
    )
    cells = list(opened(rich_error_package(rows))["Data"].cells())
    assert [(c.value, c.type) for c in cells] == [(xlio.ErrorCode.CALC, "e")] * 4 + [(5.0, "n")]


def test_a_new_value_drops_the_stale_rich_value(tmp_path):
    # Excel shows and saves the rich value a cell's vm points to, not the cached value beside it,
    # when the file is not recalculated on opening; so a written value drops vm, even when it is
    # the #VALUE! the cell already caches. Given the error the rich value stands for (#CALC!),
    # the cell is left as it is, vm and all: Excel's record of it stays right.
    rows = (
        '<row r="1"><c r="A1" t="e" vm="1"><f>FILTER(D1:D2,D1:D2&gt;5)</f><v>#VALUE!</v></c>'
        '<c r="B1" t="e" vm="1"><f>FILTER(D1:D2,D1:D2&gt;5)</f><v>#VALUE!</v></c>'
        '<c r="C1" t="e" vm="1"><v>#VALUE!</v></c>'
        '<c r="E1" t="e" vm="1"><f>FILTER(D1:D2,D1:D2&gt;5)</f><v>#VALUE!</v></c>'
        '<c r="F1" t="e" vm="1"><f>FILTER(D1:D2,D1:D2&gt;5)</f><v>#VALUE!</v></c></row>'
    )
    source = tmp_path / "in.xlsx"
    source.write_bytes(rich_error_package(rows))
    target = tmp_path / "out.xlsx"
    new = {(1, 1): xlio.ErrorCode.VALUE, (1, 2): 5.0, (1, 6): xlio.ErrorCode.CALC}
    xlio.patch(
        source,
        target,
        cached={"Data": lambda r, c: new.get((r, c))},
        cells={"Data": {(1, 3): xlio.CellEdit(None, True)}},
        full_calc_on_load=False,
    )
    sheet_xml = zipfile.ZipFile(target).read("xl/worksheets/sheet1.xml").decode()
    cells = dict(re.findall(r'<c r="([A-Z]+1)"([^>]*)>', sheet_xml))
    kept = ("E1", "F1")
    assert ["vm=" in cells[ref] for ref in ("A1", "B1", "C1", *kept)] == [False] * 3 + [True] * 2
    assert rows[rows.index('<c r="F1"') :] in sheet_xml  # untouched
    values = {(c.row, c.column): c.value for c in opened(target.read_bytes())["Data"].cells()}
    assert values == {
        (1, 1): xlio.ErrorCode.VALUE, (1, 2): 5.0, (1, 3): True, (1, 5): xlio.ErrorCode.CALC,
        (1, 6): xlio.ErrorCode.CALC,
    }  # fmt: skip


@pytest.mark.parametrize("attribute", ['vm="2"', 'cm="2"'])
def test_a_metadata_index_past_its_table_is_refused(attribute):
    # Excel refuses the file (capture attributes: cm="5" and vm="5" with no metadata part).
    rows = f'<row r="1"><c r="A1" {attribute}><v>1</v></c></row>'
    with pytest.raises(xlio.PackageError, match="past its table"):
        list(opened(rich_error_package(rows))["Data"].cells())


ROWS = "".join(
    f'<row r="{r}">' + "".join(f'<c r="{c}{r}"><v>{r}</v></c>' for c in "ABCD") + "</row>"
    for r in range(1, 6)
)


def test_a_range_bounds_what_is_read():
    sheet = opened(package(ROWS))["Data"]
    got = [(c.row, c.column) for c in sheet.cells("C4:b2")]
    assert got == [(2, 2), (2, 3), (3, 2), (3, 3), (4, 2), (4, 3)]
    assert [(c.row, c.column) for c in sheet.cells("$D$5")] == [(5, 4)]
    for ref in ("A0", "XFE1", "A1:B2:C3", "1:2", ""):
        with pytest.raises(ValueError, match=r"not a cell|off the sheet"):
            list(sheet.cells(ref))


def test_max_cells_bounds_the_range_asked_for_and_the_range_read():
    limits = xlio.ReadLimits(max_cells=12)
    sheet = opened(package(ROWS), limits)["Data"]
    assert len(list(sheet.cells("A1:C4"))) == 12
    with pytest.raises(xlio.LimitError, match="A1:D4 covers more than the limit of 12"):
        sheet.cells("A1:D4").__next__()
    with pytest.raises(xlio.LimitError, match="more than the limit of 12 cells"):
        list(sheet.cells())
    assert len(list(opened(package(ROWS), xlio.ReadLimits(max_cells=20))["Data"].cells())) == 20


def test_the_default_limits_read_a_cell_in_the_last_position():
    rows = '<row r="1"><c r="A1"><v>1</v></c></row><row r="1048576"><c r="XFD1048576"/></row>'
    assert len(list(opened(package(rows))["Data"].cells())) == 2
    with pytest.raises(xlio.LimitError):
        list(opened(package(rows), xlio.ReadLimits.UNTRUSTED)["Data"].cells())


def test_the_untrusted_preset():
    u = xlio.ReadLimits.UNTRUSTED
    assert dataclasses.astuple(u) == (10_000, 256 << 20, 1 << 30, 1_048_576, 8 << 20, 32 << 20)
    assert dataclasses.replace(u, max_cells=5).max_parts == 10_000
    defaults = xlio.ReadLimits()
    assert (defaults.max_part_bytes, defaults.max_total_bytes) == (16 << 30, 64 << 30)
    assert defaults.max_cells == 1_048_576 * 16_384  # the whole sheet: never refuses a workbook


# Parts -----------------------------------------------------------------------------------------


def test_parts_content_types_and_relationships():
    types = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="XML" ContentType="application/xml"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
        'relationships+xml"/><Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    sheet_rels = rels(
        ("rId1", f"{DOC}/hyperlink", "https://example.com/a?b=1"),
        ("rId2", f"{DOC}/drawing", "../drawings/drawing1.xml"),
        ("rId3", f"{DOC}/table", "/xl/tables/table1.xml"),
    ).replace('Target="https://example.com/a?b=1"/>', 'Target="https://example.com/a?b=1" '
              'TargetMode="External"/>')  # fmt: skip
    wb = opened(package(sheet_rels=sheet_rels, extra_parts={"[Content_Types].xml": types}))
    parts = wb.parts
    assert next(iter(parts)) == "[Content_Types].xml"
    assert "xl/workbook.xml" in parts
    assert "xl/nope.xml" not in parts
    assert parts.content_type("xl/worksheets/sheet1.xml").endswith("worksheet+xml")
    assert parts.content_type("xl/workbook.xml") == "application/xml"  # the default, any case
    assert parts.content_type("xl/media/a.png") is None
    assert parts.size("xl/workbook.xml") == len(parts.read("xl/workbook.xml"))
    with parts.open("xl/workbook.xml") as f:
        assert f.read(9) == b"<workbook"
    assert parts.xml("xl/workbook.xml").tag == f"{{{MAIN}}}workbook"
    assert wb["Data"].relationships == (
        xlio.Relationship("rId1", f"{DOC}/hyperlink", "https://example.com/a?b=1", True, None),
        xlio.Relationship(
            "rId2", f"{DOC}/drawing", "../drawings/drawing1.xml", False, "xl/drawings/drawing1.xml"
        ),
        xlio.Relationship(
            "rId3", f"{DOC}/table", "/xl/tables/table1.xml", False, "xl/tables/table1.xml"
        ),
    )
    assert parts.relationships("")[0].target_part == "xl/workbook.xml"
    assert parts.relationships("xl/nope.xml") == ()
    with pytest.raises(KeyError):
        parts.read("xl/nope.xml")


def test_the_namespace_constants_are_the_ones_the_tests_build_with():
    constants = (xlio.ns.MAIN, xlio.ns.RELATIONSHIPS, xlio.ns.PACKAGE_RELATIONSHIPS)
    assert constants == (MAIN, REL, PKG)
    assert f"{DOC}/worksheet" == xlio.ns.WORKSHEET
    # Excel reads the x14 conditional formats only under this exact URI (407d, not 407D).
    assert "407d" in xlio.ns.EXT_CONDITIONAL_FORMATTINGS


# check -----------------------------------------------------------------------------------------


def test_check_reads_every_part_even_those_nothing_refers_to():
    good = package('<row r="1"><c r="A1"><v>1</v></c></row>')
    xlio.check(io.BytesIO(good))
    hidden = package(extra_parts={"customXml/item1.xml": '<!DOCTYPE x [<!ENTITY e "x">]><x/>'})
    with xlio.open(io.BytesIO(hidden)) as wb:  # open reads only what is asked for
        assert len(list(wb["Data"].cells())) == 0
    with pytest.raises(xlio.PackageError, match=r"customXml/item1\.xml: declares a DTD"):
        xlio.check(io.BytesIO(hidden))


# Files written by Excel and by other programs -------------------------------------------------


@pytest.mark.parametrize(
    "name", ["marker_typed.xlsx", "iterate_cached.xlsx", "excel_formulas.xlsx"]
)
def test_files_saved_by_excel(name):
    path = FILES / name
    with xlio.open(path) as wb:
        assert wb.kind == "workbook"
        assert not wb.has_vba
        assert wb.parts.content_type(wb.workbook_part).endswith("sheet.main+xml")
        sheet = wb.sheets[0]
        cells = list(sheet.cells())
    with zipfile.ZipFile(path) as z:
        xml = z.read(sheet.part).decode()
    assert len(cells) == xml.count("<c ")
    openpyxl = pytest.importorskip("openpyxl")
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    for cell in cells:
        theirs = ws.cell(cell.row, cell.column).value
        if cell.type == "n" and isinstance(theirs, int | float) and not isinstance(theirs, bool):
            # openpyxl turns numbers under a date format into dates; only plain numbers compare
            assert theirs == cell.value, (cell, theirs)
        elif cell.type in ("s", "str", "inlineStr", "b", "e"):
            # openpyxl reads an empty text result as None
            assert theirs == (cell.value or None), (cell, theirs)


def test_a_workbook_written_by_xlsxwriter(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "x.xlsx"
    book = xlsxwriter.Workbook(path)
    first = book.add_worksheet("First")
    first.write_row(0, 0, [1.5, "text", True])
    first.write_formula(1, 0, "=A1*2", None, 3)
    book.add_worksheet("Second").hide()
    book.define_name("Rate", "=First!$A$1")
    book.close()
    with xlio.open(path) as wb:
        assert [(s.name, s.state) for s in wb.sheets] == [
            ("First", "visible"),
            ("Second", "hidden"),
        ]
        cells = {(c.row, c.column): c for c in wb["First"].cells()}
        assert [cells[1, k].value for k in (1, 2, 3)] == [1.5, "text", True]
        assert (cells[2, 1].formula.text, cells[2, 1].value) == ("A1*2", 3.0)
        assert [(n.name, n.text) for n in wb.defined_names] == [("Rate", "First!$A$1")]
        assert wb.parts.content_type("xl/styles.xml").endswith("styles+xml")


def test_external_links():
    workbook_rels = rels(
        ("rId1", f"{DOC}/worksheet", "worksheets/sheet1.xml"),
        ("rId4", f"{DOC}/externalLink", "externalLinks/externalLink1.xml"),
        ("rId5", f"{DOC}/externalLink", "externalLinks/externalLink2.xml"),
    )
    wb = opened(package(extra_parts={WORKBOOK_RELS: workbook_rels}))
    assert wb.external_links == (
        "xl/externalLinks/externalLink1.xml",
        "xl/externalLinks/externalLink2.xml",
    )
    assert opened(package()).external_links == ()
