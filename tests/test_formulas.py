"""``Sheet.formulas()``: the formula cells only, the same as ``Sheet.cells()`` gives them.

Each case compares the stream with the formula cells of a full read of the same part, and with
the expat reader the scan falls back to, so both readers and the fallback are held to it.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

import xlio
from tests.test_read import DOC, MAIN, PKG, REL, package
from tests.watchdog import bounded
from xlio import _formulas
from xlio._formulas import NotPlainError, read_formulas, scan_formulas

FILES = Path(__file__).with_name("files")


def from_cells(sheet: xlio.Sheet) -> list[tuple[object, ...]]:
    return [
        (c.row, c.column, c.formula, c.cm, c.value is not None)
        for c in sheet.cells()
        if c.formula is not None
    ]


def agree(data: bytes, *, plain: bool | None = None) -> list[xlio.FormulaCell]:
    """The sheet's formulas, checked against a full read and against the expat reader."""
    sheet = xlio.open(io.BytesIO(data)).sheets[0]
    got = list(sheet.formulas())
    assert [tuple(f) for f in got] == from_cells(sheet)
    part = sheet._workbook._book._package
    with part.part(sheet.part) as f:
        assert list(read_formulas(f)) == got
    if plain is not None:
        with part.part(sheet.part) as f:
            try:
                list(scan_formulas(f))
                scanned = True
            except NotPlainError:
                scanned = False
        assert scanned == plain
    return got


def sheet(rows: str, head: str = "", tail: str = "", root: str = "") -> bytes:
    xml = (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"{root}>{head}'
        f"<sheetData>{rows}</sheetData>{tail}</worksheet>"
    )
    return package(sheet_xml=xml, extra_parts=METADATA)


ROWS = (
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c>'
    '<c r="C1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c></row>'
    '<row r="2"><c r="A2"><f t="shared" ref="A2:A4" si="0">B1+1</f><v>3</v></c>'
    '<c r="B2" t="str"><f>""</f><v></v></c><c r="C2"><f ca="1">NOW()</f><v /></c></row>'
    '<row r="3"><c r="A3"><f t="shared" si="0"/><v>4</v></c>'
    '<c r="B3" cm="1"><f t="array" ref="B3:B4" aca="false">SEQUENCE(2)</f><v>1</v></c></row>'
)
METADATA = {
    "xl/metadata.xml": f'<metadata xmlns="{MAIN}"><metadataTypes count="1"><metadataType '
    'name="XLDAPR"/></metadataTypes><cellMetadata count="1"><bk><rc t="1" v="0"/></bk>'
    "</cellMetadata></metadata>",
    "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">'
    f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
    f'<Relationship Id="rId3" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
    "</Relationships>",
}


def test_the_formula_cells_of_a_plain_sheet():
    got = agree(package(ROWS, extra_parts=METADATA), plain=True)
    assert [(f.row, f.column) for f in got] == [
        (1, 2),
        (1, 3),
        (2, 1),
        (2, 2),
        (2, 3),
        (3, 1),
        (3, 2),
    ]
    shared, child = got[2], got[5]
    assert (shared.formula.kind, shared.formula.ref, shared.formula.si) == ("shared", "A2:A4", 0)
    assert (child.formula.text, child.formula.si) == (None, 0)  # given as cells() gives it
    assert got[1].formula.text == '"a"&"b"'
    assert [f.cached for f in got] == [True, True, True, True, False, True, True]
    assert got[4].formula.attributes == {"ca": "1"}
    assert got[6].cm == 1
    assert got[6].formula.attributes == {"t": "array", "ref": "B3:B4", "aca": "false"}


def test_each_way_a_cell_is_written_reads_alike():
    # Excel's attribute order, another order, single quotes and spaces, and an inline string:
    # the scan's three ways to a cell agree with each other and with cells().
    rows = (
        '<row r="1"><c r="A1" s="0" t="str" cm="1"><f>"x"</f><v>x</v></c>'
        '<c t="str" r="B1" cm="1"><f>"x"</f><v>x</v></c>'
        "<c  r = 'C1'  t = 'str' ><f  t = 'normal' >\"x\"</f><v>x</v></c>"
        '<c r="D1" t="inlineStr"><f>"x"</f><is><t>x</t></is></c></row>'
        # openpyxl's empty value, with a space: uncached, or the empty text
        '<row r="2"><c r="A2"><f>1</f><v /></c><c r="B2" t="str"><f>""</f><v    /></c>'
        '<c t="n" r="C2"><f>1</f><v /></c><c r="D2"><f>1</f><v    /></c></row>'
    )
    got = agree(package(rows, extra_parts=METADATA), plain=True)
    assert [(f.formula.text, f.cached) for f in got][:4] == [('"x"', True)] * 4
    assert [f.cached for f in got][4:] == [False, True, False, False]
    assert [f.cm for f in got] == [1, 1] + [None] * 6


def test_openpyxl_cells_take_the_quick_way():
    # openpyxl writes an empty value as <v />: its formula cells match the quick patterns too.
    for cell in (b'<c r="B1"><f>A1*2</f><v /></c>', b'<c r="B1"><f>A1*2</f><v/></c>'):
        assert _formulas._EXCEL.match(cell) is not None
        assert _formulas._USUAL.match(cell) is not None


def test_a_value_left_empty_is_cached_only_as_text():
    rows = (
        '<row r="1"><c r="A1"><f>1</f><v/></c><c r="B1" t="str"><f>""</f><v/></c>'
        '<c r="C1" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
        '<c r="D1" t="inlineStr"><f>1</f><is><t>x</t></is></c></row>'
    )
    assert [f.cached for f in agree(package(rows), plain=True)] == [False, True, True, True]


def test_references_line_ends_and_escapes_read_as_expat_reads_them():
    rows = (
        '<row r="1"><c r="A1"><f>"&lt;&gt;&amp;&quot;&apos;&#10;&#x41;"\r\n&amp;"_x0042_"\r</f></c>'
        "<c r='B1'><f\tt = 'shared'\nref=\"B1:B2\" si='0'>1</f></c></row>"
    )
    a1, b1 = agree(package(rows), plain=True)
    assert a1.formula.text == '"<>&"\'\nA"\n&"B"\n'
    assert (b1.formula.kind, b1.formula.si) == ("shared", 0)


@pytest.mark.parametrize(
    ("make", "why"),
    [
        (lambda: sheet(ROWS.replace('<row r="3">', '<!-- a note --><row r="3">')), "a comment"),
        (lambda: sheet(ROWS.replace("<v>3</v>", "<v><![CDATA[3]]></v>")), "a CDATA section"),
        (lambda: sheet(ROWS.replace('<row r="2">', '<?pi x?><row r="2">')), "an instruction"),
        (
            lambda: sheet(
                ROWS.replace('<c r="C1"', '<x:c xmlns:x="' + MAIN + '" r="C1"').replace(
                    "<v>ab</v></c>", "<v>ab</v></x:c>"
                )
            ),
            "a prefixed cell",
        ),
        (
            lambda: sheet(
                ROWS.replace('<f ca="1">', '<x:f xmlns:x="' + MAIN + '" ca="1">').replace(
                    "NOW()</f>", "NOW()</x:f>"
                )
            ),
            "a prefixed formula",
        ),
        (lambda: sheet(ROWS.replace('<c r="C2">', "<c>")), "a formula cell without r"),
        (lambda: sheet(ROWS.replace("<f>A1*2</f>", "<f>A1<x/>*2</f>")), "markup in a formula"),
        (lambda: sheet(ROWS.replace("<v>2</v>", "<v>2<x/></v>")), "markup in a value"),
        (
            lambda: sheet(ROWS.replace("<f>A1*2</f>", "<f>1</f><f>A1*2</f>")),
            "two formulas in a cell",
        ),
        (lambda: sheet(ROWS, tail='<c r="Z9"><f>1</f></c>'), "a cell after sheetData"),
        (lambda: sheet(ROWS.replace('<c r="C2">', '<c r="XFE2">')), "a column off the grid"),
    ],
)
def test_a_part_the_scan_does_not_take_is_read_by_expat_alike(make, why):
    agree(make(), plain=False)


def test_a_greater_than_sign_in_an_attribute():
    (got,) = [
        f
        for f in agree(sheet(ROWS.replace('<f ca="1">', '<f ca="1" x="&gt;>">')))
        if f.row == 2 and f.column == 3
    ]
    assert got.formula.attributes == {"ca": "1", "x": ">>"}


def test_a_formula_outside_sheet_data_is_not_a_formula_cell():
    tail = (
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings '
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:conditionalFormatting>'
        "<x14:cfRule><xm:f>A1&gt;0</xm:f></x14:cfRule></x14:conditionalFormatting>"
        "</x14:conditionalFormattings></ext></extLst>"
    )
    assert len(agree(sheet(ROWS, tail=tail), plain=True)) == 7


XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def test_prefixed_names_after_the_cells_are_scanned():
    # A control's anchor writes <xdr:row> and <xdr:col> after sheetData: not a cell.
    tail = (
        f'<!-- a note --><controls><anchor xmlns:xdr="{XDR}"><from><xdr:col>7</xdr:col>'
        "<xdr:row>5</xdr:row></from></anchor></controls>"
    )
    assert len(agree(sheet(ROWS, tail=tail), plain=True)) == 7
    prefixed = f'<x:c xmlns:x="{MAIN}" r="Z9"><x:f>1</x:f></x:c>'
    agree(sheet(ROWS, tail=prefixed), plain=False)


def test_other_encodings_are_read_by_expat():
    xml = (
        f'<?xml version="1.0" encoding="UTF-16"?><worksheet xmlns="{MAIN}">'
        f"<sheetData>{ROWS}</sheetData></worksheet>"
    )
    data = package(sheet_xml=xml.encode("utf-16"), extra_parts=METADATA)  # type: ignore[arg-type]
    got = agree(data, plain=False)
    assert len(got) == 7


@pytest.mark.parametrize("size", [1, 2, 3, 5, 7, 13, 64, 1000])
def test_every_chunk_boundary(monkeypatch, size):
    monkeypatch.setattr(_formulas, "_CHUNK", size)
    agree(package(ROWS, extra_parts=METADATA), plain=True)


@pytest.mark.parametrize("size", [5, 64])
def test_a_fallback_after_formulas_were_given_gives_each_once(monkeypatch, size):
    monkeypatch.setattr(_formulas, "_CHUNK", size)
    late = ROWS + '<!-- late --><row r="9"><c r="A9"><f>9</f></c></row>'
    got = agree(sheet(late), plain=False)
    assert [(f.row, f.column) for f in got][-2:] == [(3, 2), (9, 1)]
    assert len(got) == len({(f.row, f.column) for f in got})


def test_a_cell_longer_than_the_scan_keeps_is_read_by_expat(monkeypatch):
    monkeypatch.setattr(_formulas, "_CHUNK", 64)
    monkeypatch.setattr(_formulas, "_TAIL", 300)
    agree(sheet(ROWS), plain=True)  # every cell fits
    long = ROWS.replace("<v>ab</v>", "<v>" + "ab" * 200 + "</v>")
    agree(sheet(long), plain=False)


def test_a_long_sheet():
    rows = "".join(
        f'<row r="{r}"><c r="A{r}"><v>{r}</v></c><c r="B{r}"><f>A{r}*2</f><v>{2 * r}</v></c></row>'
        for r in range(1, 20_001)
    )
    assert len(agree(package(rows), plain=True)) == 20_000


@pytest.mark.parametrize(
    ("rows", "match"),
    [
        ('<row r="1"><c r="A1"><f t="shared" si="x"/></c></row>', "shared formula index"),
        ('<row r="1"><c r="A1" cm="1"><f>1</f></c></row>', "metadata index"),
        ('<row r="1"><c r="A1048577"><f>1</f></c></row>', "off the sheet"),
        ('<row r="1"><c r="1A"><f>1</f></c></row>', "cell reference"),
        ('<row r="1"><c r="A1"><f>1</f></row>', "mismatched tag"),
    ],
)
def test_what_cells_refuses_in_a_formula_cell_is_refused(rows, match):
    with pytest.raises(xlio.PackageError, match=match):
        list(xlio.open(io.BytesIO(package(rows))).sheets[0].formulas())


def test_formula_cells_past_the_cell_limit_are_refused():
    data = package('<row r="1"><c r="A1"><f>1</f></c></row><row r="5"><c r="E5"><f>2</f></c></row>')
    limits = xlio.ReadLimits(max_cells=24)
    with pytest.raises(xlio.LimitError):
        list(xlio.open(io.BytesIO(data), limits=limits).sheets[0].formulas())
    limits = xlio.ReadLimits(max_cells=25)
    assert len(list(xlio.open(io.BytesIO(data), limits=limits).sheets[0].formulas())) == 2


@bounded
def test_every_committed_workbook():
    found = 0
    for path in sorted(FILES.rglob("*.xls[xm]")):
        with xlio.open(path) as wb:
            for s in wb.sheets:
                if s.kind in ("worksheet", "macrosheet"):
                    got = list(s.formulas())
                    assert [tuple(f) for f in got] == from_cells(s), (path.name, s.name)
                    found += len(got)
    assert found > 0


def test_the_relationship_ids_the_sheet_names():
    rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    tail = (
        '<drawing r:id="rId5"/><legacyDrawing r:id="rId6"/><picture r:id="rId7"/>'
        '<tableParts count="2"><tablePart r:id="rId8"/><tablePart r:id="rId9"/></tableParts>'
    )
    xml = f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData/>{tail}</worksheet>'
    s = xlio.open(io.BytesIO(package(sheet_xml=xml, sheet_rels=rels))).sheets[0]
    assert (s.drawing_id, s.legacy_drawing_id, s.picture_id, s.table_ids) == (
        "rId5", "rId6", "rId7", ("rId8", "rId9"),
    )  # fmt: skip
    bare = xlio.open(io.BytesIO(package())).sheets[0]
    assert (bare.drawing_id, bare.legacy_drawing_id, bare.picture_id, bare.table_ids) == (
        None, None, None, (),
    )  # fmt: skip
