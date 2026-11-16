"""Drawings and their anchors, pictures, AlternateContent, VML, notes and threaded comments."""

from __future__ import annotations

import io
import struct
import time
import zlib

import pytest

import xlio
from tests.test_read import DOC, MAIN, PKG, REL, package
from tests.watchdog import bounded

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"


def rels(*entries: tuple[str, str, str], external: tuple[str, ...] = ()) -> str:
    body = "".join(
        f'<Relationship Id="{i}" Type="{t}" Target="{target}"'
        + (' TargetMode="External"' if i in external else "")
        + "/>"
        for i, t, target in entries
    )
    return f'<Relationships xmlns="{PKG}">{body}</Relationships>'


def sheet_with(sheet_body: str, sheet_rels: str, **parts: str | bytes) -> bytes:
    xml = f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData/>{sheet_body}</worksheet>'
    return package(sheet_xml=xml, sheet_rels=sheet_rels, extra_parts=parts)


def marker(col: int, col_off: int, row: int, row_off: int) -> str:
    return (
        f"<xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff>"
    )


def pic(id_: int, name: str, blip: str = '<a:blip r:embed="rIdImg"/>', extra: str = "") -> str:
    return (
        f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id_}" name="{name}" descr="alt {name}"{extra}/>'
        f"<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill>{blip}</xdr:blipFill>"
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr></xdr:pic>'
    )


def drawing_xml(*anchors: str) -> str:
    return (
        f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{REL}" xmlns:mc="{MC}" '
        f'xmlns:a14="{A14}" xmlns:asvg="{ASVG}" xmlns:u="urn:unknown">'
        + "".join(anchors)
        + "</xdr:wsDr>"
    )


def two_cell(frm: str, to: str, body: str, edit_as: str = "") -> str:
    attr = f' editAs="{edit_as}"' if edit_as else ""
    return (
        f"<xdr:twoCellAnchor{attr}><xdr:from>{frm}</xdr:from><xdr:to>{to}</xdr:to>{body}"
        "<xdr:clientData/></xdr:twoCellAnchor>"
    )


DRAWING_RELS = rels(
    ("rIdImg", f"{DOC}/image", "../media/image1.png"),
    ("rIdSvg", f"{DOC}/image", "../media/image2.svg"),
    ("rIdLink", f"{DOC}/image", "https://example.com/linked.png"),
    external=("rIdLink",),
)
SHEET_RELS = rels(("rId1", f"{DOC}/drawing", "../drawings/drawing1.xml"))
TYPES = (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Default Extension="svg" ContentType="image/svg+xml"/></Types>'
)


def drawing_of(*anchors: str) -> xlio.Drawing | None:
    data = sheet_with(
        '<drawing r:id="rId1"/>',
        SHEET_RELS,
        **{
            "xl/drawings/drawing1.xml": drawing_xml(*anchors),
            "xl/drawings/_rels/drawing1.xml.rels": DRAWING_RELS,
            "xl/media/image1.png": b"\x89PNG",
            "xl/media/image2.svg": "<svg/>",
            "[Content_Types].xml": TYPES,
        },
    )
    return xlio.open(io.BytesIO(data))["Data"].drawing


def test_anchors_in_z_order():
    drawing = drawing_of(
        '<xdr:absoluteAnchor><xdr:pos x="127000" y="254000"/><xdr:ext cx="10" cy="20"/>'
        + pic(4, "abs")
        + "<xdr:clientData/></xdr:absoluteAnchor>",
        two_cell(
            marker(2, 127000, 3, 50800), marker(4, 254000, 6, 76200), pic(2, "two"), "oneCell"
        ),
        "<xdr:oneCellAnchor><xdr:from>" + marker(1, 0, 1, 0) + '</xdr:from><xdr:ext cx="381000" '
        'cy="254000"/>'
        + pic(3, "one", extra=' hidden="1"')
        + "<xdr:clientData/></xdr:oneCellAnchor>",
    )
    absolute, two, one = drawing.objects  # file order is z-order, across anchor kinds
    assert absolute.anchor == xlio.Anchor("absolute", ext=(10, 20), pos=(127000, 254000))
    assert two.anchor.from_ == xlio.Marker(2, 127000, 3, 50800)
    assert (two.anchor.from_.column_offset_points, two.anchor.from_.row_offset_points) == (
        10.0,
        4.0,
    )
    assert (two.anchor.to, two.anchor.edit_as) == (xlio.Marker(4, 254000, 6, 76200), "oneCell")
    assert (one.anchor.kind, one.anchor.ext, one.hidden) == ("oneCell", (381000, 254000), True)
    assert (two.kind, two.id, two.name, two.description) == ("picture", 2, "two", "alt two")
    assert (two.picture.embed_part, two.picture.content_type) == (
        "xl/media/image1.png",
        "image/png",
    )


def test_a_reversed_two_cell_anchor_is_read_as_excel_repairs_it():
    # Capture repair: Excel swaps the corners on each axis reversed by whole cells, after asking;
    # with alerts off it refuses the file, and check() refuses it.
    reversed_both = two_cell(marker(2, 0, 15, 0), marker(1, 0, 14, 0), pic(2, "r"))
    data = sheet_with(
        '<drawing r:id="rId1"/>',
        SHEET_RELS,
        **{
            "xl/drawings/drawing1.xml": drawing_xml(reversed_both),
            "xl/drawings/_rels/drawing1.xml.rels": DRAWING_RELS,
        },
    )
    drawing = xlio.open(io.BytesIO(data))["Data"].drawing
    assert drawing is not None
    (only,) = drawing.objects
    assert (only.anchor.from_, only.anchor.to, only.anchor.repaired) == (
        marker_of(1, 0, 14, 0), marker_of(2, 0, 15, 0), True,
    )  # fmt: skip
    with pytest.raises(xlio.PackageError, match="ends before it starts"):
        xlio.check(io.BytesIO(data))
    one_axis = two_cell(marker(2, 0, 1, 0), marker(1, 0, 5, 0), pic(2, "r"))
    (only,) = drawing_of(one_axis).objects
    assert (only.anchor.from_, only.anchor.to, only.anchor.repaired) == (
        marker_of(1, 0, 1, 0), marker_of(2, 0, 5, 0), True,
    )  # fmt: skip
    # Offsets reversed within one cell: Excel opens it as written.
    in_cell = two_cell(marker(2, 300, 1, 200), marker(2, 100, 1, 50), pic(2, "r"))
    (only,) = drawing_of(in_cell).objects
    assert (only.anchor.from_, only.anchor.repaired) == (marker_of(2, 300, 1, 200), False)


def marker_of(column: int, column_offset: int, row: int, row_offset: int) -> object:
    from xlio._drawing import Marker

    return Marker(column, column_offset, row, row_offset)


def test_picture_details():
    blip = (
        '<a:blip r:embed="rIdImg" r:link="rIdLink"><a:extLst>'
        '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
        '<asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>'
        '<a:srcRect l="25000" b="50000" r="-10000"/>'
    )
    body = (
        pic(5, "p", blip)
        .replace(
            '<a:xfrm><a:off x="0" y="0"/>',
            '<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/>',
        )
        .replace(
            "</xdr:spPr>",
            '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'
            "</xdr:spPr>",
        )
    )
    (obj,) = drawing_of(two_cell(marker(0, 0, 0, 0), marker(2, 0, 2, 0), body)).objects
    p = obj.picture
    assert p.embed_part == "xl/media/image1.png"
    # The linked target is given as written, never followed.
    assert (p.link_target, p.link_external) == ("https://example.com/linked.png", True)
    assert (p.svg_part, p.crop) == ("xl/media/image2.svg", (0.25, 0.0, -0.1, 0.5))
    assert (p.rotation, p.flip_h, p.flip_v) == (90.0, True, False)
    assert p.line.get("w") == "12700"


def test_alternate_content_keeps_its_place():
    def alternate(requires: str, choice: str, fallback: str) -> str:
        return (
            f'<mc:AlternateContent><mc:Choice Requires="{requires}">{choice}</mc:Choice>'
            f"<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"
        )

    box = (marker(0, 0, 0, 0), marker(1, 0, 1, 0))
    drawing = drawing_of(
        two_cell(*box, pic(1, "first")),
        alternate("a14", two_cell(*box, pic(2, "a14 choice")), two_cell(*box, pic(3, "fb"))),
        alternate(
            "u", two_cell(*box, pic(4, "unknown choice")), two_cell(*box, pic(5, "fallback"))
        ),
        two_cell(*box, pic(6, "last")),
    )
    names = [(o.name, o.branch) for o in drawing.objects]
    assert names == [
        ("first", None),
        ("a14 choice", "choice"),
        ("fallback", "fallback"),
        ("last", None),
    ]
    alt = drawing.objects[2].alternate
    assert alt.choices[0][0] == frozenset({"urn:unknown"})  # the prefix, resolved
    assert alt.choose(frozenset({"urn:unknown"})) is alt.choices[0][1]


def test_a_chart_and_a_shape():
    frame = (
        '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="7" name="Chart 1"/>'
        "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><a:graphic><a:graphicData "
        'uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart '
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/>'
        "</a:graphicData></a:graphic></xdr:graphicFrame>"
    )
    shape = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="Box"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>'
    box = (marker(0, 0, 0, 0), marker(1, 0, 1, 0))
    chart, sp = drawing_of(two_cell(*box, frame), two_cell(*box, shape)).objects
    assert (chart.kind, chart.name, chart.picture) == ("chart", "Chart 1", None)
    assert (sp.kind, sp.id) == ("shape", 8)


# VML -------------------------------------------------------------------------------------------

VML = (
    b'<xml xmlns:v="urn:schemas-microsoft-com:vml" '
    b'xmlns:o="urn:schemas-microsoft-com:office:office"'
    b' xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap '
    b'v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600"'
    b' o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/></v:shapetype>'
    b'<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;'
    b'width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" '
    b'o:insetmode="auto">'
    b'<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
    b"<v:textbox><div style='text-align:left'><br></div></v:textbox>"  # <br> unclosed: not XML
    b'<x:ClientData ObjectType="Note"><x:MoveWithCells/>'
    b"<x:Anchor> 2, 15, 0, 10, 4, 15, 4, 4</x:Anchor>"
    b"<x:AutoFill>False</x:AutoFill><x:Visible/><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData>"
    b'</v:shape><v:shape id="_x0000_s1026" type="#_x0000_t201" strokecolor="#0000ff" stroked="t">'
    b'<undeclared:thing/><x:ClientData ObjectType="Button"><x:Anchor>1, 2</x:Anchor></x:ClientData>'
    b"</v:shape></xml>"
)


def vml_package(data: bytes) -> bytes:
    return sheet_with(
        '<legacyDrawing r:id="rId2"/>',
        rels(("rId2", f"{DOC}/vmlDrawing", "../drawings/vmlDrawing1.vml")),
        **{"xl/drawings/vmlDrawing1.vml": data},
    )


def test_vml_that_is_not_xml_is_read_by_the_extractor():
    vml = xlio.open(io.BytesIO(vml_package(VML)))["Data"].legacy_drawing
    note, button = vml.shapes
    assert (note.id, note.type, note.object_type) == ("_x0000_s1025", "#_x0000_t202", "Note")
    assert (note.anchor, note.visible, note.row, note.column) == (
        (2, 15, 0, 10, 4, 15, 4, 4),
        True,
        1,
        1,
    )
    assert "margin-left:59.25pt" in note.style
    assert note.attributes["fillcolor"] == "#ffffe1"
    assert note.fill == {"color2": "#ffffe1"}
    assert note.shadow == {"on": "t", "color": "black", "obscured": "t"}
    assert (button.object_type, button.anchor, button.visible) == ("Button", None, False)
    assert button.attributes["strokecolor"] == "#0000ff"
    assert vml.data == VML


def test_vml_is_bounded():
    limits = xlio.ReadLimits(max_vml_bytes=100)
    sheet = xlio.open(io.BytesIO(vml_package(VML)), limits=limits)["Data"]
    with pytest.raises(xlio.LimitError, match="VML"):
        sheet.legacy_drawing  # noqa: B018


@bounded(seconds=60)
def test_crafted_vml_costs_linear_time():
    data = b"<v:shape " * 400_000 + b"<x:Anchor " * 400_000
    start = time.perf_counter()
    vml = xlio.open(io.BytesIO(vml_package(data)))["Data"].legacy_drawing
    assert vml.shapes == ()
    assert time.perf_counter() - start < 10


# Notes and threaded comments -------------------------------------------------------------------

COMMENTS = (
    f'<comments xmlns="{MAIN}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">'
    "<authors><author>Ann</author><author>tc={1}</author></authors><commentList>"
    '<comment ref="B2" authorId="0"><text><r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr>'
    '<t>Ann:</t></r><r><t xml:space="preserve">\n hello</t></r></text></comment>'
    '<comment ref="C3" authorId="7" xr:uid="{U1}"><text><t>orphan</t></text></comment>'
    "</commentList></comments>"
)
THREADED = (
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
    '<threadedComment ref="C3" dT="2026-11-11T10:00:00.00" personId="{P1}" id="{T1}" done="1">'
    "<text>first</text></threadedComment>"
    '<threadedComment ref="C3" dT="2026-11-11T10:05:00.00" personId="{P9}" id="{T2}" '
    'parentId="{T1}">'
    "<text>reply</text></threadedComment></ThreadedComments>"
)
PERSONS = (
    '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
    '<person displayName="Ann" id="{P1}" userId="ann" providerId="None"/></personList>'
)


def test_notes_threads_and_persons():
    workbook_rels = rels(
        ("rId1", f"{DOC}/worksheet", "worksheets/sheet1.xml"),
        (
            "rId5",
            "http://schemas.microsoft.com/office/2017/10/relationships/person",
            "persons/person.xml",
        ),
    )
    sheet_rels = rels(
        ("rId3", f"{DOC}/comments", "../comments1.xml"),
        ("rId4", "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
         "../threadedComments/threadedComment1.xml"),
    )  # fmt: skip
    data = sheet_with(
        "",
        sheet_rels,
        **{
            "xl/_rels/workbook.xml.rels": workbook_rels,
            "xl/comments1.xml": COMMENTS,
            "xl/threadedComments/threadedComment1.xml": THREADED,
            "xl/persons/person.xml": PERSONS,
        },
    )
    with xlio.open(io.BytesIO(data)) as wb:
        sheet = wb["Data"]
        note, orphan = sheet.comments
        assert (note.ref, note.author, note.text.text) == ("B2", "Ann", "Ann:\n hello")
        assert (note.text.runs[0].font.bold, note.text.runs[0].font.name) == (True, "Tahoma")
        assert (orphan.author, orphan.author_id, orphan.uid) == (None, 7, "{U1}")
        first, reply = sheet.threaded_comments
        assert (first.ref, first.done, first.text, first.person.display_name) == (
            "C3",
            True,
            "first",
            "Ann",
        )
        assert (reply.parent_id, reply.person, reply.person_id) == ("{T1}", None, "{P9}")
        assert first.date == "2026-11-11T10:00:00.00"
        assert wb.persons["{P1}"].user_id == "ann"


# Files from other writers ---------------------------------------------------------------------


def png(width: int = 2, height: int = 2) -> bytes:
    def chunk(kind: bytes, body: bytes) -> bytes:
        return (
            struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))
        )

    raw = b"".join(b"\0" + b"\xff\x00\x00" * width for _ in range(height))
    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def test_what_xlsxwriter_writes(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    image = tmp_path / "red.png"
    image.write_bytes(png())
    path = tmp_path / "x.xlsx"
    book = xlsxwriter.Workbook(path)
    ws = book.add_worksheet("Data")
    ws.insert_image("C4", str(image), {"x_offset": 10, "y_offset": 4, "description": "red"})
    ws.insert_image("F4", str(image), {"object_position": 3})  # absolute
    ws.write_comment("B2", "a note", {"visible": True, "author": "Ann"})
    ws.set_background(str(image))
    book.close()
    with xlio.open(path) as wb:
        sheet = wb["Data"]
        first, second = sheet.drawing.objects
        assert (first.kind, first.description) == ("picture", "red")
        assert first.anchor.from_ == xlio.Marker(2, 10 * 9525, 3, 4 * 9525)  # pixels as EMU
        assert first.picture.content_type == "image/png"
        assert wb.parts.read(first.picture.embed_part) == image.read_bytes()
        assert (second.anchor.kind, second.anchor.edit_as) == ("twoCell", "absolute")
        (note,) = sheet.comments
        assert (note.ref, note.author, note.text.text) == ("B2", "Ann", "a note")
        (shape,) = sheet.legacy_drawing.shapes
        assert (shape.object_type, shape.visible, shape.row, shape.column) == ("Note", True, 1, 1)
        assert wb.parts.read(sheet.background) == image.read_bytes()
        xlio.check(path)


def test_what_openpyxl_writes(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.comments import Comment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"].comment = Comment("hello", "Bob")
    path = tmp_path / "o.xlsx"
    wb.save(path)
    with xlio.open(path) as book:
        sheet = book["Data"]
        (note,) = sheet.comments
        assert (note.ref, note.author, note.text.text) == ("A1", "Bob", "hello")
        (shape,) = sheet.legacy_drawing.shapes
        assert (shape.object_type, shape.row, shape.column, shape.anchor) == ("Note", 0, 0, None)
        assert sheet.drawing is None
