"""A sheet's drawing: its anchored objects in z-order, pictures with their media, and VML.

Anchors are read as written, in EMU (12,700 to a point, as Excel converts them); placing them on
the grid (an offset past its cell is clamped to the cell's far edge, hidden rows shrink a
two-cell box) needs the sheet's geometry and is the consumer's. A two-cell anchor whose ``to``
cell comes before its ``from`` cell is read as Excel repairs it (``Anchor.repaired``).
Objects come in file order, which is their z-order across anchor kinds; an object inside
``mc:AlternateContent`` keeps its place. A linked picture's target is given, never followed.

VML is often not well-formed XML; its shapes are read by bounded regular expressions, never by a
lenient XML parser.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Mapping
from dataclasses import dataclass, field, replace
from typing import Literal, cast

from xlio._package import LimitError, PackageError
from xlio._types import finite, flag, integer, local

__all__ = [
    "Alternate",
    "Anchor",
    "AnchorKind",
    "Drawing",
    "DrawingObject",
    "Marker",
    "ObjectKind",
    "Picture",
    "Vml",
    "VmlShape",
    "read_drawing",
    "read_vml",
]

EMU_PER_POINT = 12_700
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
_REL = "/relationships}"
UNDERSTOOD = frozenset({A14})
"""Namespaces whose ``AlternateContent`` choice Excel for Mac was seen drawing:
``a14``; a choice that requires another namespace was not drawn, its fallback was."""

AnchorKind = Literal["twoCell", "oneCell", "absolute"]
ObjectKind = Literal[
    "picture", "shape", "group", "connector", "chart", "graphicFrame", "contentPart", "other"
]
_KINDS: dict[str, ObjectKind] = {
    "pic": "picture",
    "sp": "shape",
    "grpSp": "group",
    "cxnSp": "connector",
    "graphicFrame": "graphicFrame",
    "contentPart": "contentPart",
}
_EDIT_AS: tuple[AnchorKind, ...] = ("twoCell", "oneCell", "absolute")


@dataclass(frozen=True, slots=True)
class Marker:
    """A corner of an anchor: a cell (0-based, as written) and offsets into it in EMU."""

    column: int
    column_offset: int
    row: int
    row_offset: int

    @property
    def column_offset_points(self) -> float:
        """The column offset in points (EMU ÷ 12,700)."""
        return self.column_offset / EMU_PER_POINT

    @property
    def row_offset_points(self) -> float:
        """The row offset in points."""
        return self.row_offset / EMU_PER_POINT


@dataclass(frozen=True, slots=True)
class Anchor:
    """Where an object is anchored.

    Attributes:
        kind: ``twoCell`` (``from`` and ``to``), ``oneCell`` (``from`` and ``ext``) or
            ``absolute`` (``pos`` and ``ext``).
        from_: The top-left corner, or ``None``.
        to: The bottom-right corner of a two-cell anchor, or ``None``.
        ext: Width and height in EMU, or ``None``.
        pos: Left and top in EMU from the sheet's corner, or ``None``.
        edit_as: How a two-cell anchor moves with later edits, or ``None``; it does not change
            how a loaded file is drawn.
        repaired: The file's anchor ends before it starts, by whole cells on either axis, and
            ``from_`` and ``to`` are as Excel repairs it: the corners swapped on that axis
            (capture ``repair``). Excel asks before repairing a file, and refuses it with alerts
            off; :func:`xlio.check` refuses it.
    """

    kind: AnchorKind
    from_: Marker | None = None
    to: Marker | None = None
    ext: tuple[int, int] | None = None
    pos: tuple[int, int] | None = None
    edit_as: AnchorKind | None = None
    repaired: bool = False


@dataclass(frozen=True, slots=True)
class Alternate:
    """An ``mc:AlternateContent``, raw.

    Its choices come with the namespaces each requires, then its fallback.
    """

    choices: tuple[tuple[frozenset[str], ET.Element], ...]
    fallback: ET.Element | None = None

    def choose(self, understood: frozenset[str] = UNDERSTOOD) -> ET.Element | None:
        """The branch a reader that understands these namespaces takes.

        The first choice whose requirements it meets, else the fallback.
        """
        for requires, element in self.choices:
            if requires <= understood:
                return element
        return self.fallback


@dataclass(frozen=True, slots=True)
class Picture:
    """A picture's image and how it is shown.

    Attributes:
        embed_part: The embedded image's part, or ``None``.
        content_type: Its content type (``image/png``), or ``None``.
        link_target: A linked image's target as written (a path or URL), or ``None``. Excel
            loads it on open and, with both, draws it over the embedded one; xlio never
            follows it.
        link_external: The link's target mode is external.
        svg_part: An SVG image (``asvg:svgBlip``), which Excel draws over the embedded one,
            or ``None``.
        crop: Left, top, right and bottom crop as fractions (``a:srcRect``, negative outsets),
            or ``None``.
        rotation: Clockwise rotation in degrees.
        flip_h: Mirrored horizontally (before rotation).
        flip_v: Mirrored vertically.
        line: The outline (``a:ln``) or a style reference (``a:style/a:lnRef``), raw, or ``None``.
    """

    embed_part: str | None = None
    content_type: str | None = None
    link_target: str | None = None
    link_external: bool = False
    svg_part: str | None = None
    crop: tuple[float, float, float, float] | None = None
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    line: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class DrawingObject:
    """An anchored object, in z-order (file order).

    Attributes:
        anchor: Where it is anchored.
        kind: ``picture``, ``shape``, ``group``, ``connector``, ``chart`` (a graphic frame
            holding a chart), ``graphicFrame``, ``contentPart`` or ``other``.
        id: ``cNvPr id``, or ``None``.
        name: ``cNvPr name``, or ``None``.
        description: The alternative text (``descr``), or ``None``.
        title: The title, or ``None``.
        hidden: Not drawn (``cNvPr hidden``).
        picture: For a picture, its image; else ``None``.
        alternate: The ``AlternateContent`` it came from, or ``None``.
        branch: ``choice`` or ``fallback``: the branch read, for an object from
            ``AlternateContent``; else ``None``.
        element: The object's element (``xdr:pic``, ``xdr:sp``, ...), or ``None`` when the
            anchor holds none.
        source: The anchor element as written.
    """

    anchor: Anchor
    kind: ObjectKind
    id: int | None = None
    name: str | None = None
    description: str | None = None
    title: str | None = None
    hidden: bool = False
    picture: Picture | None = None
    alternate: Alternate | None = None
    branch: Literal["choice", "fallback"] | None = None
    element: ET.Element | None = field(default=None, compare=False, repr=False)
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class Drawing:
    """A sheet's drawing part: its objects in z-order."""

    part: str
    objects: tuple[DrawingObject, ...] = ()


def _marker(element: ET.Element | None) -> Marker | None:
    if element is None:
        return None
    values = {local(c.tag): (c.text or "").strip() for c in element}
    return Marker(
        integer(values.get("col"), 0) or 0,
        integer(values.get("colOff"), 0) or 0,
        integer(values.get("row"), 0) or 0,
        integer(values.get("rowOff"), 0) or 0,
    )


def _first(element: ET.Element, name: str) -> ET.Element | None:
    return next((c for c in element if local(c.tag) == name), None)


def _deep(element: ET.Element, name: str) -> ET.Element | None:
    return next((e for e in element.iter() if local(e.tag) == name), None)


def _pair(element: ET.Element | None, a: str, b: str) -> tuple[int, int] | None:
    if element is None:
        return None
    return integer(element.get(a), 0) or 0, integer(element.get(b), 0) or 0


def _rel_attr(element: ET.Element, name: str) -> str | None:
    for key, value in element.attrib.items():
        if key.endswith(_REL + name) or key == name:
            return value
    return None


def _picture(
    pic: ET.Element,
    rels: Mapping[str, tuple[str, str | None, bool]],
    content_type: Mapping[str, str | None],
) -> Picture:
    blip = _deep(pic, "blip")
    embed_part = link_target = svg_part = None
    link_external = False
    crop = None
    if blip is not None:
        embed = _rel_attr(blip, "embed")
        if embed is not None and embed in rels:
            embed_part = rels[embed][1]
        link = _rel_attr(blip, "link")
        if link is not None and link in rels:
            link_target, _, link_external = rels[link]
        svg = next((e for e in blip.iter() if e.tag == f"{{{SVG}}}svgBlip"), None)
        if svg is not None:
            svg_id = _rel_attr(svg, "embed")
            svg_part = rels[svg_id][1] if svg_id is not None and svg_id in rels else None
    rect = _deep(pic, "srcRect")
    if rect is not None:
        crop = tuple(finite(rect.get(k), 0.0) / 100_000 for k in ("l", "t", "r", "b"))
    xfrm = _deep(pic, "xfrm")
    ln = _deep(pic, "ln")
    if ln is None:
        style = _deep(pic, "style")
        ln = _first(style, "lnRef") if style is not None else None
    return Picture(
        embed_part,
        content_type.get(embed_part) if embed_part is not None else None,
        link_target,
        link_external,
        svg_part,
        crop,  # type: ignore[arg-type]
        finite(xfrm.get("rot"), 0.0) / 60_000 if xfrm is not None else 0.0,
        flag(xfrm.get("flipH"), False) if xfrm is not None else False,
        flag(xfrm.get("flipV"), False) if xfrm is not None else False,
        ln,
    )


def _kind(element: ET.Element) -> ObjectKind:
    name = local(element.tag)
    kind = _KINDS.get(name, "other")
    if kind == "graphicFrame" and _deep(element, "chart") is not None:
        return "chart"
    return kind


_OBJECTS = frozenset(_KINDS)


def _object(
    anchor_element: ET.Element,
    rels: Mapping[str, tuple[str, str | None, bool]],
    content_type: Mapping[str, str | None],
    part: str,
    alternate: Alternate | None = None,
    branch: Literal["choice", "fallback"] | None = None,
) -> DrawingObject:
    name = local(anchor_element.tag)
    kind: AnchorKind = (
        "twoCell"
        if name == "twoCellAnchor"
        else "oneCell"
        if name == "oneCellAnchor"
        else "absolute"
    )
    from_ = _marker(_first(anchor_element, "from"))
    to = _marker(_first(anchor_element, "to"))
    repaired = False
    if kind == "twoCell" and from_ is not None and to is not None:
        # Excel repairs an anchor whose to cell comes before its from cell, on either axis, by
        # swapping the corners on that axis; offsets reversed within one cell are kept
        # (capture repair).
        if to.column < from_.column:
            from_, to = (
                replace(from_, column=to.column, column_offset=to.column_offset),
                replace(to, column=from_.column, column_offset=from_.column_offset),
            )
            repaired = True
        if to.row < from_.row:
            from_, to = (
                replace(from_, row=to.row, row_offset=to.row_offset),
                replace(to, row=from_.row, row_offset=from_.row_offset),
            )
            repaired = True
    edit_as = anchor_element.get("editAs")
    anchor = Anchor(
        kind,
        from_,
        to,
        _pair(_first(anchor_element, "ext"), "cx", "cy"),
        _pair(_first(anchor_element, "pos"), "x", "y"),
        next((k for k in _EDIT_AS if k == edit_as), None),
        repaired,
    )
    element = next((c for c in anchor_element if local(c.tag) in _OBJECTS), None)
    if element is None:
        return DrawingObject(
            anchor, "other", alternate=alternate, branch=branch, source=anchor_element
        )
    props = _deep(element, "cNvPr")
    get = props.get if props is not None else ET.Element("x").get
    kind_of = _kind(element)
    return DrawingObject(
        anchor,
        kind_of,
        integer(get("id")),
        get("name"),
        get("descr"),
        get("title"),
        flag(get("hidden"), False),
        _picture(element, rels, content_type) if kind_of == "picture" else None,
        alternate,
        branch,
        element,
        anchor_element,
    )


def _requires(text: str, namespaces: Mapping[str, str]) -> frozenset[str]:
    return frozenset(namespaces.get(prefix, prefix) for prefix in text.split())


def read_drawing(
    data: bytes,
    part: str,
    rels: Mapping[str, tuple[str, str | None, bool]],
    content_type: Mapping[str, str | None],
) -> Drawing:
    """A drawing part's objects.

    Args:
        data: The part's bytes (already checked for a DTD and bounded).
        part: Its name.
        rels: Relationship id → (target as written, target part or ``None``, external).
        content_type: Part name → content type, for the images.

    Raises:
        PackageError: The part is not well-formed.
    """
    namespaces: dict[str, str] = {}
    root: ET.Element | None = None
    parser: ET.XMLPullParser[ET.Element] = ET.XMLPullParser(events=("start-ns", "start"))
    try:
        parser.feed(data)
        parser.close()
    except ET.ParseError as exc:
        raise PackageError(f"{part}: {exc}") from exc
    for pair in parser.read_events():
        event, item = cast(tuple[str, object], pair)
        if event == "start-ns":
            prefix, uri = cast(tuple[str, str], item)
            namespaces.setdefault(prefix, uri)
        elif root is None:
            root = cast(ET.Element, item)
    if root is None:
        raise PackageError(f"{part}: no root element")
    objects = []
    for child in root:
        name = local(child.tag)
        if name in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            objects.append(_object(child, rels, content_type, part))
        elif name == "AlternateContent":
            choices = tuple(
                (_requires(c.get("Requires", ""), namespaces), c)
                for c in child
                if local(c.tag) == "Choice"
            )
            fallback = _first(child, "Fallback")
            alternate = Alternate(choices, fallback)
            chosen = alternate.choose()
            branch: Literal["choice", "fallback"] = "fallback" if chosen is fallback else "choice"
            objects.extend(
                _object(anchor, rels, content_type, part, alternate, branch)
                for anchor in (chosen if chosen is not None else ())
                if local(anchor.tag) in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")
            )
    return Drawing(part, tuple(objects))


# VML -------------------------------------------------------------------------------------------

# Every pattern stops at the next "<", so the scan is linear in the part's size.
_SHAPE = re.compile(rb"<(?:[A-Za-z_][\w.-]{0,40}:)?shape\b([^<>]{0,65536})>")
_ATTR = re.compile(rb"""([\w:.-]{1,80})\s*=\s*(?:"([^"]{0,65536})"|'([^']{0,65536})')""")
_CLIENT = re.compile(rb"<(?:[\w.-]{1,40}:)?ClientData\b([^<>]{0,4096})>")
_FIELD = re.compile(
    rb"<(?:[\w.-]{1,40}:)?(Anchor|Visible|Row|Column)\b[^<>]{0,1024}?(/>|>([^<]{0,4096})<)"
)
_CHILD = re.compile(rb"<(?:[\w.-]{1,40}:)?(fill|stroke|shadow)\b([^<>]{0,4096})>")
_END = re.compile(rb"</(?:[\w.-]{1,40}:)?shape\s*>")


@dataclass(frozen=True, slots=True)
class VmlShape:
    """A VML shape: a note's box, a form control, a VML picture.

    Attributes:
        id: The shape's ``id``, or ``None``.
        type: Its ``type`` (``#_x0000_t202`` for a note), or ``None``.
        object_type: ``x:ClientData ObjectType`` (``Note``, ``Button``, ...), or ``None``.
        anchor: ``x:Anchor`` as eight integers (left column, left offset, top row, top offset,
            right column, right offset, bottom row, bottom offset; offsets in VML pixels, one
            a point on the Mac), or ``None``.
        visible: ``x:Visible`` is present (a shown note).
        row: ``x:Row`` (0-based), or ``None``.
        column: ``x:Column`` (0-based), or ``None``.
        style: The ``style`` attribute as written, or ``None``.
        attributes: The shape's attributes as written (``fillcolor``, ``strokecolor``, ...).
        fill: ``v:fill``'s attributes, or ``None``.
        stroke: ``v:stroke``'s attributes, or ``None``.
        shadow: ``v:shadow``'s attributes, or ``None``.
    """

    id: str | None
    type: str | None
    object_type: str | None
    anchor: tuple[int, ...] | None
    visible: bool
    row: int | None
    column: int | None
    style: str | None
    attributes: Mapping[str, str]
    fill: Mapping[str, str] | None = None
    stroke: Mapping[str, str] | None = None
    shadow: Mapping[str, str] | None = None


@dataclass(frozen=True, slots=True)
class Vml:
    """A VML drawing part: its bytes, and the shapes the extractor found."""

    part: str
    data: bytes = field(repr=False)
    shapes: tuple[VmlShape, ...] = ()


def _attributes(text: bytes) -> dict[str, str]:
    out = {}
    for m in _ATTR.finditer(text):
        value = m[2] if m[2] is not None else m[3]
        out[m[1].decode("utf-8", "replace")] = value.decode("utf-8", "replace")
    return out


def read_vml(data: bytes, part: str, limit: int) -> Vml:
    """A VML part's shapes, read by bounded regular expressions.

    Raises:
        LimitError: The part is larger than ``limit`` (``ReadLimits.max_vml_bytes``).
    """
    if len(data) > limit:
        raise LimitError(f"{part}: VML of {len(data)} bytes, more than the limit of {limit}")
    shapes = []
    for m in _SHAPE.finditer(data):
        attrs = _attributes(m[1])
        start = m.end()
        end_match = _END.search(data, start) if not m[1].rstrip().endswith(b"/") else None
        body = data[start : end_match.start()] if end_match is not None else b""
        client = _CLIENT.search(body)
        fields: dict[str, bytes | None] = {}
        for f in _FIELD.finditer(body):
            fields.setdefault(f[1].decode(), f[3])
        anchor = None
        if fields.get("Anchor"):
            numbers = [integer(n.decode(errors="replace")) for n in fields["Anchor"].split(b",")]  # type: ignore[union-attr]
            if len(numbers) == 8 and all(n is not None for n in numbers):
                anchor = tuple(n for n in numbers if n is not None)
        children = {c[1].decode(): _attributes(c[2]) for c in reversed(list(_CHILD.finditer(body)))}
        row, column = fields.get("Row"), fields.get("Column")
        shapes.append(
            VmlShape(
                attrs.get("id"),
                attrs.get("type"),
                _attributes(client[1]).get("ObjectType") if client is not None else None,
                anchor,
                "Visible" in fields,
                integer(row.decode(errors="replace")) if row else None,
                integer(column.decode(errors="replace")) if column else None,
                attrs.get("style"),
                attrs,
                children.get("fill"),
                children.get("stroke"),
                children.get("shadow"),
            )
        )
    return Vml(part, data, tuple(shapes))
