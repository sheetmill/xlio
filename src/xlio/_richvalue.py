"""Rich values: what stands behind a cell's ``vm`` (a picture in a cell, a data type, an error).

A cell with a rich value caches a plain value (``#VALUE!`` for a picture in a cell); the rich
value itself lives in the workbook's rich data parts, reached through the value metadata. These
records are distinct from rich text (``Cell.rich``, ``RichString``).
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field

from xlio._types import integer, local

__all__ = ["ImageValue", "RichValue", "read_rich_values"]


@dataclass(frozen=True, slots=True)
class RichValue:
    """A rich value, raw: its structure's type and its keyed values as written.

    Attributes:
        type: The structure's type (``_localImage``, ``_webImage``, ``_error``, ``_entity``,
            ...), as written.
        values: Key name → the value's text as written, in the structure's key order.
        source: The ``rv`` element as written.
    """

    type: str
    values: Mapping[str, str | None]
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class ImageValue(RichValue):
    """A picture in a cell: a ``_localImage`` (placed in the cell) or ``_webImage`` (``IMAGE``).

    The cell caches ``#VALUE!``; Excel draws the picture in the cell instead.

    Attributes:
        image_part: A placed picture's part, through the rich value relationships, or
            ``None`` (a web image, whose identifier indexes the web image part, not modeled;
            or a relationship that leads nowhere).
        alt_text: Alternative text (``Text``), or ``None``.
        calc_origin: How the value came to be (``CalcOrigin``), as written, or ``None``.
        sizing: How the picture is sized in the cell (``ImageSizing``, ``IMAGE``'s sizing
            argument), or ``None``; a placed picture has none and fits the cell.
        height: The height ``IMAGE`` was given (``ImageHeight``), or ``None``.
        width: The width ``IMAGE`` was given (``ImageWidth``), or ``None``.
    """

    image_part: str | None = None
    alt_text: str | None = None
    calc_origin: int | None = None
    sizing: int | None = None
    height: int | None = None
    width: int | None = None


_IMAGES = {"_localImage": "_rvRel:LocalImageIdentifier", "_webImage": None}


def read_rich_values(
    structures: ET.Element | None,
    values: ET.Element | None,
    relations: Sequence[str | None],
) -> tuple[RichValue, ...]:
    """Every rich value, by its index; ``relations`` are the targets ``_rvRel`` keys index."""
    if structures is None or values is None:
        return ()
    shapes = [
        (s.get("t", ""), [k.get("n", "") for k in s if local(k.tag) == "k"])
        for s in structures
        if local(s.tag) == "s"
    ]
    out: list[RichValue] = []
    for rv in (e for e in values if local(e.tag) == "rv"):
        index = integer(rv.get("s"))
        kind, keys = shapes[index] if index is not None and 0 <= index < len(shapes) else ("", [])
        texts = [v.text for v in rv if local(v.tag) == "v"]
        fields = dict(zip(keys, texts, strict=False))
        if kind not in _IMAGES:
            out.append(RichValue(kind, fields, rv))
            continue
        identifier = _IMAGES[kind]
        rel = integer(fields.get(identifier)) if identifier is not None else None
        out.append(
            ImageValue(
                kind,
                fields,
                rv,
                relations[rel] if rel is not None and 0 <= rel < len(relations) else None,
                fields.get("Text"),
                integer(fields.get("CalcOrigin")),
                integer(fields.get("ImageSizing")),
                integer(fields.get("ImageHeight")),
                integer(fields.get("ImageWidth")),
            )
        )
    return tuple(out)
