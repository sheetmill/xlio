"""Rich text: a shared or inline string's runs with their fonts, and its phonetic guides."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import BinaryIO, Literal, cast

from xlio._package import PackageError
from xlio._styles import Font
from xlio._types import choice, unsigned

__all__ = ["PhoneticProperties", "PhoneticRun", "RichString", "Run", "read_rich_strings", "rich"]

_SPACE = " \t\n\r"
_CHUNK = 1 << 16
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PhoneticType = Literal["halfwidthKatakana", "fullwidthKatakana", "Hiragana", "noConversion"]
_TYPES: tuple[PhoneticType, ...] = (
    "halfwidthKatakana", "fullwidthKatakana", "Hiragana", "noConversion",
)  # fmt: skip
PhoneticAlignment = Literal["noControl", "left", "center", "distributed"]
_ALIGNMENTS: tuple[PhoneticAlignment, ...] = ("noControl", "left", "center", "distributed")


@dataclass(frozen=True, slots=True)
class Run:
    """A run of rich text: its text and the font properties it sets (``None``: the cell's)."""

    text: str
    font: Font | None = None


@dataclass(frozen=True, slots=True)
class PhoneticRun:
    """A phonetic guide (``rPh``).

    Its text stands over the base text's characters ``start`` to ``end`` (0-based, end
    excluded).
    """

    text: str
    start: int
    end: int


@dataclass(frozen=True, slots=True)
class PhoneticProperties:
    """How phonetic guides are shown (``phoneticPr``).

    Attributes:
        font_id: The font of the guides (an index into ``Styles.fonts``).
        type: The characters used.
        alignment: The guides' alignment over the base text.
    """

    font_id: int
    type: PhoneticType = "fullwidthKatakana"
    alignment: PhoneticAlignment = "left"


@dataclass(frozen=True, slots=True)
class RichString:
    """A string with its runs and phonetic guides.

    Attributes:
        text: The text as Excel reads it, the same as the plain string.
        runs: The runs, in order; empty for a string without runs.
        phonetic: The phonetic guides.
        phonetic_properties: How the guides are shown, or ``None``.
        source: The element as written (``si`` or ``is``).
    """

    text: str
    runs: tuple[Run, ...] = ()
    phonetic: tuple[PhoneticRun, ...] = ()
    phonetic_properties: PhoneticProperties | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


def _name(tag: str) -> str:
    """A local name, for trees built with or without namespace processing."""
    return tag.rpartition("}")[2].rpartition(":")[2]


def _text(t: ET.Element, unescape: Callable[[str], str]) -> str:
    raw = t.text or ""
    preserve = (t.get(_XML_SPACE) or t.get("xml:space")) == "preserve"
    return unescape(raw if preserve else raw.strip(_SPACE))


def rich(element: ET.Element) -> RichString:
    """A ``si`` or ``is`` element as a rich string, text read as Excel reads it.

    Each ``<t>`` not marked ``xml:space="preserve"`` loses its outer whitespace (captured by
    xlrecalc:
    https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/whitespace.tsv),
    and ``_xHHHH_`` escapes are decoded.
    """
    from xlio._cells import unescape

    runs: list[Run] = []
    phonetic: list[PhoneticRun] = []
    properties = None
    plain: list[str] = []
    for child in element:
        name = _name(child.tag)
        if name == "t":
            plain.append(_text(child, unescape))
        elif name == "r":
            text = "".join(_text(t, unescape) for t in child if _name(t.tag) == "t")
            props = next((p for p in child if _name(p.tag) == "rPr"), None)
            font = Font.read(_renamed(props), kind="run") if props is not None else None
            runs.append(Run(text, font))
        elif name == "rPh":
            text = "".join(_text(t, unescape) for t in child if _name(t.tag) == "t")
            phonetic.append(
                PhoneticRun(text, unsigned(child.get("sb"), 0), unsigned(child.get("eb"), 0))
            )
        elif name == "phoneticPr":
            properties = PhoneticProperties(
                unsigned(child.get("fontId"), 0),
                choice(child.get("type"), _TYPES, "fullwidthKatakana"),
                choice(child.get("alignment"), _ALIGNMENTS, "left"),
            )
    text = "".join(plain) + "".join(r.text for r in runs)
    return RichString(text, tuple(runs), tuple(phonetic), properties, element)


def _renamed(element: ET.Element) -> ET.Element:
    """A run's properties with prefixed names (a tree built without namespaces) made local."""
    if all(":" not in e.tag for e in element.iter()):
        return element
    copy = ET.Element(_name(element.tag), dict(element.attrib))
    for child in element:
        ET.SubElement(copy, _name(child.tag), dict(child.attrib))
    return copy


def read_rich_strings(stream: BinaryIO, part: str) -> list[RichString]:
    """Every shared string with its runs, streamed one ``<si>`` at a time."""
    parser: ET.XMLPullParser[ET.Element] = ET.XMLPullParser(events=("start", "end"))
    out: list[RichString] = []
    root: ET.Element | None = None

    def drain() -> None:
        nonlocal root
        for pair in parser.read_events():
            event, element = cast(tuple[str, ET.Element], pair)
            if root is None:
                root = element
            elif event == "end" and _name(element.tag) == "si":
                out.append(rich(element))
                root.clear()  # strings already read are not kept twice

    try:
        while chunk := stream.read(_CHUNK):
            parser.feed(chunk)
            drain()
        parser.close()
        drain()
    except ET.ParseError as exc:
        raise PackageError(f"{part}: {exc}") from exc
    return out
