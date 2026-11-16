"""The workbook's theme: its color scheme and font scheme."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Mapping
from dataclasses import dataclass, field

from xlio._types import local

__all__ = ["ColorScheme", "SchemeColor", "SchemeFonts", "Theme", "read_theme"]

_SLOTS = (
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
    "accent6", "hlink", "folHlink",
)  # fmt: skip
# Excel's theme color index → the scheme's slot: the light and dark pairs are swapped.
_THEME_ORDER = (1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11)
_RGB = re.compile(r"[0-9A-Fa-f]{6}")


@dataclass(frozen=True, slots=True)
class SchemeColor:
    """One color of the scheme.

    Attributes:
        rgb: ``RRGGBB`` in upper case: an ``srgbClr``'s value, or a ``sysClr``'s ``lastClr``
            (the system color when the file was saved); ``None`` when neither is six hex digits.
        system: The ``sysClr`` value (``windowText``, ``window``), or ``None`` for an RGB color.
        source: The element as written.
    """

    rgb: str | None
    system: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class ColorScheme:
    """The theme's twelve colors (``a:clrScheme``), by their names in the file.

    ``dk1``/``lt1`` are text and background 1, ``dk2``/``lt2`` text and background 2. A cell's
    theme color index counts them in the order lt1, dk1, lt2, dk2, then accent1 to folHlink
    (:meth:`by_index`).
    """

    name: str
    dk1: SchemeColor
    lt1: SchemeColor
    dk2: SchemeColor
    lt2: SchemeColor
    accent1: SchemeColor
    accent2: SchemeColor
    accent3: SchemeColor
    accent4: SchemeColor
    accent5: SchemeColor
    accent6: SchemeColor
    hlink: SchemeColor
    folHlink: SchemeColor  # noqa: N815 - the scheme's own name

    def by_index(self, index: int) -> SchemeColor | None:
        """The color a cell's theme color index names, in Excel's order (pairs swapped)."""
        if not 0 <= index < len(_SLOTS):
            return None
        return getattr(self, _SLOTS[_THEME_ORDER[index]])  # type: ignore[no-any-return]


@dataclass(frozen=True, slots=True)
class SchemeFonts:
    """The major (headings) or minor (body) fonts of the scheme.

    Attributes:
        latin: The Latin typeface.
        east_asian: The East Asian typeface (``a:ea``), often empty.
        complex: The complex-script typeface (``a:cs``), often empty.
        scripts: Script code (``Jpan``, ``Hans``) → typeface, from ``a:font``.
    """

    latin: str = ""
    east_asian: str = ""
    complex: str = ""
    scripts: Mapping[str, str] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class Theme:
    """The workbook's theme.

    Attributes:
        name: The theme's name, or ``None``.
        colors: The color scheme, or ``None`` when the part has none.
        font_scheme_name: The font scheme's name, or ``None``.
        major_fonts: The headings fonts.
        minor_fonts: The body fonts.
        part: The theme part's name.
        source: The whole part, for what is not modeled (the format scheme, effects).
    """

    name: str | None
    colors: ColorScheme | None
    font_scheme_name: str | None
    major_fonts: SchemeFonts
    minor_fonts: SchemeFonts
    part: str
    source: ET.Element = field(compare=False, repr=False)


def _color(element: ET.Element | None) -> SchemeColor:
    if element is None:
        return SchemeColor(None)
    for child in element:
        name = local(child.tag)
        if name == "srgbClr":
            value = child.get("val", "")
            return SchemeColor(value.upper() if _RGB.fullmatch(value) else None, None, element)
        if name == "sysClr":
            last = child.get("lastClr", "")
            return SchemeColor(
                last.upper() if _RGB.fullmatch(last) else None, child.get("val"), element
            )
    return SchemeColor(None, None, element)


def _fonts(element: ET.Element | None) -> SchemeFonts:
    if element is None:
        return SchemeFonts()
    found: dict[str, str] = {}
    scripts: dict[str, str] = {}
    for child in element:
        name = local(child.tag)
        if name in ("latin", "ea", "cs"):
            found[name] = child.get("typeface", "")
        elif name == "font" and child.get("script"):
            scripts[child.get("script", "")] = child.get("typeface", "")
    return SchemeFonts(found.get("latin", ""), found.get("ea", ""), found.get("cs", ""), scripts)


def _first(element: ET.Element, name: str) -> ET.Element | None:
    return next((e for e in element.iter() if local(e.tag) == name), None)


def read_theme(root: ET.Element, part: str) -> Theme:
    """The theme from its part's element tree."""
    scheme = _first(root, "clrScheme")
    colors = None
    if scheme is not None:
        slots = {local(child.tag): child for child in scheme}
        colors = ColorScheme(scheme.get("name", ""), *(_color(slots.get(s)) for s in _SLOTS))
    fonts = _first(root, "fontScheme")
    major = _first(fonts, "majorFont") if fonts is not None else None
    minor = _first(fonts, "minorFont") if fonts is not None else None
    return Theme(
        root.get("name"),
        colors,
        fonts.get("name") if fonts is not None else None,
        _fonts(major),
        _fonts(minor),
        part,
        root,
    )
