"""Excel's built-in table and pivot table styles, as data.

The 144 presets of ECMA-376 Part 1, 5th edition (``presetTableStyles.xml``, shipped in
``xlio/data`` as Ecma International publishes it, its notice in ``NOTICE``): 60 table styles
(Light1 to Light21, Medium1 to Medium28, Dark1 to Dark11) and 84 pivot table styles. Excel
draws all 60 table presets exactly as the file defines them. Each preset comes with its own
differential formats, read into the same records as a workbook's own (:class:`xlio.TableStyle`,
:class:`xlio.DifferentialFormat`); :meth:`xlio.Styles.table_style` looks a workbook's own styles
up first, then these.
"""

from __future__ import annotations

import dataclasses
import functools
import importlib.resources
import xml.etree.ElementTree as ET

from xlio._styles import read_differential_formats
from xlio._tables import TableStyle, read_table_styles
from xlio._types import local

__all__ = ["names", "table_style"]


@functools.cache
def _presets() -> dict[str, TableStyle]:
    data = importlib.resources.files("xlio").joinpath("data/presetTableStyles.xml").read_bytes()
    out: dict[str, TableStyle] = {}
    for preset in ET.fromstring(data):
        dxfs = next((c for c in preset if local(c.tag) == "dxfs"), None)
        formats = read_differential_formats(dxfs) if dxfs is not None else ()
        for group in preset:
            if local(group.tag) == "tableStyles":
                for style in read_table_styles(group, formats, builtin=True).styles:
                    # Named by its container: two inner names in the file differ from it
                    # ("TableStyleLight14 " with a space, "TableStyleDark9 2").
                    named = dataclasses.replace(style, name=preset.tag)
                    out[named.name] = named
    return out


def names() -> tuple[str, ...]:
    """The presets' names, in the order of the data file."""
    return tuple(_presets())


def table_style(name: str) -> TableStyle | None:
    """The preset of this name (exact, case included), or ``None``."""
    return _presets().get(name)
