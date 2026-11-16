"""The public writer: ``patch``, over the surgical writer in ``_write``."""

from __future__ import annotations

import os
from collections.abc import Callable, Mapping, Sequence
from typing import BinaryIO, NamedTuple

from xlio import _write
from xlio._package import ReadLimits
from xlio._workbook import ErrorCode

__all__ = ["ArrayFill", "CachedValues", "CellEdit", "CellValue", "Translate", "patch"]

CellValue = float | str | bool | ErrorCode
"""A value to write: a number (``int`` is taken as a number), text, a boolean or an error."""

CachedValues = Callable[[int, int], CellValue | None]
"""A sheet's new cached formula results: ``lookup(row, col)`` (1-based) is called for each
formula cell, in file order, and returns its value, or ``None`` to keep the cell as it is."""

Translate = Callable[[str, int, int, int, int], str]
"""``translate(text, row, col, to_row, to_col)``: a formula's text moved from one cell to
another (0-based), used when a shared formula's master cell is replaced and its dependents need
their own text."""


class CellEdit(NamedTuple):
    """A cell's new content.

    Attributes:
        formula: The formula's text as a file stores it (no leading ``=``), or ``None`` for a
            constant.
        value: The constant, or the formula's cached result; ``None`` for a blank cell or a
            formula without a cached result.
        array_ref: Makes the formula an array formula over this range (``"B1:B5"``).
        dynamic: Marks the array formula as a dynamic array (with the workbook's cell metadata).
    """

    formula: str | None
    value: CellValue | None
    array_ref: str | None = None
    dynamic: bool = False


class ArrayFill(NamedTuple):
    """A range of constant cells, as an array formula fills them (1-based, inclusive).

    ``value_at(row, col)`` gives each cell's value when it is written, so the range is never
    held in memory; an edit of the same cell takes precedence.
    """

    r1: int
    c1: int
    r2: int
    c2: int
    value_at: Callable[[int, int], CellValue | None]


def _file_value(value: CellValue | None) -> _write.FileValue | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "b", value
    if isinstance(value, ErrorCode):
        return "e", value.value
    if isinstance(value, int | float):
        return "n", float(value)
    if isinstance(value, str):
        return "str", value
    raise TypeError(f"not a cell value: {value!r}")


def _lookup(cached: CachedValues) -> _write.Lookup:
    return lambda row, col: _file_value(cached(row, col))


def _fill(fill: ArrayFill) -> _write.Fill:
    value_at = fill.value_at
    return _write.Fill(fill.r1, fill.c1, fill.r2, fill.c2, lambda r, c: _file_value(value_at(r, c)))


def patch(
    source: str | os.PathLike[str] | BinaryIO,
    target: str | os.PathLike[str],
    *,
    cached: Mapping[str, CachedValues] | None = None,
    cells: Mapping[str, Mapping[tuple[int, int], CellEdit]] | None = None,
    fills: Mapping[str, Sequence[ArrayFill]] | None = None,
    translate: Translate | None = None,
    full_calc_on_load: bool = True,
    drop_calc_chain: bool = False,
    limits: ReadLimits | None = None,
) -> int:
    """Write ``source`` to ``target`` with changed cells, copying everything else byte for byte.

    Every part that does not change is copied as stored, compressed bytes included; only the
    worksheets with changes and, for ``full_calc_on_load``, the workbook part are rewritten, and
    within them only the changed cells. The target is written to a temporary file and moved into
    place, so ``target`` may be ``source``, and a failure leaves it as it was.

    Args:
        source: The workbook to change: a path, or a binary file open for reading and seekable.
        target: Where to write the result (a path).
        cached: Sheet name → new cached results of its formula cells.
        cells: Sheet name → (row, column), 1-based → the cell's new content.
        fills: Sheet name → ranges of constants that array formulas fill.
        translate: Moves a formula between cells; needed when an edit replaces a shared
            formula's master cell while its dependents remain.
        full_calc_on_load: Ask Excel to recalculate every formula when it opens the file
            (``calcPr fullCalcOnLoad``), since changed values leave dependent cached results
            stale. ``False`` leaves the setting as it was.
        drop_calc_chain: Remove the calculation chain, which Excel needs when a formula cell
            was removed or replaced by a constant.
        limits: What reading the source may cost.

    Returns:
        The number of cells changed.

    Raises:
        PackageError: The source is refused, or a part to rewrite is not UTF-8.
        KeyError: A sheet named in the changes does not exist.
        TypeError: A value is not a :data:`CellValue`.
    """
    edits = {
        sheet: {
            at: _write.Edit(e.formula, _file_value(e.value), e.array_ref, e.dynamic)
            for at, e in sheet_cells.items()
        }
        for sheet, sheet_cells in (cells or {}).items()
    }
    return _write.write(
        source,
        target,
        {sheet: _lookup(lookup) for sheet, lookup in (cached or {}).items()},
        edits=edits,
        fills={
            sheet: [_fill(f) for f in sheet_fills] for sheet, sheet_fills in (fills or {}).items()
        },
        translate=translate,
        drop_calc_chain=drop_calc_chain,
        full_calc_on_load=full_calc_on_load,
        limits=limits,
    )
