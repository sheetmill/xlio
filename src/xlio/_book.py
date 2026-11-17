"""The workbook: sheets, defined names, the date system, calculation settings."""

from __future__ import annotations

import os
import posixpath
import xml.parsers.expat
from collections.abc import Iterator
from dataclasses import dataclass
from types import TracebackType
from typing import Any, BinaryIO

from xlio._cells import Cell, read_cells, read_shared_strings
from xlio._package import Package, PackageError, ReadLimits
from xlio._rich import RichString, read_rich_strings
from xlio._types import bounded_int

__all__ = ["Book", "CalcProperties", "DefinedName", "SheetInfo"]

_MAINS = (
    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}",
    "{http://purl.oclc.org/ooxml/spreadsheetml/main}",
)
_RELS = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}",
    "{http://purl.oclc.org/ooxml/officeDocument/relationships}",
)
_TRUE = ("1", "true")
# Rich-value error types and the error Excel shows for each (capture errors). Another type is
# shown as the #VALUE! the cell caches; 7 and 17 are refused.
_RICH_ERRORS = {
    "3": "#REF!", "4": "#NAME?", "6": "#N/A", "8": "#SPILL!", "9": "#CONNECT!",
    "10": "#BLOCKED!", "11": "#UNKNOWN!", "12": "#FIELD!", "13": "#CALC!",
    "15": "#BUSY!", "18": "#UNKNOWN!", "19": "#TIMEOUT!",
}  # fmt: skip
_REFUSED_ERROR_TYPES = frozenset({"7", "17"})
# errorType 14: Excel saves the cell as the number 0 (its rich value dropped) and a formula
# referring to it caches 0, so it is read as 0; at load-time calculation ERROR.TYPE gave 15 and
# VALUETOTEXT "#CALC!", so what Excel shows in the cell itself is still to be captured.
_RICH_NUMBERS = {"14": 0.0}


CELL_KINDS = frozenset({"worksheet", "xlMacrosheet", "xlIntlMacrosheet"})
"""The relationship types of sheets that hold cells: worksheets and macro sheets."""


@dataclass(frozen=True, slots=True)
class SheetInfo:
    """A sheet: name, part, ``visible``/``hidden``/``veryHidden``, and its kind.

    ``kind`` is the last segment of the relationship type: ``worksheet``, ``chartsheet``,
    ``dialogsheet``, ``xlMacrosheet`` or ``xlIntlMacrosheet``; worksheets and macro sheets have
    cells to read (:data:`CELL_KINDS`).
    """

    name: str
    part: str
    state: str
    kind: str
    position: int = 0
    """Its place in ``<sheets>`` as written, 0-based: what ``localSheetId`` and ``activeTab``
    count, a sheet without a part included."""


@dataclass(frozen=True, slots=True)
class DefinedName:
    """A defined name: its text (a formula, file form), scope, hidden.

    ``sheet`` is ``localSheetId`` as written: the sheet's place in ``<sheets>``, a sheet without a
    part counted (``Sheet.position``), or ``None`` for a name of the workbook.
    """

    name: str
    text: str
    sheet: int | None
    hidden: bool


@dataclass(frozen=True, slots=True)
class CalcProperties:
    """``<calcPr>`` as written; ``None`` where the file leaves an attribute out."""

    calc_id: int | None = None
    full_calc_on_load: bool = False
    calc_mode: str | None = None
    iterate: bool = False
    iterate_count: int | None = None
    iterate_delta: float | None = None


def _find(element: object, local: str) -> list[object]:
    return [e for ns in _MAINS for e in element.iter(ns + local)]  # type: ignore[attr-defined]


def _local(element: object, local: str) -> list[Any]:
    """Descendants with this local name, in any namespace."""
    return [
        e
        for e in element.iter()  # type: ignore[attr-defined]
        if isinstance(e.tag, str) and e.tag.rpartition("}")[2] == local
    ]


def _attr(element: object, local: str) -> str | None:
    for ns in _RELS:
        value = element.get(ns + local)  # type: ignore[attr-defined]
        if value is not None:
            return str(value)
    return None


class Book:
    """An xlsx workbook opened for reading. Use as a context manager, or ``close`` it."""

    def __init__(
        self,
        source: str | os.PathLike[str] | BinaryIO,
        limits: ReadLimits | None = None,
    ) -> None:
        self._package = Package(source, limits or ReadLimits())
        try:
            self._load()
        except BaseException:
            self._package.close()
            raise
        self._strings: list[str] | None = None
        self._rich: list[RichString] | None = None

    def __enter__(self) -> Book:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        """Close the file."""
        self._package.close()

    @property
    def package(self) -> Package:
        """The package, for the writer."""
        return self._package

    def _load(self) -> None:
        package = self._package
        root_rels = package.relationships("")
        workbook = next(
            (
                t
                for typ, t, ext in root_rels.values()
                if typ.endswith("/officeDocument") and not ext
            ),
            None,
        )
        if workbook is None or workbook not in package.names:
            raise PackageError("the package has no workbook part")
        self.workbook_part = workbook
        rels = package.relationships(workbook)
        root = package.xml(workbook)
        self.root = root
        """The workbook part's element tree."""
        self.sheets: list[SheetInfo] = []
        for position, sheet in enumerate(_find(root, "sheet")):
            rel = rels.get(_attr(sheet, "id") or "")
            if rel is None:
                continue
            kind = rel[0].rsplit("/", 1)[-1]
            self.sheets.append(
                SheetInfo(
                    sheet.get("name", ""),  # type: ignore[attr-defined]
                    rel[1],
                    sheet.get("state", "visible"),  # type: ignore[attr-defined]
                    kind,
                    position,
                )
            )
        views = _find(root, "workbookView")
        tab = views[0].get("activeTab", "0") if views else "0"  # type: ignore[attr-defined]
        self.active_tab = bounded_int(tab) if tab.isdigit() and tab.isascii() else 0
        """The sheet the workbook opens on (``workbookView activeTab``), 0-based."""
        props = _find(root, "workbookPr")
        self.date1904 = bool(props) and props[0].get("date1904", "0") in _TRUE  # type: ignore[attr-defined]
        calc = _find(root, "calcPr")
        self.calc = _calc(calc[0]) if calc else CalcProperties()
        self.defined_names: list[DefinedName] = []
        for dn in _find(root, "definedName"):
            local = dn.get("localSheetId")  # type: ignore[attr-defined]
            self.defined_names.append(
                DefinedName(
                    dn.get("name", ""),  # type: ignore[attr-defined]
                    dn.text or "",  # type: ignore[attr-defined]
                    _integer(local, "localSheetId") if local is not None else None,
                    dn.get("hidden", "0") in _TRUE,  # type: ignore[attr-defined]
                )
            )
        self._strings_part = next(
            (t for typ, t, ext in rels.values() if typ.endswith("/sharedStrings") and not ext),
            None,
        )
        self.dynamic_cm: frozenset[int] = frozenset()
        """The ``cm`` values that mark a dynamic-array formula (metadata type ``XLDAPR``)."""
        self.metadata_counts = (0, 0)  # the cellMetadata and valueMetadata entries
        self.value_rich: dict[int, int] = {}  # vm -> the rich value's index (XLRICHVALUE)
        self.value_errors: dict[int, str | float] = {}
        """``vm`` value -> what a rich value stands for, for a cell that caches ``#VALUE!``: an
        error (``#SPILL!``), or a number (errorType 14)."""
        meta = next(
            (t for typ, t, ext in rels.values() if typ.endswith("/sheetMetadata") and not ext),
            None,
        )
        if meta is not None and meta in package.names:
            self._metadata(meta, rels)
        self.external_links = [
            t for typ, t, ext in rels.values() if typ.endswith("/externalLink") and not ext
        ]

    def _metadata(self, part: str, rels: dict[str, tuple[str, str, bool]]) -> None:
        """The dynamic-array and rich-value metadata.

        As Excel writes it (captured by xlrecalc: https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/arrays.tsv).
        """
        root = self._package.xml(part)
        types = [t.get("name", "") for t in _local(root, "metadataType")]
        future: dict[str, list[object]] = {}
        for fm in _local(root, "futureMetadata"):
            future[fm.get("name", "")] = list(_local(fm, "bk"))

        def entries(block: str) -> list[tuple[str, int]]:
            out = []
            for container in _local(root, block):
                for bk in _local(container, "bk"):
                    rc = next(iter(_local(bk, "rc")), None)
                    if rc is None:
                        out.append(("", 0))
                        continue
                    t = _int_or(rc.get("t", "0"), 0)
                    v = _int_or(rc.get("v", "0"), -1)
                    out.append((types[t - 1] if 0 < t <= len(types) else "", v))
            return out

        cell_entries, value_entries = entries("cellMetadata"), entries("valueMetadata")
        self.metadata_counts = (len(cell_entries), len(value_entries))
        dynamic = set()
        for k, (name, v) in enumerate(cell_entries, 1):
            blocks = future.get(name, [])
            if name == "XLDAPR" and 0 <= v < len(blocks):
                props = next(iter(_local(blocks[v], "dynamicArrayProperties")), None)
                if props is not None and props.get("fDynamic", "0") in _TRUE:
                    dynamic.add(k)
        self.dynamic_cm = frozenset(dynamic)
        values = self._rich_errors(rels)
        for k, (name, v) in enumerate(value_entries, 1):
            blocks = future.get(name, [])
            if name == "XLRICHVALUE" and 0 <= v < len(blocks):
                rvb = next(iter(_local(blocks[v], "rvb")), None)
                if rvb is not None:
                    index = _int_or(rvb.get("i", "-1"), -1)
                    if index >= 0:
                        self.value_rich[k] = index
                    error = values.get(index)
                    if error is not None:
                        self.value_errors[k] = error

    def _rich_errors(self, rels: dict[str, tuple[str, str, bool]]) -> dict[int, str | float]:
        """Rich value index -> the error Excel shows for it, for the error types captured."""
        parts = {typ.rsplit("/", 1)[-1].lower(): t for typ, t, ext in rels.values() if not ext}
        values_part = parts.get("rdrichvalue")
        structures_part = parts.get("rdrichvaluestructure")
        if values_part not in self._package.names or structures_part not in self._package.names:
            return {}
        assert values_part is not None
        assert structures_part is not None
        structures = []
        for st in _local(self._package.xml(structures_part), "s"):
            keys = [k.get("n", "") for k in _local(st, "k")]
            structures.append((st.get("t", ""), keys))
        out = {}
        for i, rv in enumerate(_local(self._package.xml(values_part), "rv")):
            s = _int_or(rv.get("s", "-1"), -1)
            if not 0 <= s < len(structures) or structures[s][0] != "_error":
                continue
            fields = dict(zip(structures[s][1], (v.text for v in _local(rv, "v")), strict=False))
            error_type = (fields.get("errorType") or "").strip()
            if error_type in _REFUSED_ERROR_TYPES:
                raise PackageError(f"{values_part}: a rich value of errorType {error_type}, "
                                   "which Excel refuses")  # fmt: skip
            code: str | float | None = _RICH_ERRORS.get(error_type, _RICH_NUMBERS.get(error_type))
            if code is not None:
                out[i] = code
        return out

    @property
    def shared_strings(self) -> list[str]:
        """The shared string table, read once."""
        if self._strings is None:
            part = self._strings_part
            if part is None or part not in self._package.names:
                self._strings = []
            else:
                with self._package.part(part) as f:
                    try:
                        self._strings = read_shared_strings(f)
                    except xml.parsers.expat.ExpatError as exc:
                        raise PackageError(f"{part}: {exc}") from exc
        return self._strings

    @property
    def rich_strings(self) -> list[RichString]:
        """The shared string table with runs and phonetic guides, read once."""
        if self._rich is None:
            part = self._strings_part
            if part is None or part not in self._package.names:
                self._rich = []
            else:
                with self._package.part(part) as f:
                    self._rich = read_rich_strings(f, part)
        return self._rich

    def sheet(self, name: str) -> SheetInfo:
        """The sheet with this name (Excel compares sheet names without case)."""
        folded = name.casefold()
        for sheet in self.sheets:
            if sheet.name.casefold() == folded:
                return sheet
        raise KeyError(f"no sheet named {name!r}")

    def cells(
        self,
        sheet: str | SheetInfo,
        rich_strings: list[RichString] | None = None,
        **options: Any,
    ) -> Iterator[Cell]:
        """A worksheet's or macro sheet's cells, in file order; ``options`` go to ``read_cells``."""
        info = sheet if isinstance(sheet, SheetInfo) else self.sheet(sheet)
        if info.kind not in CELL_KINDS:
            raise ValueError(f"{info.name!r} is a {info.kind}, not a worksheet")
        strings = self.shared_strings
        with self._package.part(info.part) as f:
            try:
                yield from read_cells(f, strings, rich_strings, **options)
            except xml.parsers.expat.ExpatError as exc:
                raise PackageError(f"{info.part}: {exc}") from exc

    def part_name(self, relative: str) -> str:
        """A part name relative to the workbook part's folder."""
        return posixpath.normpath(posixpath.join(posixpath.dirname(self.workbook_part), relative))


def _int_or(text: str, default: int) -> int:
    """An integer attribute the reader can do without: its default when it is not one."""
    try:
        return int(text)
    except ValueError:
        return default


def _integer(text: str, what: str) -> int:
    try:
        return int(text)
    except ValueError:
        raise PackageError(f"{what}: not an integer: {text[:40]!r}") from None


def _calc(element: object) -> CalcProperties:
    """The calculation settings; a number that is not one is taken as absent (none computes)."""
    get = element.get  # type: ignore[attr-defined]
    calc_id = get("calcId")
    count = get("iterateCount")
    delta = get("iterateDelta")
    try:
        iterate_delta = float(delta) if delta is not None else None
    except ValueError:
        iterate_delta = None
    return CalcProperties(
        calc_id=_int_or(calc_id, -1) if calc_id is not None else None,
        full_calc_on_load=get("fullCalcOnLoad", "0") in _TRUE,
        calc_mode=get("calcMode"),
        iterate=get("iterate", "0") in _TRUE,
        iterate_count=_int_or(count, -1) if count is not None else None,
        iterate_delta=iterate_delta,
    )
