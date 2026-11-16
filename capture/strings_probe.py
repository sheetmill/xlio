"""What Excel reads from escaped surrogates, and from a text cell with an empty value only.

    uv run python -m capture.strings_probe

Two suites. In ``strings``, the control (openpyxl's) has one sheet whose first row this probe
writes whole. A1 holds the case's text; B1 to H1 are formulas without cached values, which Excel
computes on opening: ``LEN(A1)``, ``UNICODE(A1)``, ``UNICODE(MID(A1,2,1))``,
``UNICODE(RIGHT(A1,1))``, ``A1``, ``ISTEXT(A1)``, ``ISBLANK(A1)``. The cases:

- escapes (``_xHHHH_``): a surrogate pair, a lone high and a lone low surrogate, the pair in the
  wrong order and in lowercase hex, and an emoji written as UTF-8 (the control), each in a
  shared string and an inline string;
- ``t="str"`` without a formula: an empty ``<v/>``, no ``<v>``, a space; and Excel's own form
  with the formula ``""``.

In ``strings_cached``, the same texts are a formula's cached text. Excel computes every formula
of a workbook openpyxl wrote, whatever its ``calcId``, and keeps the cached values of one it
saved itself (xlrecalc's ``open_calc`` capture). So Excel first saves a copy of a workbook with
A1 = ``="q"``; each case plants its text as A1's cached value in that copy, with B1 to H1 kept
as Excel cached them or without values (Excel then computes them, from the cached text). The
control ``zzz`` shows that the cached text is kept.

Each row gives, from Excel's copy: A1's element, the shared strings Excel writes, and B1 to H1's
cached values. Text outside ASCII is written as Python escapes (``\\U0001f600``), and bytes
that are not UTF-8 as ``\\xNN``. Output: ``tests/captures/strings.tsv`` and
``tests/captures/strings_cached.tsv``.
"""

from __future__ import annotations

import re
import zipfile

from capture.excel import refuse_open_workbooks
from capture.probe import (
    CONTAINER,
    SHEET,
    Case,
    build,
    control_workbook,
    part_text,
    run,
    save_copy,
)

WORKBOOK = "xl/workbook.xml"
STRINGS = "xl/sharedStrings.xml"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STRINGS_TYPE = (
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
)
STRINGS_REL = (
    '<Relationship Id="rIdStrings" Type="http://schemas.openxmlformats.org/officeDocument/'
    '2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
)
CALC_OPENPYXL = '<calcPr calcId="124519" fullCalcOnLoad="1" />'
CALC_EXCEL = '<calcPr calcId="191029" />'
FORMULAS = {
    "B1": "LEN(A1)",
    "C1": "_xlfn.UNICODE(A1)",
    "D1": "_xlfn.UNICODE(MID(A1,2,1))",
    "E1": "_xlfn.UNICODE(RIGHT(A1,1))",
    "F1": "A1",
    "G1": "ISTEXT(A1)",
    "H1": "ISBLANK(A1)",
}
LABELS = ["A1", "sharedStrings", *(f"{ref} {f}" for ref, f in FORMULAS.items())]

TEXTS = {
    "pair": "_xD83D__xDE00_",
    "lone high": "a_xD83D_b",
    "lone low": "a_xDE00_b",
    "reversed pair": "_xDE00__xD83D_",
    "lowercase pair": "_xd83d__xde00_",
    "UTF-8 emoji (control)": "\U0001f600",
}
EMPTY = {
    "str <v/>, no formula": '<c r="A1" t="str"><v/></c>',
    "str without <v>, no formula": '<c r="A1" t="str"/>',
    "str <v> </v>, no formula": '<c r="A1" t="str"><v xml:space="preserve"> </v></c>',
    'str <v/>, formula "" (control)': '<c r="A1" t="str"><f>""</f><v/></c>',
}


def fill(wb: object) -> None:
    """The control as openpyxl writes it; the probe adds the shared strings part."""


def sheet(a1: str) -> str:
    formulas = "".join(f'<c r="{ref}"><f>{f}</f></c>' for ref, f in FORMULAS.items())
    row = f'<row r="1">{a1}{formulas}</row>'
    return f'<worksheet xmlns="{MAIN}"><sheetData>{row}</sheetData></worksheet>'


def sst(*texts: str) -> str:
    items = "".join(f"<si><t>{t}</t></si>" for t in texts)
    return f'<sst xmlns="{MAIN}" count="{len(texts)}" uniqueCount="{len(texts)}">{items}</sst>'


def case(name: str, a1: str, strings: tuple[str, ...] = ("x",)) -> Case:
    return Case(
        name,
        edits=(
            (WORKBOOK, CALC_OPENPYXL, CALC_EXCEL),  # changes nothing: openpyxl's file
            ("[Content_Types].xml", "</Types>", STRINGS_TYPE + "</Types>"),
            ("xl/_rels/workbook.xml.rels", "</Relationships>", STRINGS_REL + "</Relationships>"),
        ),
        parts={SHEET: sheet(a1), STRINGS: sst(*strings)},
    )


def cases() -> list[Case]:
    out = []
    for name, text in TEXTS.items():
        out.append(case(f"shared: {name}", '<c r="A1" t="s"><v>0</v></c>', (text,)))
        out.append(case(f"inline: {name}", f'<c r="A1" t="inlineStr"><is><t>{text}</t></is></c>'))
    out += [case(name, a1) for name, a1 in EMPTY.items()]
    return out


BASE_A1 = '<c r="A1" t="str"><f>"q"</f><v>q</v></c>'


def excel_base() -> bytes:
    """Excel's own copy of the control with A1 = ``="q"``: Excel keeps the cached values of a
    file it saved, where it computes every formula of one openpyxl wrote (xlrecalc's
    ``open_calc`` capture)."""
    src, copy = CONTAINER / "xlio_strings_base.xlsx", CONTAINER / "xlio_strings_base_copy.xlsx"
    CONTAINER.mkdir(parents=True, exist_ok=True)
    try:
        src.write_bytes(build(control_workbook(fill), case("base", BASE_A1)))
        copy.unlink(missing_ok=True)
        error = save_copy(src, copy)
        if error:
            raise SystemExit(f"Excel did not save the base workbook: {error}")
        return copy.read_bytes()
    finally:
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)


def cached_cases(base: bytes) -> list[Case]:
    """A1's cached text planted in Excel's copy; B1:H1 keep their values, or lose them."""
    sheet_xml = part_text(base, SHEET)
    if sheet_xml.count(BASE_A1) != 1:
        raise SystemExit(f"Excel's base does not hold {BASE_A1}")
    uncached = re.sub(r'(<c r="[B-H]1"[^>]*>)(<f>.*?</f>)<v>.*?</v>', r"\1\2", sheet_xml)
    out = []
    texts = {"zzz (control)": "zzz", **TEXTS}
    for name, text in texts.items():
        a1 = BASE_A1.replace("<v>q</v>", f"<v>{text}</v>")
        out.append(Case(f"cached: {name}", edits=((SHEET, BASE_A1, a1),)))
        out.append(Case(f"cached, B1:H1 uncached: {name}",
                        parts={SHEET: uncached.replace(BASE_A1, a1)}))  # fmt: skip
    return out


def shown(data: bytes) -> str:
    """Text as ASCII: bytes that are not UTF-8 as ``\\xNN``, other characters as escapes."""
    text = data.decode("utf-8", "backslashreplace")
    return " ".join(text.encode("ascii", "backslashreplace").decode().split()).replace("\t", " ")


def reader(copy: zipfile.ZipFile, case: Case) -> list[str]:
    names = copy.namelist()
    part = copy.read(SHEET) if SHEET in names else b""
    a1 = re.search(rb'<c r="A1"[^>]*?(?:/>|>.*?</c>)', part, re.S)
    strings = re.findall(rb"<si>.*?</si>", copy.read(STRINGS), re.S) if STRINGS in names else []
    values = [shown(a1[0]) if a1 else "-", shown(b"|".join(strings)) or "-"]
    for ref in FORMULAS:
        cell = re.search(rf'<c r="{ref}"([^>]*?)(?:/>|>(.*?)</c>)'.encode(), part, re.S)
        if cell is None:
            values.append("-")
            continue
        kind = re.search(rb't="([^"]*)"', cell[1])
        value = re.search(rb"<v\b[^>]*>(.*?)</v>", cell[2] or b"", re.S)
        prefix = (kind[1] + b":") if kind else b"n:"
        values.append(shown(prefix + (value[1] if value else b"(none)")))
    return values


if __name__ == "__main__":
    run(
        "strings",
        control_workbook(fill),
        cases(),
        LABELS,
        "AppleScript open, save as .xlsx, read the copy's XML",
        "A1 = the case's text; B1:H1 computed by Excel on opening; text outside ASCII as escapes",
        reader=reader,
    )
    refuse_open_workbooks()
    base = excel_base()
    run(
        "strings_cached",
        base,
        cached_cases(base),
        LABELS,
        'Excel saves a copy of a workbook with A1 = "q" (formula); A1\'s cached text in that copy '
        "replaced; AppleScript open, save as .xlsx, read the copy's XML",
        "A1's cached text planted in Excel's own file; B1:H1 kept as Excel cached them, or "
        "without values (computed on opening); text outside ASCII as escapes",
        reader=reader,
    )
