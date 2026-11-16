"""How Excel reads and writes its errors, above all the newer ones (``#FIELD!``, ``#BUSY!``).

    uv run python -m capture.errors_probe            # the errors suite
    uv run python -m capture.errors_probe --stale    # the errors_stale suite

The control (openpyxl's): A1 = 1, B1 = ``=A1``, C1 = ``=ERROR.TYPE(A1)``, D1 = ``=ISERROR(A1)``,
E1 = ``=VALUETOTEXT(A1)`` (the error's text),
A2 = 5, A3 = 6; ``fullCalcOnLoad``, so Excel computes B1:D1 on open. Four groups of cases:

- ``plain <code>``: A1 is ``<c t="e"><v>code</v></c>``, every code Excel defines and two it does
  not; does Excel open it, and which error does it read (B1, C1)?
- ``rich <n>``: A1 is ``#VALUE!`` with a rich value of type ``_error`` and ``errorType`` n (the
  parts Excel writes for a blocked spill; 8 is ``#SPILL!`` and 13 ``#CALC!``, both captured, as
  controls): which error each ``errorType`` stands for.
- ``cached plain <code>``: A1 is a formula whose cached value is the code, written plainly; does
  Excel open it (a writer must not write a code that way if not)?
- ``formula ...``: A1 holds a formula whose result is an error; how Excel writes it (plain, or
  ``#VALUE!`` with a rich value, and its ``errorType``).

Each row gives, from Excel's copy: A1 as type:value, A1's ``vm`` and the ``errorType`` and
``subType`` it leads to, B1, C1 and D1. Output: ``tests/captures/errors.tsv``.

``errors_stale``: what Excel shows for a formula cell whose ``vm`` leads to an error rich value
while its ``<v>`` caches something else, as a writer that drops nothing leaves it. Excel keeps
the cached values of a file it saved itself (xlrecalc's ``open_calc`` capture,
https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/open_calc.tsv),
so Excel first saves a copy of the control with A1 = ``FILTER(A2:A3,A2:A3>100)`` (no match:
``#CALC!``, written as ``#VALUE!`` with a rich value); each case then plants a cached value in
A1 of that copy (5, TRUE, text, ``#DIV/0!``; and ``#VALUE!`` without the ``vm``, the
control), with B1:E1 kept as Excel cached them or without values (Excel then computes them from
A1 as it holds it). Output: ``tests/captures/errors_stale.tsv``.
"""

from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
import zipfile

from capture.excel import refuse_open_workbooks
from capture.probe import (
    CONTAINER,
    SHEET,
    Case,
    build,
    control_workbook,
    part_text,
    read,
    run,
    save_copy,
)

A1 = '<c r="A1" t="n"><v>1</v></c>'
DEFINED = [
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#CONNECT!",
    "#BUSY!",
    "#EXTERNAL!",
    "#UNKNOWN!",
    "#PYTHON!",
    "#TIMEOUT!",
]
UNDEFINED = ["#UNHEARD!", "#field!"]
ERROR_TYPES = range(0, 21)
FORMULAS = {
    "formula 1/0": "1/0",
    "formula FIELDVALUE of a number": '_xlfn.FIELDVALUE(A2,"Population")',
    "formula FILTER with no match": "_xlfn._xlws.FILTER(A2:A3,A2:A3&gt;100)",
    "formula ERROR.TYPE of FIELDVALUE": 'ERROR.TYPE(_xlfn.FIELDVALUE(A2,"Population"))',
}

RD = "http://schemas.microsoft.com/office/2017/06/relationships/"
CONTENT = "application/vnd."
METADATA = (
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
    '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" '
    'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
    'clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>'
    '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst>'
    '<ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk>'
    '</futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>'
    "</metadata>"
)
VALUES = (
    '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">'
    '<rv s="0"><v>{n}</v><v>0</v></rv></rvData>'
)
STRUCTURES = (
    '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" '
    'count="1"><s t="_error"><k n="errorType" t="i"/><k n="subType" t="i"/></s></rvStructures>'
)
RELS_END = "</Relationships>"
RICH_RELS = (
    '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    'sheetMetadata" Target="metadata.xml" Id="rId90" />'
    f'<Relationship Type="{RD}rdRichValue" Target="richData/rdrichvalue.xml" Id="rId91" />'
    f'<Relationship Type="{RD}rdRichValueStructure" Target="richData/rdrichvaluestructure.xml" '
    'Id="rId92" />' + RELS_END
)
TYPES_END = "</Types>"
RICH_TYPES = (
    '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.spreadsheetml.sheetMetadata+xml" />'
    f'<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="{CONTENT}ms-excel.'
    'rdrichvalue+xml" />'
    '<Override PartName="/xl/richData/rdrichvaluestructure.xml" '
    f'ContentType="{CONTENT}ms-excel.rdrichvaluestructure+xml" />' + TYPES_END
)


def fill(wb: object) -> None:
    ws = wb.active  # type: ignore[attr-defined]
    ws["A1"] = 1
    ws["B1"] = "=A1"
    ws["C1"] = "=ERROR.TYPE(A1)"
    ws["D1"] = "=ISERROR(A1)"
    ws["E1"] = "=_xlfn.VALUETOTEXT(A1)"
    ws["A2"] = 5
    ws["A3"] = 6


def cases() -> list[Case]:
    out = [Case("control")]
    for code in DEFINED + UNDEFINED:
        out.append(Case(f"plain {code}", ((SHEET, A1, f'<c r="A1" t="e"><v>{code}</v></c>'),)))
    for n in ERROR_TYPES:
        out.append(
            Case(
                f"rich {n}",
                (
                    (SHEET, A1, '<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>'),
                    ("xl/_rels/workbook.xml.rels", RELS_END, RICH_RELS),
                    ("[Content_Types].xml", TYPES_END, RICH_TYPES),
                ),
                {
                    "xl/metadata.xml": METADATA,
                    "xl/richData/rdrichvalue.xml": VALUES.format(n=n),
                    "xl/richData/rdrichvaluestructure.xml": STRUCTURES,
                },
            )
        )
    for code in [*DEFINED[8:], "#GETTING_DATA"]:
        out.append(
            Case(
                f"cached plain {code}",
                ((SHEET, A1, f'<c r="A1" t="e"><f>1/0</f><v>{code}</v></c>'),),
            )
        )
    for name, formula in FORMULAS.items():
        out.append(Case(name, ((SHEET, A1, f'<c r="A1"><f>{formula}</f><v /></c>'),)))
    return out


def cell(z: zipfile.ZipFile, ref: str) -> str:
    found = read(z, SHEET, rf'<c r="{ref}"[^>]*?(?:/>|>.*?</c>)')
    if found in ("absent", "-"):
        return "absent"
    t = ET.fromstring(found.replace("<c ", '<c xmlns="m" ', 1))
    v = t.find("{m}v")
    return f"{t.get('t', 'n')}:{v.text if v is not None and v.text else ''}"


def rich_error(z: zipfile.ZipFile, vm: str | None) -> str:
    """The ``errorType`` and ``subType`` a ``vm`` leads to, through Excel's own parts."""
    if vm is None:
        return "-"
    try:
        main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        rd = "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}"
        meta = ET.fromstring(z.read("xl/metadata.xml"))
        types = [t.get("name") for t in meta.iter(main + "metadataType")]
        rc = next(iter(meta.iter(main + "valueMetadata"))).findall(f"{main}bk")[int(vm) - 1][0]
        kind = types[int(rc.get("t", "0")) - 1]
        future = next(f for f in meta.iter(main + "futureMetadata") if f.get("name") == kind)
        rvb = list(future.iter(rd + "rvb"))[int(rc.get("v", "0"))]
        names = [n for n in z.namelist() if n.lower().endswith("rdrichvalue.xml")]
        snames = [n for n in z.namelist() if n.lower().endswith("rdrichvaluestructure.xml")]
        rv = ET.fromstring(z.read(names[0]))[int(rvb.get("i", "0"))]
        s = ET.fromstring(z.read(snames[0]))[int(rv.get("s", "0"))]
        keys = [k.get("n") for k in s]
        values = dict(zip(keys, (v.text for v in rv), strict=False))
        return f"{s.get('t')} " + " ".join(f"{k}={values[k]}" for k in keys)
    except (KeyError, IndexError, StopIteration, ValueError, ET.ParseError) as exc:
        return f"unresolved ({type(exc).__name__})"


def reader(z: zipfile.ZipFile, case: Case) -> list[str]:
    a1 = read(z, SHEET, r'<c r="A1"([^>]*)')
    vm = next((p.split('"')[1] for p in a1.split() if p.startswith("vm=")), None)
    return [
        cell(z, "A1"),
        vm or "-",
        rich_error(z, vm),
        cell(z, "B1"),
        cell(z, "C1"),
        cell(z, "D1"),
        cell(z, "E1"),
    ]


LABELS = [
    "A1",
    "A1 vm",
    "A1 rich value",
    "B1 =A1",
    "C1 =ERROR.TYPE(A1)",
    "D1 =ISERROR(A1)",
    "E1 =VALUETOTEXT(A1)",
]
STALE_FORMULA = FORMULAS["formula FILTER with no match"]
STALE = {  # case -> (t attribute, cached value) planted in A1
    "#VALUE! without vm (control)": ("e", "#VALUE!"),
    "cached 5": ("n", "5"),
    "cached TRUE": ("b", "1"),
    "cached text": ("str", "x"),
    "cached #DIV/0!": ("e", "#DIV/0!"),
}


def stale_base() -> bytes:
    """Excel's own copy of the control with A1 = the FILTER formula that matches nothing."""
    base = Case("base", ((SHEET, A1, f'<c r="A1"><f>{STALE_FORMULA}</f><v /></c>'),))
    src, copy = CONTAINER / "xlio_errors_base.xlsx", CONTAINER / "xlio_errors_base_copy.xlsx"
    CONTAINER.mkdir(parents=True, exist_ok=True)
    try:
        src.write_bytes(build(control_workbook(fill), base))
        copy.unlink(missing_ok=True)
        error = save_copy(src, copy)
        if error:
            raise SystemExit(f"Excel did not save the base workbook: {error}")
        return copy.read_bytes()
    finally:
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)


def stale_cases(base: bytes) -> list[Case]:
    sheet = part_text(base, SHEET)
    found = re.search(r'<c r="A1"[^>]*\bvm="\d+"[^>]*>.*?</c>', sheet, re.S)
    if found is None:
        raise SystemExit("Excel's base has no A1 with a vm")
    a1 = found[0]
    uncached = re.sub(r'(<c r="[B-E]1"[^>]*>)(<f>.*?</f>)<v>.*?</v>', r"\1\2", sheet)
    out = [Case("base (Excel's own)"), Case("base, B1:E1 uncached", parts={SHEET: uncached})]
    for name, (kind, value) in STALE.items():
        planted = re.sub(r'\st="[^"]*"', "", a1, count=1)
        if kind != "n":
            planted = planted.replace('<c r="A1"', f'<c r="A1" t="{kind}"', 1)
        planted = re.sub(r"<v>.*?</v>", f"<v>{value}</v>", planted, count=1)
        if name.endswith("(control)"):
            planted = re.sub(r'\svm="\d+"', "", planted, count=1)
        out.append(Case(name, ((SHEET, a1, planted),)))
        out.append(Case(f"{name}, B1:E1 uncached", parts={SHEET: uncached.replace(a1, planted)}))
    return out


def stale() -> None:
    refuse_open_workbooks()
    base = stale_base()
    run(
        "errors_stale",
        base,
        stale_cases(base),
        LABELS,
        "Excel saves a copy of openpyxl's workbook with A1 = FILTER(...) matching nothing (#CALC!, "
        "a rich value); A1's cached value in that copy replaced, its vm kept; opened with alerts "
        "off and saved as a copy",
        "cells as type:value from Excel's copy; a rich value as its structure type and keys; "
        "B1:E1 as Excel cached them, or computed on opening (uncached)",
        reader,
    )


def main() -> None:
    run(
        "errors",
        control_workbook(fill),
        cases(),
        LABELS,
        "openpyxl's workbook with A1 rewritten (and rich-value parts added) in the XML, opened "
        "with alerts off and saved as a copy",
        "cells as type:value from Excel's copy; a rich value as its structure type and keys",
        reader,
    )


if __name__ == "__main__":
    if sys.argv[1:] == ["--stale"]:
        stale()
    else:
        main()
