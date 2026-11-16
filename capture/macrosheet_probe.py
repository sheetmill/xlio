"""What Excel keeps of an Excel 4.0 macro sheet it opens and saves.

    uv run python -m capture.macrosheet_probe --open          # opens the cases as a person does
    uv run python -m capture.macrosheet_probe --save ANSWER [REPORT]  # saves each copy, reads it

A workbook with a macro sheet may bring up Excel's notice about macros, which is answered by a
person. So the cases are opened through Launch Services (``open -a``), as a person opens a file,
and the script sends Excel nothing more. Answer each notice with **Disable Macros** (nothing is
run either way: the sheets hold no macro), then run ``--save`` with the answer you gave
(``disable``, or ``none`` if no notice appeared), followed by what else you saw (other dialogs,
in their words), which goes into the TSV's header. Each workbook still open is then saved as an
``.xlsm`` copy, closed, and read.

The control (openpyxl's, A1 = 1 on the worksheet) gains a second sheet, a macro sheet with a
column width, a number, an inline string, a formula over each, and page margins:

- ``.xlsm``, relationship ``xlMacrosheet``, root ``xm:macrosheet``;
- ``.xlsm``, relationship ``xlIntlMacrosheet`` (an international macro sheet);
- ``.xlsx`` (the workbook's content type not macro-enabled), the same macro sheet;
- ``.xlsm``, the root ``macrosheet`` in SpreadsheetML's namespace instead of ``xm``.

Each row gives, from Excel's copy: the workbook part's content type, the ``<sheet>`` elements,
the macro sheet's relationship type and target, its content type, its root element, ``cols``,
``sheetData`` and ``pageMargins``. Output: ``tests/captures/macrosheet.tsv``.
"""

from __future__ import annotations

import datetime
import re
import subprocess
import sys
import zipfile
from pathlib import Path

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.probe import (
    CAPTURES,
    CONTAINER,
    HERE,
    ROOT,
    TIME_LIMIT,
    WIDTH,
    Case,
    build,
    control_workbook,
    excel_version,
)

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MS_REL = "http://schemas.microsoft.com/office/2006/relationships/"
WORKBOOK = "xl/workbook.xml"
TYPES = "[Content_Types].xml"
RELS = "xl/_rels/workbook.xml.rels"
PART = "xl/macrosheets/sheet1.xml"
WORKBOOK_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
MACRO_WORKBOOK_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
BODY = (
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    '<cols><col min="1" max="1" width="20" customWidth="1"/></cols>'
    "<sheetData>"
    '<row r="1"><c r="A1"><v>1</v></c></row>'
    '<row r="2"><c r="A2" t="inlineStr"><is><t>text</t></is></c></row>'
    '<row r="3"><c r="A3"><f>A1+1</f><v>2</v></c></row>'
    '<row r="4"><c r="A4" t="str"><f>A2&amp;"!"</f><v>text!</v></c></row>'
    "</sheetData>"
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
)
LABELS = ["content type", "sheets", "relationship", "part type", "root", "cols", "sheetData",
          "pageMargins"]  # fmt: skip


def macro_sheet(root: str) -> str:
    if root == "xm":
        return (
            f'<xm:macrosheet xmlns="{MAIN}" xmlns:r="{REL}" xmlns:xm="{XM}">{BODY}</xm:macrosheet>'
        )
    return f'<macrosheet xmlns="{MAIN}" xmlns:r="{REL}">{BODY}</macrosheet>'


def case(name: str, rel: str, part_type: str, root: str, macro_enabled: bool) -> Case:
    edits = [
        (WORKBOOK, "</sheets>", '<sheet name="Macro1" sheetId="2" r:id="rIdMacro" /></sheets>'),
        (RELS, "</Relationships>",
         f'<Relationship Id="rIdMacro" Type="{MS_REL}{rel}" Target="macrosheets/sheet1.xml"/>'
         "</Relationships>"),
        (TYPES, "</Types>",
         f'<Override PartName="/{PART}" ContentType="{part_type}"/></Types>'),
    ]  # fmt: skip
    if macro_enabled:
        edits.append((TYPES, WORKBOOK_TYPE, MACRO_WORKBOOK_TYPE))
    return Case(name, edits=tuple(edits), parts={PART: macro_sheet(root)})


CASES: list[tuple[str, Case]] = [  # (the file's suffix, the case)
    ("xlsm", case("xlsm, xlMacrosheet", "xlMacrosheet",
                  "application/vnd.ms-excel.macrosheet+xml", "xm", True)),
    ("xlsm", case("xlsm, xlIntlMacrosheet", "xlIntlMacrosheet",
                  "application/vnd.ms-excel.intlmacrosheet+xml", "xm", True)),
    ("xlsx", case("xlsx, xlMacrosheet", "xlMacrosheet",
                  "application/vnd.ms-excel.macrosheet+xml", "xm", False)),
    ("xlsm", case("xlsm, root in SpreadsheetML's namespace", "xlMacrosheet",
                  "application/vnd.ms-excel.macrosheet+xml", "main", True)),
]  # fmt: skip


def fill(wb: object) -> None:
    wb.active["A1"] = 1  # type: ignore[attr-defined]


def source(i: int, suffix: str) -> Path:
    return CONTAINER / f"xlio_macro_{i}.{suffix}"


def open_cases() -> None:
    """Write the cases into Excel's container and open each as a person does; no more."""
    refuse_open_workbooks()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    control = control_workbook(fill)
    for i, (suffix, c) in enumerate(CASES):
        src = source(i, suffix)
        src.write_bytes(build(control, c))
        subprocess.run(["open", "-a", "Microsoft Excel", str(src)], check=True, timeout=30)
        print(f"{c.name}: opened {src.name}; answer any notice in Excel", flush=True)


def first(pattern: str, text: str) -> str:
    found = re.search(pattern, text, re.S)
    if found is None:
        return "-"
    value = found.group(1) if found.groups() else found.group(0)
    return " ".join(value.split())[:WIDTH].replace("\t", " ")


TYPE = r'Type="([^"]*)"'


def reader(z: zipfile.ZipFile) -> list[str]:
    types = z.read(TYPES).decode()
    rels = z.read(RELS).decode()
    rel = re.search(r'<Relationship [^>]*Type="[^"]*[Mm]acrosheet"[^>]*>', rels)
    target = re.search(r'Target="([^"]*)"', rel[0]) if rel else None
    part = ""
    if target is not None:  # relative to xl/, or absolute from the package root
        part = target[1][1:] if target[1].startswith("/") else "xl/" + target[1]
    body = z.read(part).decode() if part in z.namelist() else ""
    return [
        first(r'PartName="/xl/workbook.xml" ContentType="([^"]*)"', types),
        "|".join(re.findall(r"<sheet [^>]*>", z.read(WORKBOOK).decode())),
        f"{first(TYPE, rel[0])} -> {target[1]}" if rel and target else "absent",
        first(rf'PartName="/{re.escape(part)}" ContentType="([^"]*)"', types) if part else "-",
        first(r"<(?:\w+:)?macrosheet\b[^>]*>", body) if body else "absent",
        first(r"<cols>.*?</cols>", body),
        first(r"<sheetData>.*?</sheetData>", body),
        first(r"<pageMargins\b[^>]*>", body),
    ]


def save(answer: str, report: str) -> None:
    version = excel_version()
    rows = []
    for i, (suffix, c) in enumerate(CASES):
        src = source(i, suffix)
        copy = CONTAINER / f"xlio_macro_{i}_copy.xlsm"
        try:
            copy.unlink(missing_ok=True)
            result = subprocess.run(
                ["gtimeout", "-k", "5", TIME_LIMIT, "osascript",
                 str(HERE / "save_open.applescript"), src.name, str(copy), "xlsm"],
                capture_output=True, text=True, check=False,
            )  # fmt: skip
            if result.returncode != 0:
                # AppleScript's message starts with the script's path: keep Excel's words only
                error = result.stderr.strip().rsplit("error: ", 1)[-1].replace("\t", " ")[:120]
                rows.append("\t".join([c.name, "not saved", error] + [""] * len(LABELS)))
            else:
                with zipfile.ZipFile(copy) as z:
                    rows.append("\t".join([c.name, "opened", "", *reader(z)]))
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print(rows[-1], flush=True)
    header = [
        "# suite: macrosheet (capture); Excel " + version + " for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: opened through Launch Services (open -a) "
        f"as a person opens a file, the macro notice answered '{answer}' by the person (none: "
        "not shown), then saved as an .xlsm copy through AppleScript",
        "# what Excel's copy keeps of the macro sheet; not saved = the workbook was not open "
        "(Excel did not open it)" + (f"; {report}" if report else ""),
        "\t".join(["case", "outcome", "error", *LABELS]),
    ]
    out = CAPTURES / "macrosheet.tsv"
    out.write_text("\n".join(header + rows) + "\n", encoding="utf-8")
    print(f"wrote {out.relative_to(ROOT)}: {len(rows)} cases")


if __name__ == "__main__":
    args = sys.argv[1:]
    if args == ["--open"]:
        open_cases()
    elif args[:1] == ["--save"] and args[1:2] in (["disable"], ["enable"], ["none"]):
        save(args[1], " ".join(args[2:]))
    else:
        raise SystemExit("--open, then --save disable|none [what you saw, for the header]")
