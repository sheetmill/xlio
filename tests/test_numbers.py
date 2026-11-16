"""Numbers in attributes that no file means: thousands of digits, or a double near its limit.

Python refuses to convert a string of thousands of digits to an integer, and a double near its
largest value overflows once scaled. Each case is a workbook Excel saved with every value of one
numeric attribute in one part replaced; reading everything xlio reads, and ``check()``, must
give a value or refuse the file (``PackageError``, ``LimitError``), never another exception.
Each workbook's cases run in a child process under the memory watchdog.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest

import xlio
from tests.test_read import MAIN, REL, package
from tests.watchdog import bounded

FILES = Path(__file__).parent / "files" / "excel"
NUMERIC = re.compile(rb'\s([A-Za-z:]+)="(-?[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)"')
HUGE = ("9" * 5000, "-" + "9" * 5000, "1" + "0" * 5000, "1e308", "-1.7976931348623157e308")


def cases(data: bytes) -> list[tuple[str, str, str]]:
    """(part, attribute, value) for each numeric attribute of each XML part."""
    out = []
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for part in z.namelist():
            if not part.endswith((".xml", ".rels", ".vml")):
                continue
            names = {m[1].decode() for m in NUMERIC.finditer(z.read(part))}
            out += [(part, name, value) for name in sorted(names) for value in HUGE]
    return out


def rewritten(data: bytes, part: str, attribute: str, value: str) -> bytes:
    out = io.BytesIO()
    pattern = re.compile(rb"(\s" + re.escape(attribute.encode()) + rb'=")[^"]*(")')
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            body = src.read(info)
            if info.filename == part:
                body = pattern.sub(lambda m: m[1] + value.encode() + m[2], body)
            dst.writestr(info, body)
    return out.getvalue()


def read_everything(data: bytes) -> None:
    wb = xlio.open(io.BytesIO(data))
    book = ("styles", "theme", "views", "properties", "rich_values", "persons", "external_links",
            "cell_metadata", "calc", "defined_names", "unparsed")  # fmt: skip
    for name in book:
        getattr(wb, name)
    for sheet in wb.sheets:
        if sheet.kind != "worksheet":
            continue
        for _ in sheet.cells(rich=True):
            pass
        for name in ("unparsed", "rows", "columns", "views", "format", "properties", "drawing",
                     "legacy_drawing", "comments", "threaded_comments", "tables",
                     "conditional_formats", "autofilter", "sort_state", "page_setup",
                     "print_area", "print_titles", "hyperlinks", "merges", "row_breaks",
                     "column_breaks", "background"):  # fmt: skip
            getattr(sheet, name)
    xlio.check(io.BytesIO(data))


SHEET = (
    f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetPr><tabColor theme="4" tint="0.5"/>'
    '<outlinePr summaryBelow="0"/></sheetPr><dimension ref="A1:B2"/><sheetViews>'
    '<sheetView zoomScale="120" colorId="10" workbookViewId="0"><pane xSplit="1" ySplit="1" '
    'topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" '
    'activeCell="B2" sqref="B2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" '
    'baseColWidth="8" outlineLevelRow="1" outlineLevelCol="1"/><cols><col min="1" max="2" '
    'width="12" style="0" outlineLevel="1"/></cols><sheetData><row r="1" ht="20" '
    'customHeight="1" outlineLevel="1" s="0"><c r="A1" s="0"><v>1</v></c></row></sheetData>'
    '<rowBreaks count="1" manualBreakCount="1"><brk id="1" max="16383" man="1"/></rowBreaks>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    '<pageSetup paperSize="9" scale="80" fitToWidth="1" fitToHeight="0" firstPageNumber="3" '
    'copies="2" horizontalDpi="600" verticalDpi="600"/></worksheet>'
)
WORKBOOKS = {
    **{p.stem: p.read_bytes() for p in sorted(FILES.glob("*.xlsx"))},
    "crafted": package(
        sheet_xml=SHEET,
        workbook_extra='<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>',
    ),
}


@pytest.mark.parametrize("name", sorted(WORKBOOKS))
@bounded
def test_huge_numbers_give_a_value_or_a_refusal(name):
    data = WORKBOOKS[name]
    failures = []
    found = cases(data)
    assert found
    for part, attribute, value in found:
        try:
            read_everything(rewritten(data, part, attribute, value))
        except (xlio.PackageError, xlio.LimitError):
            pass
        except Exception as exc:  # the finding: any other exception
            what = f"{part} {attribute}={value[:12]}…: {type(exc).__name__}: {exc}"
            failures.append(what[:200])
    assert failures == []
