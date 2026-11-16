"""The public writer, ``xlio.patch``: Python values in, the file's own forms out."""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest
from capture.paths import HOME_PATH

import xlio
from tests.test_read import package
from tests.test_write import ROWS, raw_entries, sheet_xml

FILES = Path(__file__).with_name("files")


def save(tmp_path: Path, data: bytes) -> Path:
    path = tmp_path / "book.xlsx"
    path.write_bytes(data)
    return path


def read(path: Path) -> dict[tuple[int, int], xlio.Cell]:
    with xlio.open(path) as wb:
        return {(c.row, c.column): c for c in wb["Data"].cells()}


def test_cached_results_as_python_values(tmp_path):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    results = {(1, 2): 6, (1, 3): "x2", (1, 4): True, (1, 5): xlio.ErrorCode.DIV0}
    count = xlio.patch(src, out, cached={"Data": lambda r, c: results.get((r, c))})
    assert count == 3  # E1 already holds #DIV/0!
    cells = read(out)
    assert [(cells[1, k].type, cells[1, k].value) for k in (2, 3, 4, 5)] == [
        ("n", 6.0),
        ("str", "x2"),
        ("b", True),
        ("e", xlio.ErrorCode.DIV0),
    ]
    with xlio.open(out) as wb:
        assert wb.calc.full_calc_on_load


def test_an_error_code_and_text_that_looks_like_one_differ(tmp_path):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    xlio.patch(src, out, cached={"Data": lambda r, c: "#N/A" if c == 3 else None})
    assert (read(out)[1, 3].type, read(out)[1, 3].value) == ("str", "#N/A")


def test_cell_edits_and_fills(tmp_path):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    edits = {
        (1, 1): xlio.CellEdit(None, 7),
        (2, 1): xlio.CellEdit(None, " padded "),
        (2, 2): xlio.CellEdit("A1*2", 14.0),
        (3, 1): xlio.CellEdit("SEQUENCE(2)", 1.0, array_ref="A3:A4", dynamic=True),
    }
    fill = xlio.ArrayFill(4, 1, 4, 1, lambda r, c: 2.0)
    xlio.patch(src, out, cells={"Data": edits}, fills={"Data": [fill]})
    cells = read(out)
    assert cells[1, 1].value == 7.0
    assert (cells[2, 1].type, cells[2, 1].value) == ("inlineStr", " padded ")
    assert (cells[2, 2].formula.text, cells[2, 2].value) == ("A1*2", 14.0)
    assert (cells[3, 1].formula.kind, cells[3, 1].formula.ref) == ("array", "A3:A4")
    assert cells[4, 1].value == 2.0
    with xlio.open(out) as wb:
        assert cells[3, 1].cm in wb.cell_metadata.dynamic_arrays
    assert '<c r="A1"><v>7</v></c>' in sheet_xml(out)


def test_a_value_that_is_not_a_cell_value_is_refused(tmp_path):
    src = save(tmp_path, package(ROWS))
    with pytest.raises(TypeError, match="not a cell value"):
        xlio.patch(src, tmp_path / "out.xlsx", cells={"Data": {(1, 1): xlio.CellEdit(None, [1])}})  # type: ignore[arg-type]


def test_nothing_else_changes(tmp_path):
    src = save(tmp_path, package(ROWS, strings=["<si><t>a</t></si>"]))
    out = tmp_path / "out.xlsx"
    xlio.patch(src, out, cached={"Data": lambda r, c: 6.0 if c == 2 else None})
    before, after = raw_entries(src), raw_entries(out)
    assert {n for n in before if before[n] != after[n]} == {
        "xl/worksheets/sheet1.xml",
        "xl/workbook.xml",
    }
    kept = tmp_path / "kept.xlsx"
    xlio.patch(src, kept, full_calc_on_load=False)
    assert raw_entries(src) == raw_entries(kept)


def test_a_file_saved_by_excel_keeps_every_other_part(tmp_path):
    src = FILES / "excel_formulas.xlsx"
    with xlio.open(src) as wb:
        sheet = wb.sheets[0]
        formula = next(c for c in sheet.cells() if c.formula is not None and c.type == "n")
    out = tmp_path / "out.xlsx"
    at = (formula.row, formula.column)
    changed = xlio.patch(
        src, out, cached={sheet.name: lambda r, c: 12345.5 if (r, c) == at else None}
    )
    assert changed == 1
    with xlio.open(out) as wb:
        got = {(c.row, c.column): c.value for c in wb.sheets[0].cells()}
    assert got[at] == 12345.5
    before, after = raw_entries(src), raw_entries(out)
    assert list(before) == list(after)
    assert {n for n in before if before[n] != after[n]} == {sheet.part, "xl/workbook.xml"}


# Committed workbooks ----------------------------------------------------------------------------

COMMITTED = sorted(Path(__file__).parent.rglob("*.xlsx"))
TOOLS = {b"", b"openpyxl", b"XlsxWriter"}  # creators that name a library, not a person
PLACEHOLDER = b"Writer"


def problems(data: bytes) -> list[str]:
    """What in a workbook could identify a person or a machine; empty when nothing does."""
    found: list[str] = []
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for name in z.namelist():
            body = z.read(name)
            found += [
                f"{name}: {m.decode()}"
                for m in (b"absPath", b"lastModifiedBy", b"MSIP_Label_")
                if m in body
            ]
            if HOME_PATH.search(body):
                found.append(f"{name}: a home-directory path")
            if name in ("docProps/custom.xml", "docMetadata/LabelInfo.xml"):
                found.append(f"{name}: present")
            found += [
                f"{name}: creator {c!r}"
                for c in re.findall(rb"<dc:creator>([^<]*)</dc:creator>", body)
                if c not in TOOLS
            ]
            found += [
                f"{name}: {t.decode()} filled"
                for t in (b"Company", b"Manager")
                if re.search(rb"<" + t + rb">[^<]+</" + t + rb">", body)
            ]
            if name.startswith("xl/persons/"):
                found += [
                    f"{name}: a person other than the placeholder"
                    for p in re.findall(rb"<person\b[^>]*>", body)
                    if b'displayName="Writer"' not in p or b'providerId="None"' not in p
                ]
            found += [
                f"{name}: author {a!r}"
                for a in re.findall(rb"<author>([^<]*)</author>", body)
                if a not in (b"", PLACEHOLDER) and not a.startswith(b"tc={")
            ]
    return found


def test_every_committed_workbook_is_checked():
    assert len(COMMITTED) == 8


@pytest.mark.parametrize("path", COMMITTED, ids=lambda p: str(p.relative_to(Path(__file__).parent)))
def test_committed_workbooks_carry_no_paths_or_people(path):
    assert problems(path.read_bytes()) == []


def test_a_real_name_is_caught():
    path = Path(__file__).parent / "files" / "excel" / "handmade.xlsx"
    out = io.BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            body = src.read(info)
            if info.filename == "xl/persons/person.xml":
                body = body.replace(b'displayName="Writer"', b'displayName="Pat Example"')
            dst.writestr(info, body)
    assert problems(out.getvalue()) == [
        "xl/persons/person.xml: a person other than the placeholder"
    ]
