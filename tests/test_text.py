"""CSV files read as Excel for Mac opens them.

Captures ``text``, ``text_launch_*``, ``text_extra`` and ``text_limits``
(``capture/text_probe.py``, ``capture/text_extra_probe.py``): each case is a ``.csv`` file Excel
opened and saved as a workbook; the TSV holds every cell of the copy (type, value and number
format) and the sheet's name. Here :func:`xlio.read_csv` must give the same cells.
"""

from __future__ import annotations

import datetime
import io
from pathlib import Path

import pytest
from capture.text_extra_probe import YEARS, extra_files
from capture.text_probe import LAUNCH, files

import xlio
from xlio._text import sheet_name_of

CAPTURES = Path(__file__).with_name("captures")
TODAY = datetime.date(2026, 11, 13)  # the day of capture text: a date without a year takes it


def rows(suite: str) -> list[list[str]]:
    lines = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith(("# ", "case\t"))]


def unescape(text: str) -> str:
    out, i = [], 0
    while i < len(text):
        if text[i] == "\\" and i + 1 < len(text):
            out.append({"n": "\n", "t": "\t", "r": "\r", "\\": "\\"}[text[i + 1]])
            i += 2
        else:
            out.append(text[i])
            i += 1
    return "".join(out)


def expected(suite: str) -> dict[str, dict[str, tuple[str, str]]]:
    """Case -> cell -> (type:value, format), the font line left out."""
    out: dict[str, dict[str, tuple[str, str]]] = {}
    for row in rows(suite):
        if row[1] == "opened" and row[3] != "font":
            out.setdefault(row[0], {})[row[3]] = (row[4], row[5].strip())
    return out


def got(wb: xlio.Workbook) -> dict[str, tuple[str, str]]:
    formats = wb.styles.cell_formats
    out = {}
    for cell in wb.sheets[0].cells():
        ref = f"{xlio.CellRef(cell.row, cell.column)}"
        fmt = formats[cell.style]
        shown = (
            f"builtin {fmt.number_format_id}" if fmt.number_format is None else fmt.number_format
        )
        if cell.formula is not None:
            out[ref] = (f"={cell.formula.text}", shown)
        elif cell.type == "n":
            out[ref] = (f"n:{cell.value!r}", shown)
        elif cell.type == "s":
            out[ref] = (f"s:{cell.value}", shown)
        elif cell.type == "b":
            out[ref] = (f"b:{int(bool(cell.value))}", shown)
        else:
            out[ref] = (f"{cell.type}:{cell.value}", shown)
    return out


def same(want: tuple[str, str], have: tuple[str, str]) -> bool:
    (value, fmt), (got_value, got_fmt) = want, have
    if fmt != got_fmt:
        return False
    if value.startswith("n:="):  # a formula, cached by Excel: xlio holds no value
        return got_value == value[2:].split(" -> ")[0]
    if value.startswith("n:") and got_value.startswith("n:"):
        return float(value[2:]) == float(got_value[2:])
    return unescape(value) == got_value


def compare(suite: str, cases: dict[str, bytes], options: xlio.CsvOptions) -> None:
    want = expected(suite)
    failed = {}
    for case, data in cases.items():
        with xlio.read_csv(io.BytesIO(data), options) as wb:
            have = got(wb)
        cells = want.get(case, {})
        wrong = {r: (cells.get(r), have.get(r)) for r in set(cells) | set(have)
                 if r not in cells or r not in have or not same(cells[r], have[r])}  # fmt: skip
        if wrong:
            failed[case] = wrong
    assert failed == {}


def test_every_case_of_capture_text():
    compare("text", files(), xlio.CsvOptions(today=TODAY))


@pytest.mark.parametrize(
    ("answer", "conversions"), [("convert", "excel"), ("dont", "keep"), ("none", "excel")]
)
def test_the_notice_answered(answer, conversions):
    suite = f"text_launch_{answer}"
    cases = {k: v for k, v in files().items() if k in LAUNCH and k in expected(suite)}
    assert cases
    compare(suite, cases, xlio.CsvOptions(today=TODAY, conversions=conversions))


def test_two_digit_years_and_sheet_names():
    want = expected("text_extra")
    with xlio.read_csv(io.BytesIO(extra_files()["years"][1]), xlio.CsvOptions(today=TODAY)) as wb:
        have = got(wb)
    assert len(have) == len(YEARS)
    assert all(same(want["years"][r], have[r]) for r in want["years"])
    names = {row[0]: row[2] for row in rows("text_extra") if row[0].startswith("name ")}
    for case, (file_name, _) in extra_files().items():
        if case.startswith("name ") and case not in ("name colon", "name backslash"):
            assert sheet_name_of(file_name) == names[case], case


def test_a_long_line_continues_on_the_next_row():
    for case, (_, _, rows_, _cols, cells, length) in {
        r[0]: tuple(r[1:]) for r in rows("text_limits") if r[0].startswith("field")
    }.items():
        size = int(case.split()[1])
        with xlio.read_csv(io.BytesIO(b"x" * size + b"\n")) as wb:
            have = list(wb.sheets[0].cells())
        assert (len(have), len(have[0].value)) == (int(cells), int(length)), case
        assert max(c.row for c in have) == int(rows_)


def test_none_keeps_every_field_as_text():
    with xlio.read_csv(io.BytesIO(b"1,=1+1,TRUE\n"), xlio.CsvOptions(conversions="none")) as wb:
        assert [(c.type, c.value) for c in wb.sheets[0].cells()] == [
            ("s", "1"), ("s", "=1+1"), ("s", "TRUE"),
        ]  # fmt: skip


def test_formulas_as_text():
    with xlio.read_csv(io.BytesIO(b"=1+1\n"), xlio.CsvOptions(formulas="text")) as wb:
        (cell,) = wb.sheets[0].cells()
    assert (cell.type, cell.value, cell.formula) == ("s", "=1+1", None)


def test_columns_past_the_grid_are_dropped():
    case = next(r for r in rows("text_limits") if r[0] == "columns 16385")
    data = (",".join(str(n) for n in range(1, 16386)) + "\n").encode()
    with xlio.read_csv(io.BytesIO(data)) as wb:
        cells = list(wb.sheets[0].cells())
    assert (len(cells), cells[-1].column) == (int(case[5]), int(case[4]))


def test_a_csv_file_is_not_a_package(tmp_path):
    source = tmp_path / "data.csv"
    source.write_bytes(b"a,b\n1,2\n")
    for call in (
        lambda: xlio.open(source),
        lambda: xlio.patch(
            source, tmp_path / "out.xlsx", cells={"data": {(1, 1): xlio.CellEdit(None, 5.0)}}
        ),
    ):
        with pytest.raises(xlio.PackageError, match=r"a CSV file is read with xlio\.read_csv"):
            call()
    with xlio.read_csv(source) as wb:  # its parts are built for it
        assert wb.sheets[0].name == "data"
        assert "xl/worksheets/sheet1.xml" in wb.parts


def test_a_text_file_past_its_limit_is_refused():
    limits = xlio.ReadLimits(max_text_bytes=8)
    with pytest.raises(xlio.LimitError, match="limit of 8 bytes"):
        xlio.read_csv(io.BytesIO(b"1,2,3,4,5\n"), limits=limits)
    with xlio.read_csv(io.BytesIO(b"1,2,3\n"), limits=limits) as wb:
        assert len(list(wb.sheets[0].cells())) == 3
    assert xlio.ReadLimits.UNTRUSTED.max_text_bytes == 32 << 20
