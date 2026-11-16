"""What Excel makes of a CSV file it opens (File > Open, en_US): conversions, encodings, structure.

    uv run python -m capture.text_probe            # every case, opened through AppleScript
    uv run python -m capture.text_probe --launch-open   # three cases, opened as by a person
    uv run python -m capture.text_probe --launch convert|dont|none  # then save, read

Each case is a ``.csv`` file built in Excel's container. Excel opens it with alerts off and saves
a copy as a workbook (``save_copy.applescript``), as File > Open does with Excel's defaults; no
import dialog is involved. From the copy the script reads every cell of the first sheet (row,
column, type and value as the XML has them), the number format Excel gave it, and the sheet's
name. Output: ``tests/captures/text.tsv``, one line per cell.

- ``convert``: one input per row in column A, quoted so commas stay inside (``convert
  unquoted``: those that need no quotes, bare): numbers in every
  form (signs, separators, currency, percent, fractions, exponents, leading zeros, precision),
  dates and times in several orders, booleans, errors, formulas and text;
- ``encoding <name>``: the same five words written in each encoding (with and without a BOM);
- ``structure <name>``: line endings, ragged rows, quotes, a semicolon or tab inside a ``.csv``,
  trailing empty lines and delimiters.

With ``--launch-open``, three cases (one of each kind) are written and opened through Launch
Services (``open -a "Microsoft Excel"``), as a person opens a file from the Finder, since an
AppleScript open can suppress dialogs a person sees; the script then exits, sending Excel
nothing while a notice may be open. Answer any notice of default conversions ("Convert" or
"Don't Convert"), then run ``--launch CHOICE`` (``convert`` or ``dont``, the answer you gave;
``none`` if no notice appeared), which saves each open workbook as a copy
(``save_open.applescript``), reads it and deletes the files. Output:
``tests/captures/text_launch_CHOICE.tsv``.

Each case also records the standard font of the copy (a CSV takes the application's), as a line
with the cell ``font``.

Run it only while Excel is running and answers AppleScript, no other capture is running, and no
test or run is feeding crafted input to the reader (CAPTURING.md).
Files live in Excel's container and are deleted afterwards.
"""

from __future__ import annotations

import datetime
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.probe import CAPTURES, CONTAINER, HERE, ROOT, TIME_LIMIT, excel_version, save_copy

CONVERT = [
    # numbers
    "1", "1.5", "-1", "1e3", "1E+03", "3E5", "1D5", "0012", "00123", "01234", "1,234", "1,234.5",
    "1,23", "1 000", "$5", "$1,234.50", "($5)", "(5)", "5-", "-5-", "5%", "50%", "1.5%", "%5",
    " 5", "5 ", "+5", ".5", "5.", "-.5", "1/2", "1 1/2", "0 1/2", "0x10", "1e400", "1e-400",
    "4.9e-324", "123456789012345678", "1234567890123456789012", "0.1234567890123456789", "-0",
    "inf", "nan", "NaN", "€5", "£5", "5 USD", "5€", "1.234,5",
    # dates and times
    "1/2/2026", "2026-01-02", "2026/01/02", "2026-01-02T03:04:05", "2026-01-02 03:04:05",
    "2-Jan", "Jan-2", "2 Jan 2026", "Jan 2, 2026", "January 2026", "13/1/2026", "1/13/2026",
    "1/2/26", "1/2", "1-2", "1-2-3", "12:30", "12:30:45", "12:30 PM", "12:30:45.5", "25:00",
    "1/2/2026 12:30", "2/29/2025", "2/29/2024", "1/1/1900", "12/31/1899", "1/1/10000",
    "555-1234", "123-45-6789",
    # booleans, errors
    "TRUE", "true", "False", "yes", "#N/A", "#DIV/0!", "#SPILL!", "#n/a",
    # formulas
    "=1+1", "=A1", "=SUM(1,2)", "+1+1", "-1+1", "@SUM(1)", "'=1", "=", "==1",
    # text
    "abc", "'abc", "'123", "", "  spaces  ", "a\nb", 'a"b', "\t5",
]  # fmt: skip
WORDS = ["café", "naïve", "€", "日本", "Ünïcödé"]
ENCODINGS = {
    "utf-8": b"",
    "utf-8 with BOM": b"\xef\xbb\xbf",
    "utf-16-le with BOM": b"\xff\xfe",
    "utf-16-be with BOM": b"\xfe\xff",
    "cp1252": b"",
}
STRUCTURE: dict[str, bytes] = {
    "crlf": b"a,b\r\nc,d\r\n",
    "lf": b"a,b\nc,d\n",
    "cr": b"a,b\rc,d\r",
    "no final newline": b"a,b\nc,d",
    "ragged rows": b"a\nb,c,d\ne,f\n",
    "trailing empty lines": b"a,b\n\n\n",
    "trailing delimiters": b"a,b,,\nc,,\n",
    "empty fields": b",a,,b\n",
    "quoted delimiter and newline": b'"a,b","c\nd"\n',
    "doubled quote": b'"a""b",c\n',
    "quote inside a field": b'a"b,c\n',
    "quote then text": b'"a"b,c\n',
    "unterminated quote": b'"abc,d\ne,f\n',
    "semicolons": b"a;b;c\n1;2;3\n",
    "tabs": b"a\tb\tc\n1\t2\t3\n",
    "spaces around a quoted field": b' "a" ,b\n',
    "leading blank line": b"\na,b\n",
}


def csv_text(values: list[str]) -> str:
    return "".join('"' + v.replace('"', '""') + '"\n' for v in values)


def files() -> dict[str, bytes]:
    out = {"convert": csv_text(CONVERT).encode("utf-8")}
    # The same inputs unquoted where a bare field can hold them: does quoting change a reading?
    bare = [v for v in CONVERT if not re.search(r'[,"\n\t]', v) and v.strip() == v and v]
    out["convert unquoted"] = "".join(v + "\n" for v in bare).encode("utf-8")
    text = csv_text(WORDS)
    for name, bom in ENCODINGS.items():
        codec = name.split(" ")[0]
        body = text.encode(codec, errors="replace")
        out[f"encoding {name}"] = bom + body
    for name, body in STRUCTURE.items():
        out[f"structure {name}"] = body
    return out


MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def formats(z: zipfile.ZipFile) -> list[str]:
    """Each cell format's number format code, by the cell format index."""
    root = ET.fromstring(z.read("xl/styles.xml"))
    codes = {n.get("numFmtId"): n.get("formatCode", "") for n in root.iter(MAIN + "numFmt")}
    xfs = root.find(f"{MAIN}cellXfs")
    out = []
    for xf in xfs if xfs is not None else ():
        fid = xf.get("numFmtId", "0")
        out.append(codes.get(fid, f"builtin {fid}"))
    return out


def cells(z: zipfile.ZipFile) -> list[str]:
    """``ref  type:value  format`` for every cell of the first sheet."""
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(MAIN + "si"):
            shared.append("".join(t.text or "" for t in si.iter(MAIN + "t")))
    fmts = formats(z)
    out = []
    sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    for c in sheet.iter(MAIN + "c"):
        t = c.get("t", "n")
        v = c.find(f"{MAIN}v")
        f = c.find(f"{MAIN}f")
        value = v.text if v is not None and v.text is not None else ""
        if t == "s" and value.isdigit():
            value = shared[int(value)]
        if f is not None:
            value = f"={f.text or ''} -> {value}"
        style = int(c.get("s", "0"))
        fmt = fmts[style] if style < len(fmts) else "?"
        shown = value.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n")
        shown = shown.replace("\r", "\\r")
        out.append(f"{c.get('r')}\t{t}:{shown}\t{fmt}")
    return out


def font(z: zipfile.ZipFile) -> str:
    """The copy's first font (the standard font), as name and size."""
    root = ET.fromstring(z.read("xl/styles.xml"))
    first = root.find(f"{MAIN}fonts/{MAIN}font")
    if first is None:
        return "?"
    name, size = first.find(f"{MAIN}name"), first.find(f"{MAIN}sz")
    named = name.get("val") if name is not None else "?"
    sized = size.get("val") if size is not None else "?"
    return f"{named} {sized}"


LAUNCH = ["convert", "encoding utf-8", "structure quoted delimiter and newline"]


def launch_open() -> None:
    """Write the launch cases into Excel's container and open each as a person does; no more.

    Nothing is sent to Excel over AppleScript while it may show a notice (sending queries
    while Excel's notice of default conversions was open, then answering it, crashed Excel
    twice on 16.115).
    """
    refuse_open_workbooks()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    for i, (name, data) in enumerate(files().items()):
        if name in LAUNCH:
            src = CONTAINER / f"xlio_text_{i}.csv"
            src.write_bytes(data)
            subprocess.run(["open", "-a", "Microsoft Excel", str(src)], check=True, timeout=30)
            print(f"{name}: opened {src.name}; answer any notice in Excel", flush=True)


def sheet_name(z: zipfile.ZipFile) -> str:
    found = re.search(r'<sheet [^>]*name="([^"]*)"', z.read("xl/workbook.xml").decode())
    return found.group(1) if found else "?"


def main(launch: str | None) -> None:
    if not launch:
        refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    rows = []
    for i, (name, data) in enumerate(files().items()):
        if launch and name not in LAUNCH:
            continue
        src = CONTAINER / f"xlio_text_{i}.csv"
        copy = CONTAINER / f"xlio_text_{i}_copy.xlsx"
        try:
            copy.unlink(missing_ok=True)
            if launch:
                result = subprocess.run(
                    ["gtimeout", "-k", "5", TIME_LIMIT, "osascript",
                     str(HERE / "save_open.applescript"), src.name, str(copy)],
                    capture_output=True, text=True, check=False,
                )  # fmt: skip
                error = "" if result.returncode == 0 else result.stderr.strip()[:120]
            else:
                src.write_bytes(data)
                error = save_copy(src, copy)
            if error:
                rows.append(f"{name}\trefused\t{error}\t\t\t")
            else:
                with zipfile.ZipFile(copy) as z:
                    title = sheet_name(z)
                    rows.append(f"{name}\topened\t{title}\tfont\t{font(z)}\t")
                    for line in cells(z):
                        rows.append(f"{name}\topened\t{title}\t{line}")
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print(f"{name}: done", flush=True)
    suite = f"text_launch_{launch}" if launch else "text"
    method = (
        "a .csv file opened through Launch Services (open -a), as a person opens it, Excel's "
        f"notice of default conversions answered '{launch}' by the person (none: not shown), "
        "then saved as a workbook copy through AppleScript"
        if launch
        else "a .csv file opened with alerts off (File > Open's defaults) and saved as a "
        "workbook copy"
    )
    header = [
        f"# suite: {suite} (capture.text_probe); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: {method}",
        "# one line per cell of the copy's first sheet: type:value as the XML has it (a shared "
        "string resolved, a formula as =text -> cached), and the cell's number format; \\n, \\r, "
        "\\t escaped",
        "case\toutcome\tsheet\tcell\tvalue\tformat",
    ]
    out = CAPTURES / f"{suite}.tsv"
    out.write_text("\n".join(header + rows) + "\n", encoding="utf-8")
    print(f"wrote {out.relative_to(ROOT)}: {len(rows)} lines")


if __name__ == "__main__":
    args = sys.argv[1:]
    if args == ["--launch-open"]:
        launch_open()
    elif args[:1] == ["--launch"]:
        if args[1:2] not in (["convert"], ["dont"], ["none"]):
            raise SystemExit("--launch takes the answer you gave: convert, dont or none")
        main(args[1])
    else:
        main(None)
