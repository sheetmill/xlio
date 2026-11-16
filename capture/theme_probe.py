"""The theme color index past which Excel refuses a file (it refused 99; the theme has 12).

    uv run python -m capture.theme_probe

The control is the ``attributes`` probe's (openpyxl's workbook, palette removed); each case sets
the default font's color to ``theme="n"``. The script opens 11, 12 and 98 first, then bisects
between the largest index Excel opened and the smallest it refused, one file at a time, so a few
files settle the bound. Each row gives the outcome and the color Excel's copy writes. Output:
``tests/captures/theme_bound.tsv``.

Run it only while Excel is running and answers AppleScript, no other capture is running, and no
test or run is feeding crafted input to the reader (CAPTURING.md).
Workbooks live in Excel's container and are deleted afterwards.
"""

from __future__ import annotations

import datetime
import zipfile

from capture.attributes_probe import FONT_COLOR, STYLES, control
from capture.edition import edition_line
from capture.excel import refuse_open_workbooks
from capture.probe import CAPTURES, CONTAINER, ROOT, Case, build, excel_version, read, save_copy


def case(n: int) -> Case:
    return Case(f"theme {n}", ((STYLES, FONT_COLOR, f'<color theme="{n}" />'),))


def main() -> None:
    refuse_open_workbooks()
    version = excel_version()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    data = control()
    src, copy = CONTAINER / "xlio_theme_src.xlsx", CONTAINER / "xlio_theme_copy.xlsx"
    outcomes: dict[int, bool] = {}
    rows = []

    def probe(n: int) -> bool:
        try:
            src.write_bytes(build(data, case(n)))
            copy.unlink(missing_ok=True)
            error = save_copy(src, copy)
            if error:
                rows.append(f"theme {n}\trefused\t{error}\t")
            else:
                with zipfile.ZipFile(copy) as z:
                    saved = read(z, STYLES, r"<fonts\b[^>]*>\s*<font>.*?</font>")
                rows.append(f"theme {n}\topened\t\t{saved}")
        finally:
            src.unlink(missing_ok=True)
            copy.unlink(missing_ok=True)
        print(rows[-1], flush=True)
        outcomes[n] = not error
        return not error

    for n in (11, 12, 98):
        probe(n)
    opened = max((n for n, ok in outcomes.items() if ok), default=None)
    refused = min((n for n, ok in outcomes.items() if not ok), default=99)
    if opened is not None and opened < refused:
        while refused - opened > 1:
            middle = (opened + refused) // 2
            if probe(middle):
                opened = middle
            else:
                refused = middle
    header = [
        f"# suite: theme_bound (capture.theme_probe); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: the attributes control with the default "
        "font's color set to theme n, opened with alerts off and saved as a copy; 11, 12 and 98 "
        "first, then bisected",
        "# refused = error -50; font 0 as Excel's copy writes it",
        "case\toutcome\terror\tfont 0",
    ]
    out = CAPTURES / "theme_bound.tsv"
    out.write_text("\n".join(header + rows) + "\n", encoding="utf-8")
    print(f"wrote {out.relative_to(ROOT)}: largest opened {opened}, smallest refused {refused}")


if __name__ == "__main__":
    main()
