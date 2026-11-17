# Capturing Excel's behavior

What xlio reads, refuses and writes is checked against what Excel 16 for Mac does with the same
file. Such behavior is captured from Excel, never written from memory or taken from another
program's output. Captured data lives in `tests/captures/`, the scripts that produce it in
`capture/`, and each test that uses a capture names it.

## Method

1. A script in `capture/` builds small workbooks, each differing from a control in one way (a
   `<v>` rewritten as `nan`, a row moved past the grid).
2. Excel opens each workbook with alerts off and saves a copy (`capture/save_copy.applescript`).
   A file Excel would offer to repair fails to open with error -50 while alerts are off; that
   is recorded as `refused`.
3. The script reads Excel's verdict and the values Excel saved from the copy's XML, and writes a
   TSV with a header naming the suite, the Excel build, the edition and update channel, the
   locale, the date and the method.

## Steps that need a person

Some behavior shows only to a person: Excel's repair dialog, its notice about macros, and its
notice of default conversions for a `.csv` file, which an AppleScript open does not show (it applies the default,
"Convert"). Such a step opens the file through Launch Services (`open -a "Microsoft Excel"`) and
then exits; it sends Excel nothing more until you have answered the dialog and run its next step
(answering the conversion notice while a script was querying Excel crashed Excel twice, on
16.115). Copy a repair report's text into the next step's `--report`; the script keeps only its
records, without paths.

## Rules for the scripts

- Non-interactive: `display alerts` is turned off and restored on every exit path, including
  errors; otherwise Excel would silently fail to open or repair files afterwards.
- A script refuses to run while Excel has any workbook open, opens only workbooks it creates,
  and closes them without saving unless saving is the point.
- Excel is sandboxed: workbooks are built inside its container
  (`~/Library/Containers/com.microsoft.Excel/Data/Documents/xlio/`), so it opens them without
  asking for access, and are deleted afterwards.
- The edition is stated by the person capturing (`XLIO_EXCEL_EDITION`, or
  `capture/build/excel_edition.txt`), since the build number alone does not say which features
  Excel has (`capture/edition.py`).
- Floats are written into workbooks as their shortest exact text (`repr`).
- One capture at a time, and none while a test or a run feeds crafted input to the reader (the
  memory watchdog would compete with Excel).

## Captures

| capture | script | what it records |
|---|---|---|
| `package` | `capture/package_probe.py` | archives, DTDs and sizes no column, row or font can have: whether Excel opens the file, and the widths, heights and font size it reads |
| `value_forms` | `capture/value_probe.py` | number forms Excel never writes in a `<v>` (`nan`, `INF`, `1e400`, `+5`, ` 5`, a subnormal), and rows and columns off the grid: whether Excel opens the file, and the value it reads |
| `errors` | `capture/errors_probe.py` | each error code written plainly in a `<v>` (as a constant and as a formula's cached value), `#VALUE!` with a rich value of each `errorType` 0 to 20, and formulas returning the newer errors: whether Excel opens the file, and the error it reads (`ERROR.TYPE`, `VALUETOTEXT`) and writes |
| `errors_stale` | `capture/errors_probe.py --stale` | a formula cell whose `vm` leads to an error rich value (`#CALC!`) while its `<v>` caches 5, TRUE, text or another error, planted in a file Excel saved: the error Excel shows and saves |
| `views` | `capture/views_probe.py` | `activeCell` values Excel does not write, selections for absent panes, zoom, `activeTab` and sheet `state` out of range, head records after `sheetData`: whether Excel opens the file, and the view it saves |
| `attributes` | `capture/attributes_probe.py` | flags that are not booleans, sizes that are not numbers, indexes past their tables, colors, rotations and font sizes out of range, page setup values, absent margins, print names: whether Excel opens the file, and what it saves |
| `theme_bound` | `capture/theme_probe.py` | the theme color index past which Excel refuses a file, bisected |
| `text` | `capture/text_probe.py` | `.csv` files opened through File > Open's defaults (AppleScript): every cell of the copy with its type, value and number format, the sheet name and the standard font, for conversions (quoted and bare), encodings and structure |
| `text_launch_*` | `capture/text_probe.py --launch-open`, `--launch CHOICE` | three of those files opened as a person opens them (Launch Services), Excel's notice of default conversions answered by the person (`convert`, `dont`, or `none` when it did not appear) |
| `strings`, `strings_cached` | `capture/strings_probe.py` | escaped surrogates (a pair, lone halves, the pair reversed or in lowercase hex) in a shared string, an inline string and a formula's cached text, and `t="str"` cells with an empty value and no formula: the text Excel writes back and what `LEN`, `UNICODE`, `ISTEXT` and `ISBLANK` give for it |
| `macrosheet` | `capture/macrosheet_probe.py` | Excel 4.0 macro sheets (in an `.xlsm` and an `.xlsx`, international, with another root element) opened by a person, Excel's macro notice answered: what the `.xlsm` copy keeps |
| `outline` | `capture/outline_probe.py` | outline levels past 7, negative or not integers, on a row, a column and the sheet's `sheetFormatPr`: whether Excel opens the file, and the levels it saves |
| `text_extra` | `capture/text_extra_probe.py` | `.csv` files with two-digit years, and file names that make sheet names Excel would not take (long, reserved characters, `History`): the cells and the sheet's name |
| `text_limits` | `capture/text_extra_probe.py --limits-open`, `--limits-save` | `.csv` files past Excel's limits (a field of 32,767 characters and more, 1,048,577 rows, 16,385 columns), opened by a person: what the copy keeps, and the dialog Excel shows |
| `repair` | `capture/repair_probe.py` | reversed drawing anchors and invalid conditional formats opened by a person, Excel's repair answered Yes: what the repaired workbook keeps, and the repair report's records |

`capture/fixtures.py` is not a capture: it has Excel save workbooks XlsxWriter builds (tables
and filters, conditional formats, drawings, notes) into `tests/files/excel/`, and with `--adopt`
takes a workbook made by hand in Excel (threaded comments, icons, pictures in cells), each
stripped of paths and people.

`package` and `value_forms` gave the same results under Office 2021 (Excel 16.114.3) and
Microsoft 365 (Excel 16.115, Current Channel).

The `text`, `text_launch_*`, `text_extra` and `text_limits` captures are not used by the code yet: they record how Excel
opens CSV files, for the planned text formats (CSV, TSV, fixed width).

## Captures kept elsewhere

Some rules were captured by xlrecalc (https://github.com/sheetmill/xlrecalc/), where this code
was first written, and are cited with a link: the outer whitespace Excel drops
([`whitespace.tsv`](https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/whitespace.tsv)),
dynamic-array and rich-value metadata
([`arrays.tsv`](https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/arrays.tsv)) and
text results read by Excel
([`test_excel_live.py`](https://github.com/sheetmill/xlrecalc/blob/main/tests/test_excel_live.py)),
and when Excel recalculates on opening: it keeps the cached values of a file it saved itself
and computes every formula of one openpyxl wrote
([`open_calc.tsv`](https://github.com/sheetmill/xlrecalc/blob/main/tests/captures/open_calc.tsv)).

## Platforms

Every capture here is from Excel for Mac. Excel for Windows is not captured; where the two may
differ, the reader follows the Mac.

## Test workbooks

`tests/files/excel/` holds the workbooks `capture/fixtures.py` had Excel save.
`tests/files/marker_typed.xlsx`, `iterate_cached.xlsx` and `excel_formulas.xlsx` were saved by
Excel for Mac for xlrecalc's tests (https://github.com/sheetmill/xlrecalc/blob/main/CAPTURING.md)
and copied from there, stripped of paths and editors.
