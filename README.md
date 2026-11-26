# xlio

**xlio is an internal package of the sheetmill project.** It reads xlsx files as Excel reads
them and changes their cells without disturbing anything else. It was extracted from
[xlrecalc](https://github.com/sheetmill/xlrecalc), whose own reader and writer it began as, to be
shared by the sheetmill packages. It is not intended for direct use: its API may change in any
minor release (such as 0.3 to 0.4) without the deprecation period the public sheetmill packages
give, while patch releases only fix bugs and keep the API. If you depend on it anyway, pin a
minor version (for example `xlio==0.4.*`).

It is written in pure Python with the standard library only. The reader checks a package before
it parses anything (part count and sizes, compression methods, encryption, names used twice, a
DTD in any part) and refuses what Excel refuses; the writer copies every part it does not change
byte for byte and rewrites only the cells it is given.

Which Excel it follows: `read_csv` reads a CSV file as Excel for Windows opens it, unless it is
given `platform="mac"`, which reads it as Excel for Mac does; both readings rest on results
captured from Excel under Microsoft 365. The readings of xlsx files (`open`, `check`) take no
platform: they were captured from Excel 16 for Mac and are not yet verified on Windows.

xlio is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats,
[xlrecalc](https://github.com/sheetmill/xlrecalc), which recalculates workbooks, and
[sheetpeek](https://github.com/sheetmill/sheetpeek), which shows workbooks as Excel does.

## Files from people you do not trust

Open them with `limits=xlio.ReadLimits.UNTRUSTED`, which bounds what reading one may cost (its
parts, their sizes, its cells, its text). To report a security problem, see
[`SECURITY.md`](https://github.com/sheetmill/xlio/blob/main/SECURITY.md).

## Development

`uv sync --group docs` then `uv run pytest`;
[`CONTRIBUTING.md`](https://github.com/sheetmill/xlio/blob/main/CONTRIBUTING.md) has the pip route, the
full set of checks and the commit conventions.

Most of the code is written with Claude Code, reviewed before each commit and tested against
results captured from Excel; see [`CONTRIBUTING.md`](https://github.com/sheetmill/xlio/blob/main/CONTRIBUTING.md).

## License

The code is under the MIT license ([`LICENSE`](https://github.com/sheetmill/xlio/blob/main/LICENSE)).
`src/xlio/data/presetTableStyles.xml`, the built-in table and pivot table styles from ECMA-376,
is under Ecma International's text copyright policy, as published by Ecma; see
[`NOTICE`](https://github.com/sheetmill/xlio/blob/main/NOTICE).
