# xlio

**xlio is an internal package of the sheetmill project.** It reads xlsx files as Excel reads
them and changes their cells without disturbing anything else. It was extracted from
[xlrecalc](https://github.com/sheetmill/xlrecalc), whose own reader and writer it began as, to be
shared by the sheetmill packages. It is not intended for direct use: its API may change in any
minor release (such as 0.3 to 0.4) without the deprecation period the public sheetmill packages
give, while patch releases only fix bugs and keep the API. If you depend on it anyway, pin a
minor version (for example `xlio==0.5.*`).

It is written in pure Python with the standard library only; this documentation covers its
public names.

Which Excel it follows: `read_csv` reads a CSV file as Excel for Windows opens it, unless it is
given `platform="mac"`, which reads it as Excel for Mac does; both readings rest on results
captured from Excel under Microsoft 365. The readings of xlsx files (`open`, `check`) take no
platform: they were captured from Excel 16 for Mac and are not yet verified on Windows, except a
cell font's size, read as Excel for Windows holds it and not yet verified on the Mac.

Python 3.11 or later. The source is at <https://github.com/sheetmill/xlio>.

xlio is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats,
[xlrecalc](https://github.com/sheetmill/xlrecalc), which recalculates workbooks, and
[sheetpeek](https://github.com/sheetmill/sheetpeek), which shows workbooks as Excel does.

xlio includes Excel's built-in table and pivot table styles from ECMA-376, under Ecma
International's copyright notice, which is reproduced in the `NOTICE` file of the source.

```{toctree}
:maxdepth: 2

api
changelog
```
