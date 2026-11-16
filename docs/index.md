# xlio

**xlio is an internal package of the sheetmill project.** It reads xlsx files as Excel reads
them and changes their cells without disturbing anything else. It was extracted from
[xlrecalc](https://github.com/sheetmill/xlrecalc), whose own reader and writer it began as, to be
shared by the sheetmill packages. It is not intended for direct use: its API may change in any
minor release (such as 0.2 to 0.3) without the deprecation period the public sheetmill packages
give, while patch releases only fix bugs and keep the API. If you depend on it anyway, pin a
minor version (for example `xlio==0.2.*`).

It is written in pure Python with the standard library only; this documentation covers its
public names.

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
