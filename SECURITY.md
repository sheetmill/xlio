# Security policy

xlio reads xlsx and CSV files, including files from people the caller does not trust (with
`limits=xlio.ReadLimits.UNTRUSTED`), and changes cells of an xlsx file by copying every other part
as it is. A security problem here is a file that crashes the reader, hangs it, makes it take time
or memory beyond what its limits allow, or makes it read a part it should refuse; or a change the
writer makes that damages the rest of the file.

## Reporting a vulnerability

Please report security problems privately, through GitHub's private vulnerability reporting:
open the repository's **Security** tab and choose **Report a vulnerability**
(https://github.com/sheetmill/xlio/security/advisories/new). Do not open a public issue.

Useful in a report: the version, a file or steps that reproduce the problem (a crafted workbook or
CSV file is fine), the limits used, and what happens (a crash, a hang, memory use, a limit that
does not hold, a damaged file).

A problem in one of the other sheetmill packages belongs in that package's repository, reported
the same way: sheetpeek, which renders workbooks
(https://github.com/sheetmill/sheetpeek/security/advisories/new); xlrecalc, which evaluates
formulas (https://github.com/sheetmill/xlrecalc/security/advisories/new); and xlformat, which
formats values and measures text (https://github.com/sheetmill/xlformat/security/advisories/new).
If you are not sure which, report it here.

## Supported versions

xlio is in alpha. Fixes go into the next release; only the latest release receives fixes.

## What happens after a report

The report is acknowledged, reproduced and fixed in the next release, whose changelog states the
fix in its Security section. Reporters are credited there unless they prefer not to be.
