# Contributing

## Issues and bug reports

Bug reports are welcome. The most useful report includes a small workbook that shows the
problem (or the part of it that matters, such as one sheet's XML), what Excel does with it, what
xlio does, and your Excel version and platform. Please do not attach a workbook that holds
anyone's private data.

## Pull requests

The project is just getting started, so pull requests aren't being accepted yet while
the API, release process and review practices settle. In the meantime, if you have a change in
mind, please open an issue describing it.

When pull requests open, these rules will apply:

- AI-assisted contributions must be disclosed in the pull request description.
- The contributor must understand every line of the change and be able to explain it when asked.
- Undisclosed or unexplained AI-generated pull requests may be closed without review.

## Setting up

Python 3.11 or later. Either route installs the package in editable mode with the development
and documentation tools:

```bash
uv sync --group docs
```

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e . --group dev --group docs      # needs pip 25.1 or later
```

`uv.lock` pins the versions CI tests with, and `uv sync` installs exactly those versions.
pip and other tools ignore it and resolve the latest versions the dependency groups in
`pyproject.toml` allow, which is normally fine but is not what CI runs.

## Checks to run before proposing a change

All of these must pass. CI runs the same set on Python 3.11 to 3.14, plus a run without
openpyxl, a development dependency. Prefix each with `uv run` on the uv route.

```bash
pytest
ruff check .
ruff format --check .
mypy
sphinx-build -n -W docs docs/_build/html # the strict documentation build
```

`pytest --tier slow`, as CI runs it, adds the costly cases. Tests that feed crafted packages to
the reader run in a child process under a memory limit (`tests/watchdog.py`), so a guard that
regresses fails the test instead of exhausting the machine.

## Expected results

What the reader accepts, refuses and reads is what Excel 16 for Mac does with the same file.
Such behavior is captured from Excel, never written from memory or taken from another program
(`CAPTURING.md`). Captured files are never edited by hand.

## Commit messages

A commit made with the help of an AI tool ends with the trailer

```
Assisted-by: <tool>:<model ID>
```

as its last line, after a blank line, naming the tool and the exact model, for example
`Assisted-by: Claude Code:claude-opus-5-5`. One trailer per tool and model used.

## How this project's own commits are made

Most of the code is written with Claude Code, reviewed before each commit and tested against
results captured from real Excel. Every such commit carries the git trailer
`Assisted-by: Claude Code:<model ID>`, so the history records which model took part.
