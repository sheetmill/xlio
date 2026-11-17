"""What the README and the documentation state about the package agrees with the package."""

from __future__ import annotations

import re
from pathlib import Path

import xlio


def test_the_readme_pins_this_minor_version():
    version = re.match(r"(\d+\.\d+)\.", xlio.__version__)
    assert version is not None
    for doc in ("README.md", "docs/index.md"):
        text = (Path(__file__).parents[1] / doc).read_text(encoding="utf-8")
        assert re.findall(r"`xlio==([\d.]+)\.\*`", text) == [version[1]], doc
