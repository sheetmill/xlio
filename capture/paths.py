"""What must not be in a committed workbook: a path into someone's home directory.

Shared by the fixture tooling, which refuses to write such a workbook, and the tests, which
check every committed one. A cached value such as ``CELL("filename")`` holds the saved file's
path, so every part is searched, in each form a path takes in XML.
"""

from __future__ import annotations

import re

HOME_PATH = re.compile(rb"/Users/|/home/|[A-Za-z]:\\\\Users\\\\|[A-Za-z]:\\Users\\|file:///")
