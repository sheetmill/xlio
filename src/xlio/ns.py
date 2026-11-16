"""Namespaces, relationship types and extension URIs, as Excel writes and matches them.

Extension URIs are compared exactly, case included, because Excel does: it reads the 2010
conditional-formatting extension of a worksheet only under :data:`EXT_CONDITIONAL_FORMATTINGS`
(``407d`` in lower case) and ignores the same URI printed with ``407D``, as [MS-XLSX] 2.3.6
prints it. Where a constant differs from the specification's text, its comment says so.
"""

from __future__ import annotations

# Namespaces ------------------------------------------------------------------------------------

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
"""SpreadsheetML, transitional."""
MAIN_STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main"
"""SpreadsheetML, strict."""
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
"""The ``r:`` namespace of relationship ids in a part, transitional."""
RELATIONSHIPS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"
"""The ``r:`` namespace, strict."""
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
"""The namespace of ``.rels`` parts."""
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
"""The namespace of ``[Content_Types].xml``."""
MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006"
"""``mc:``, of ``AlternateContent``."""
DRAWING_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main"
"""DrawingML's ``a:``, of themes and drawings."""
SPREADSHEET_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
"""``xdr:``, of drawing parts."""
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
"""``x14:``, Excel 2010's extensions."""
XM = "http://schemas.microsoft.com/office/excel/2006/main"
"""``xm:``, formulas and ranges inside ``x14:`` elements."""

# Relationship types ----------------------------------------------------------------------------

_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

OFFICE_DOCUMENT = _DOC + "officeDocument"
"""From the package to the workbook part."""
WORKSHEET = _DOC + "worksheet"
"""From the workbook to a worksheet."""
CHARTSHEET = _DOC + "chartsheet"
"""From the workbook to a chartsheet."""
DIALOGSHEET = _DOC + "dialogsheet"
"""From the workbook to a dialog sheet."""
SHARED_STRINGS = _DOC + "sharedStrings"
"""From the workbook to the shared strings."""
STYLES = _DOC + "styles"
"""From the workbook to the styles."""
THEME = _DOC + "theme"
"""From the workbook to its theme."""
SHEET_METADATA = _DOC + "sheetMetadata"
"""From the workbook to the cell and value metadata."""
TABLE = _DOC + "table"
"""From a worksheet to a table part."""
HYPERLINK = _DOC + "hyperlink"
"""From a worksheet to a hyperlink's external target."""
DRAWING = _DOC + "drawing"
"""From a worksheet to its drawing part."""
VML_DRAWING = _DOC + "vmlDrawing"
"""From a worksheet to its VML drawing (notes, form controls)."""
COMMENTS = _DOC + "comments"
"""From a worksheet to its notes."""
THREADED_COMMENT = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
"""From a worksheet to its threaded comments."""
PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"
"""From the workbook to the people of threaded comments."""

# Extension URIs (exact, as Excel matches them) ------------------------------------------------

EXT_CONDITIONAL_FORMATTINGS = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
"""A worksheet's ``x14:conditionalFormattings``. [MS-XLSX] prints ``407D``, which Excel ignores."""
EXT_CONDITIONAL_FORMATTING_ID = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
"""A rule's ``x14:id``, which joins it to its ``x14:cfRule``, as [MS-XLSX] prints it."""
