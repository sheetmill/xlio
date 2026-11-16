# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- `patch(full_calc_on_load=False)` removes a `fullCalcOnLoad` request the file holds, as Excel
  removes it when it saves a file. In 0.1.0 it left the request in place.

### Fixed
- `Sheet.cells()` reads an Excel 4.0 macro sheet's cells as Excel reads them, as it reads a
  worksheet's, where 0.1.0 raised `ValueError`; `check()` reads macro sheets whole too. Dialog
  sheets still raise.
- A macro sheet whose root element is in SpreadsheetML's namespace, rather than Excel's
  `xm:macrosheet`, reads without its columns, rows and cells, as Excel opens it.
- A character outside the Basic Multilingual Plane written as an escaped surrogate pair
  (`_xD83D__xDE00_`) reads as one character (U+1F600), as Excel reads it, not as two lone
  surrogates that no encoding can write. A lone escaped surrogate is still read as one
  character, as Excel holds it.
- Outline levels read as Excel reads them: a row's or column's `outlineLevel` is kept modulo 8
  (9 reads as 1), and one that is not decimal digits (`-1`, `1.5`, `+2`) reads as no level; a
  sheet's `outlineLevelRow` or `outlineLevelCol` of 8 to 255 is refused (`PackageError`, and by
  `check()`), as Excel refuses it, and one that is not decimal digits reads as 0.
- `patch()` leaves a formula cell untouched, `vm` included, when its new value is the error
  its rich value already stands for (`#CALC!`, say), so Excel's own record of it stays; any
  other new value still drops the `vm`.

## [0.1.0] - 2026-11-15

### Added
- The xlsx reader and surgical writer, extracted from xlrecalc (where they are private) to be
  shared by the sheetmill packages.
- `xlio.open()` returns a `Workbook`: every sheet with its kind and state (`Sheet.kind`,
  `Sheet.state`), defined names, calculation settings, the date system, dynamic-array and
  rich-value metadata, and whether the file is a template or macro-enabled. `Sheet.cells()`
  streams a worksheet's cells, optionally within a range, each with its value, type, formula,
  style index and metadata indexes.
- Values are read as Excel reads them: a subnormal number reads as 0; a cell with an error rich
  value behind its `vm` reads as the error Excel shows (`#SPILL!`, `#CALC!`, `#FIELD!`,
  `#CONNECT!`, `#BLOCKED!`, `#UNKNOWN!`, `#BUSY!`, `#TIMEOUT!`, by its `errorType`), whatever
  the cell caches, and type 14 as the number 0, as Excel saves it; errors are `ErrorCode`
  values (a `str` enum holding every error Excel defines). A file with an error code written
  as text that Excel refuses there (any but the first seven, `#GETTING_DATA` and `#BUSY!`), or
  a rich-value error type it refuses, is refused. Numbers too long or too large for any
  attribute are bounded, never an error.
- `Workbook.parts` gives every part of the package raw (bytes, stream or element tree), with
  its content type, and `Relationship` records for any part's relationships. `xlio.ns` holds
  the namespace, relationship-type and extension-URI constants, with extension URIs in the
  exact case Excel matches.
- `xlio.patch()` writes changed cells (cached formula results, constants, formulas, array
  ranges) into a copy of a workbook, every other part copied byte for byte; it takes Python
  values (`CellEdit`, `ArrayFill`). A cell given a new value loses its `vm`, since Excel shows
  the rich value a `vm` points to rather than the value beside it.
- `xlio.check()` checks that Excel would open a package as it is, without a repair, for a
  file handed to another reader: every part's size, checksum and DTD, then every worksheet
  read whole (structure, cells, drawing anchors, conditional formats). It costs about as much
  as reading every cell; `xlio.open()` stays lazy. It refuses the font sizes and font colors
  Excel refuses (a theme index past the theme's 12 colors), as the styles do. A known limit:
  formulas are not parsed, so a file Excel repairs for a formula that does not parse passes.
- `ReadLimits.max_total_bytes` counts each part once, however often it is read.
- `ReadLimits.max_cells` bounds the grid area one read of a worksheet covers, and
  `ReadLimits.UNTRUSTED` is a preset for files from people the caller does not trust.
- The package checks refuse an OLE compound file (an `.xls` workbook or a password-protected
  xlsx) and a binary `.xlsb` workbook with a message saying so, and check every part's stated
  size against the limits before anything is read.
- A sheet's structure: `Sheet.properties` (tab color, outline directions, code name),
  `Sheet.views` (zoom, gridlines and their color, headings, zeros, panes, selections, the
  active cell as a `CellRef`, taken as Excel takes it, with the text as written kept; a
  selection for a pane the view does not have is refused, as Excel refuses it),
  `Sheet.format` (default sizes) and `Sheet.columns`, read from the part's head up to its
  cells; `Sheet.rows`, gathered by a whole pass of `Sheet.cells()` (otherwise one
  pass that skips the cells); `Sheet.merges` and `Sheet.hyperlinks` with their external
  targets. A head record written after the cells is refused by the pass, as Excel refuses
  it.
- Sizes are read as Excel reads them: a column width is kept in points modulo 65,536 (a
  negative width wraps; NaN and infinity are 0), a row height is its magnitude and wraps past
  409.5, and the defaults Excel refuses (`defaultColWidth` of 300, `defaultRowHeight` of NaN)
  are refused; the text as written stays beside each (`raw_width`, `raw_height`). A width
  that is not a number is 0 and an empty height is 0; a column range Excel refuses (a `min`
  that is not a number, or past `max`) is refused.
- Attribute values Excel never writes are read as Excel reads them: a color's `rgb` of fewer
  than eight digits, or not hex, and a tint outside -1 to 1; a `textRotation` past 255; a
  rich-text run's font size (clamped to 1 to 409.55) and a differential format's (refused
  at 0 or past 409.55, also by `check()`). A cell whose style, cell-metadata or
  value-metadata index is past its table is refused, as Excel refuses it.
- `Workbook.theme` (color scheme, with Excel's theme index order, and font scheme),
  `Workbook.styles.indexed_colors` (a custom palette), `Workbook.views` and
  `Workbook.properties`.
- `unparsed` on `Workbook` and `Sheet`: every element xlio does not model, whole, so nothing is
  dropped without a trace.
- `Workbook.styles` in full: custom number formats, fonts, pattern and gradient fills,
  borders, cell formats with their font, fill, border, alignment, protection and named style
  resolved (an index that leads nowhere is `None`), named styles and differential formats.
  A font size is its magnitude, and a size outside 1 to 409 is refused, as Excel refuses it.
- Tables and filters: `Sheet.tables` (each table part with its columns, totals, calculated
  formulas, style options, filter and sort), `Sheet.autofilter` (values, date groups, custom,
  top 10, dynamic, color and icon filters, button settings) and `Sheet.sort_state`. An absent
  table style option is off, as Excel draws it.
- Table styles: `Styles.table_styles` (the workbook's own, with their differential formats
  resolved) and `Styles.table_style(name)`, which falls back to `xlio.presets`: Excel's 144
  built-in table and pivot table styles from ECMA-376, shipped as data with Ecma
  International's notice (`NOTICE`).
- Conditional formats: `Sheet.conditional_formats`, every rule typed (cell value, expression,
  text, blank, error, date, top and bottom, average, unique and duplicate rules, color scales,
  data bars, icon sets) with its formulas and its differential format, and each rule Excel 2010
  extended joined with its extension rule by `x14:id` (`Rule.x14`: data bar borders, direction,
  negative colors and axis, new and custom icon sets). Rules that exist only in the extension are
  formats of their own. The extension is read only under the URI Excel reads, compared exactly.
  `Sheet.extensions` lists every `extLst` entry with its URI as written. A format whose
  `sqref` is not a list of cells and ranges is dropped, as Excel drops it; one of a type
  outside the schema or with a `dxfId` past the table, which Excel repairs, is refused by
  `check()`.
- Drawings: `Sheet.drawing` gives the sheet's anchored objects in z-order (file order, across
  anchor kinds and through `AlternateContent`, whose branch is chosen as Excel is captured
  choosing it and kept raw beside), each with its anchor in EMU, kind, name, alternative text
  and hidden flag; pictures with their image part and content type, SVG image, crop, rotation,
  flips and outline, and a linked picture's target, which is never followed. A two-cell anchor
  that ends before it starts, by whole cells on either axis, is read as Excel repairs it (its
  corners swapped on that axis, `Anchor.repaired`) and refused by `check()`; offsets reversed
  within one cell are kept, as Excel keeps them.
  `Sheet.background` gives the background picture's part.
- VML: `Sheet.legacy_drawing` gives the VML part's bytes and its shapes (note boxes, form
  controls) read by a bounded extractor, since VML is often not well-formed XML: the anchor,
  visibility, cell, style, fill, stroke and shadow. `ReadLimits.max_vml_bytes` bounds it.
- Notes and threaded comments: `Sheet.comments` (cell, author, rich text),
  `Sheet.threaded_comments` (thread, author, date, resolved) and `Workbook.persons`.
- Printing: `Sheet.row_breaks` and `Sheet.column_breaks`, `Sheet.page_setup` (paper, orientation,
  scale and fitting, page numbers, margins, print options, headers and footers), and
  `Sheet.print_area` and `Sheet.print_titles` from the defined names scoped to the sheet by its
  place in `<sheets>` as written (`Sheet.position`: hidden sheets and sheets without a part
  count, as `localSheetId` and `activeTab` count them), each range as written, with the reason
  a part was not usable (`#REF!`, a formula, another sheet's range).
- `Workbook.external_links`: the parts of the workbook's links to other workbooks.
- Rich values: `Workbook.rich_values` gives the rich value behind each value-metadata index (a
  cell's `vm`): a picture placed in a cell as an `ImageValue` (its image part, alternative
  text, origin and sizing as stored; the cell itself caches `#VALUE!`), anything else (data
  types, errors) as a raw `RichValue`.
- Rich text: `Workbook.rich_strings` (runs with their fonts, phonetic guides and their
  properties) and `Sheet.cells(rich=True)`, which gives shared and inline strings their runs
  (`Cell.rich`); `Cell.phonetic` for the `ph` flag.
- Each rule read from Excel's behavior rests on a capture from Excel for Mac, with its probe
  (`CAPTURING.md`), and the tests read workbooks Excel saved.

[Unreleased]: https://github.com/sheetmill/xlio/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/sheetmill/xlio/releases/tag/v0.1.0
