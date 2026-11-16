# API reference

## Reading

```{eval-rst}
.. autofunction:: xlio.open

.. autoclass:: xlio.Workbook
   :members:

.. autoclass:: xlio.Sheet
   :members:

.. autoclass:: xlio.Cell
   :members:

.. autoclass:: xlio.Formula
   :members:

.. autoclass:: xlio.ErrorCode
   :members:
   :undoc-members:

.. autoclass:: xlio.DefinedName
   :members:

.. autoclass:: xlio.CalcProperties
   :members:

.. autoclass:: xlio.CellMetadata
   :members:

.. autoclass:: xlio.RichValue
   :members:

.. autoclass:: xlio.ImageValue
   :members:

.. autoclass:: xlio.BookView
   :members:

.. autoclass:: xlio.WorkbookProperties
   :members:

.. autodata:: xlio.CellType

.. autodata:: xlio.SheetKind

.. autodata:: xlio.SheetState

.. autodata:: xlio.WorkbookKind
```

## A sheet's structure

```{eval-rst}
.. autoclass:: xlio.SheetProperties
   :members:

.. autoclass:: xlio.SheetView
   :members:

.. autoclass:: xlio.Pane
   :members:

.. autoclass:: xlio.Selection
   :members:

.. autoclass:: xlio.CellRef
   :members:

.. autodata:: xlio.PaneName

.. autoclass:: xlio.SheetFormat
   :members:

.. autoclass:: xlio.ColumnRange
   :members:

.. autoclass:: xlio.RowFormat
   :members:

.. autoclass:: xlio.Hyperlink
   :members:

.. autoclass:: xlio.Unparsed
   :members:
```

## Printing

```{eval-rst}
.. autoclass:: xlio.PageBreak
   :members:

.. autoclass:: xlio.PageSetup
   :members:

.. autoclass:: xlio.PageMargins
   :members:

.. autoclass:: xlio.PrintOptions
   :members:

.. autoclass:: xlio.HeaderFooter
   :members:

.. autoclass:: xlio.PrintArea
   :members:

.. autoclass:: xlio.PrintTitles
   :members:
```

## Theme, styles and colors

```{eval-rst}
.. autoclass:: xlio.Theme
   :members:

.. autoclass:: xlio.ColorScheme
   :members:

.. autoclass:: xlio.SchemeColor
   :members:

.. autoclass:: xlio.SchemeFonts
   :members:

.. autoclass:: xlio.Styles
   :members:

.. autoclass:: xlio.CellFormat
   :members:

.. autoclass:: xlio.NamedStyle
   :members:

.. autoclass:: xlio.DifferentialFormat
   :members:

.. autoclass:: xlio.Font
   :members:

.. autoclass:: xlio.PatternFill
   :members:

.. autoclass:: xlio.GradientFill
   :members:

.. autoclass:: xlio.GradientStop
   :members:

.. autodata:: xlio.Fill

.. autoclass:: xlio.Border
   :members:

.. autoclass:: xlio.BorderSide
   :members:

.. autoclass:: xlio.Alignment
   :members:

.. autoclass:: xlio.Protection
   :members:

.. autodata:: xlio.PatternType

.. autodata:: xlio.BorderStyle

.. autodata:: xlio.Underline

.. autoclass:: xlio.Color
   :members:
```

## Tables, table styles and filters

```{eval-rst}
.. autoclass:: xlio.Table
   :members:

.. autoclass:: xlio.TableColumn
   :members:

.. autoclass:: xlio.TableStyleInfo
   :members:

.. autoclass:: xlio.TableStyles
   :members:

.. autoclass:: xlio.TableStyle
   :members:

.. autoclass:: xlio.TableStyleElement
   :members:

.. autodata:: xlio.TableStyleElementType

.. automodule:: xlio.presets
   :members:

.. autoclass:: xlio.AutoFilter
   :members:

.. autoclass:: xlio.FilterColumn
   :members:

.. autoclass:: xlio.DateGroup
   :members:

.. autoclass:: xlio.CustomFilter
   :members:

.. autoclass:: xlio.Top10
   :members:

.. autoclass:: xlio.DynamicFilter
   :members:

.. autoclass:: xlio.ColorFilter
   :members:

.. autoclass:: xlio.IconFilter
   :members:

.. autoclass:: xlio.SortState
   :members:

.. autoclass:: xlio.SortCondition
   :members:
```

## Conditional formats

```{eval-rst}
.. autoclass:: xlio.ConditionalFormat
   :members:

.. autoclass:: xlio.Rule
   :members:

.. autoclass:: xlio.Cfvo
   :members:

.. autoclass:: xlio.ColorScale
   :members:

.. autoclass:: xlio.DataBar
   :members:

.. autoclass:: xlio.IconSet
   :members:

.. autoclass:: xlio.Extension
   :members:

.. autodata:: xlio.CfType

.. autodata:: xlio.CfvoType
```

## Drawings, notes and comments

```{eval-rst}
.. autoclass:: xlio.Drawing
   :members:

.. autoclass:: xlio.DrawingObject
   :members:

.. autoclass:: xlio.Anchor
   :members:

.. autoclass:: xlio.Marker
   :members:

.. autoclass:: xlio.Picture
   :members:

.. autoclass:: xlio.Alternate
   :members:

.. autodata:: xlio.AnchorKind

.. autodata:: xlio.ObjectKind

.. autoclass:: xlio.Vml
   :members:

.. autoclass:: xlio.VmlShape
   :members:

.. autoclass:: xlio.Comment
   :members:

.. autoclass:: xlio.ThreadedComment
   :members:

.. autoclass:: xlio.Person
   :members:
```

## Rich text

```{eval-rst}
.. autoclass:: xlio.RichString
   :members:

.. autoclass:: xlio.Run
   :members:

.. autoclass:: xlio.PhoneticRun
   :members:

.. autoclass:: xlio.PhoneticProperties
   :members:
```

## Parts and relationships

```{eval-rst}
.. autoclass:: xlio.Parts
   :members:

.. autoclass:: xlio.Relationship
   :members:

.. automodule:: xlio.ns
   :members:
```

## Writing

```{eval-rst}
.. autofunction:: xlio.patch

.. autoclass:: xlio.CellEdit
   :members:

.. autoclass:: xlio.ArrayFill
   :members:

.. autodata:: xlio.CellValue

.. autodata:: xlio.CachedValues

.. autodata:: xlio.Translate
```

## Limits and errors

```{eval-rst}
.. autofunction:: xlio.check

.. autoclass:: xlio.ReadLimits
   :members: UNTRUSTED

.. autoexception:: xlio.PackageError

.. autoexception:: xlio.LimitError
```
