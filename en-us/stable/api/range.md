# Range

### *class* Range(cell1=None, cell2=None, \*\*options)

Returns a Range object that represents a cell or a range of cells.

* **Parameters:**
  * **cell1** (*str* *|* *tuple* *[**int* *,* *int* *]*  *|* [*Range*](#xlwings.Range) *|* *None*) – Name of the range in the upper-left corner in A1 notation or as index-tuple or
    as name or as xw.Range object. It can also specify a range using the range
    operator (a colon), .e.g. ‘A1:B2’
  * **cell2** (*str* *|* *tuple* *[**int* *,* *int* *]*  *|* [*Range*](#xlwings.Range) *|* *None*) – Name of the range in the lower-right corner in A1 notation or as index-tuple or
    as name or as xw.Range object.

### Examples

```python
import xlwings as xw
sheet1 = xw.Book("MyBook.xlsx").sheets[0]

sheet1.range("A1")
sheet1.range("A1:C3")
sheet1.range((1,1))
sheet1.range((1,1), (3,3))
sheet1.range("NamedRange")

# Or using index/slice notation
sheet1["A1"]
sheet1["A1:C3"]
sheet1[0, 0]
sheet1[0:4, 0:4]
sheet1["NamedRange"]
```

#### add_hyperlink(address, text_to_display=None, screen_tip=None)

Adds a hyperlink to the specified Range (single Cell)

* **Parameters:**
  * **address** (*str*) – The address of the hyperlink.
  * **text_to_display** (*str* *|* *None*) – The text to be displayed for the hyperlink. Defaults to the hyperlink
    address.
  * **screen_tip** (*str* *|* *None*) – The screen tip to be displayed when the mouse pointer is paused over the
    hyperlink. Default is set to ‘<address> - Click once to follow. Click and
    hold to select this cell.’

#### Versionadded
Added in version 0.3.0.

#### *property* address *: str*

Returns a string value that represents the range reference.
Use `get_address()` to be able to provide parameters.

#### Versionadded
Added in version 0.9.0.

#### adjust_indent(amount)

Adjusts the indentation in a Range.

* **Parameters:**
  **amount** (*int*) – Number of spaces by which the indent is adjusted.
  Can be positive or negative.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.9.0.

#### autofill(destination, type_='fill_default')

Autofills the destination Range. Note that the destination Range must include
the origin Range.

* **Parameters:**
  * **destination** ([*Range*](#xlwings.Range)) – The origin.
  * **type** – One of the following strings: `"fill_copy"`, `"fill_days"`,
    `"fill_default"`, `"fill_formats"`, `"fill_months"`,
    `"fill_series"`, `"fill_values"`, `"fill_weekdays"`, `"fill_years"`,
    `"growth_trend"`, `"linear_trend"`, `"flash_fill`

#### Versionadded
Added in version 0.30.1.

#### autofit()

Autofits the width and height of all cells in the range.

* To autofit only the width of the columns use
  `myrange.columns.autofit()`
* To autofit only the height of the rows use
  `myrange.rows.autofit()`

#### Versionchanged
Changed in version 0.9.0.

#### clear()

Clears the content and the formatting of a Range.

#### clear_contents()

Clears the content of a Range but leaves the formatting.

#### clear_formats()

Clears the format of a Range but leaves the content.

#### Versionadded
Added in version 0.26.2.

#### *property* color *: tuple[int, int, int] | None*

Gets and sets the background color of the specified Range.

To set the color, either use an RGB tuple `(0, 0, 0)` or a hex string
like `#efefef` or an Excel color constant.
To remove the background, set the color to `None`, see Examples.

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = xw.sheets[0]
>>> sheet1.range('A1').color = (255, 255, 255)  # or '#ffffff'
>>> sheet1.range('A2').color
(255, 255, 255)
>>> sheet1.range('A2').color = None
>>> sheet1.range('A2').color is None
True
```

#### Versionadded
Added in version 0.3.0.

#### *property* column *: int*

Returns the number of the first column in the in the specified range. Read-only.

#### Versionadded
Added in version 0.3.5.

#### *property* column_width *: float | None*

Gets or sets the width of a Range.

The unit depends on the engine: **characters** on the classic, locally
installed xlwings, where one unit is the width of one character in the
Normal style (for proportional fonts, the character 0), and **points**
on xlwings Server and xlwings Lite.

If all columns in the Range have the same width, returns the width.
If columns in the Range have different widths, returns None.

In characters, column_width must be in the range 0 <= column_width <=
255. In points it only has to be positive.

Note: If the Range is outside the used range of the Worksheet, and columns in
the Range have different widths, returns the width of the first column.

#### Versionadded
Added in version 0.4.0.

#### *property* columns *: [RangeColumns](range_columns.md#xlwings.RangeColumns)*

Returns a `RangeColumns` object that represents the columns in the
specified range.

#### Versionadded
Added in version 0.9.0.

#### copy(destination=None)

Copy a range to a destination range or clipboard.

* **Parameters:**
  **destination** ([*Range*](#xlwings.Range) *|* *None*) – xlwings Range to which the specified range will be copied.
  If omitted, the range is copied to the clipboard.

#### copy_from(source_range, copy_type='all', skip_blanks=False, transpose=False)

A newer variant of copy that replaces copy/paste.

* **Parameters:**
  * **source_range** ([*Range*](#xlwings.Range))
  * **copy_type** (*str*) – One of “all”, “formats”, “formulas”, “link”, “values”
  * **skip_blanks** (*bool*)
  * **transpose** (*bool*)

#### copy_picture(appearance='screen', format='picture')

Copies the range to the clipboard as picture.

* **Parameters:**
  * **appearance** (*str*) – Either ‘screen’ or ‘printer’.
  * **format** (*str*) – Either ‘picture’ or ‘bitmap’.

#### Versionadded
Added in version 0.24.8.

#### *property* count *: int*

Returns the number of cells.

#### *property* current_region *: [Range](#xlwings.Range)*

This property returns a Range object representing a range bounded by (but not
including) any combination of blank rows and blank columns or the edges of the
worksheet. It corresponds to `Ctrl-*` on Windows and `Shift-Ctrl-Space` on
Mac.

#### delete(shift=None)

Deletes a cell or range of cells.

* **Parameters:**
  **shift** (*str* *|* *None*) – Use `left` or `up`. If omitted, Excel decides based on the shape of
  the range.

#### end(direction)

Returns a Range object that represents the cell at the end of the region that
contains the source range. Equivalent to pressing Ctrl+Up, Ctrl+down,
Ctrl+left, or Ctrl+right.

* **Parameters:**
  **direction** (*str*)

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = xw.sheets[0]
>>> sheet1.range('A1:B2').value = 1
>>> sheet1.range('A1').end('down')
<Range [Book1]Sheet1!$A$2>
>>> sheet1.range('B2').end('right')
<Range [Book1]Sheet1!$B$2>
```

#### Versionadded
Added in version 0.9.0.

#### expand(mode='table')

Expands the range according to the mode provided. Ignores empty top-left cells
(unlike `Range.end()`).

* **Parameters:**
  **mode** (*str*) – One of `'table'` (=down and right), `'down'`, `'right'`.

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = wb.sheets[0]
>>> sheet1.range('A1').value = [[None, 1], [2, 3]]
>>> sheet1.range('A1').expand().address
$A$1:$B$2
>>> sheet1.range('A1').expand('right').address
$A$1:$B$1
```

#### Versionadded
Added in version 0.9.0.

#### *property* formula *: str | list[str] | list[list[str]]*

Gets or sets the formula for the given Range.

#### *property* formula2 *: str | list[str] | list[list[str]]*

Gets or sets the formula2 for the given Range.

#### *property* formula_array *: str | None*

Gets or sets an  array formula for the given Range.

#### Versionadded
Added in version 0.7.1.

#### get_address(row_absolute=True, column_absolute=True, include_sheetname=False, external=False)

Returns the address of the range in the specified format. `address` can be
used instead if none of the defaults need to be changed.

* **Parameters:**
  * **row_absolute** (*bool*) – Set to True to return the row part of the reference as an absolute
    reference.
  * **column_absolute** (*bool*) – Set to True to return the column part of the reference as an absolute
    reference.
  * **include_sheetname** (*bool*) – Set to True to include the Sheet name in the address. Ignored if
    external=True.
  * **external** (*bool*) – Set to True to return an external reference with workbook and worksheet
    name.

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = wb.sheets[0]
>>> sheet1.range((1,1)).get_address()
'$A$1'
>>> sheet1.range((1,1)).get_address(False, False)
'A1'
>>> sheet1.range((1,1), (3,3)).get_address(True, False, True)
'Sheet1!A$1:C$3'
>>> sheet1.range((1,1), (3,3)).get_address(True, False, external=True)
'[Book1]Sheet1!A$1:C$3'
```

#### Versionadded
Added in version 0.2.3.

#### *async* get_color()

Fetch the fill color on demand, as an RGB tuple.

Returns `None` if the range has no fill.

Requires xlwings Lite.

#### *async* get_column_width()

Fetch the column width on demand, in points.

Returns `None` if the range’s columns aren’t all the same width.

Requires xlwings Lite.

#### *async* get_current_region()

Fetch the current region on demand.

The region bounded by blank rows and columns around this range, i.e.
`Ctrl-*`.

Requires xlwings Lite.

#### *async* get_formula()

Fetch formulas on demand.

The returned shape follows the same rules as reading `value`: a single
cell gives a string, a 1-by-n or n-by-1 range a flat list, and anything
else a nested list. `options(ndim=...)` applies as usual.

Requires xlwings Lite.

#### *async* get_formula_array()

Fetch the array formula for this range on demand.

A single string, or `None` if the range holds no array formula. Unlike
`get_formula()`, this isn’t a value per cell.

Requires xlwings Lite.

#### *async* get_height()

Fetch the range’s height in points, on demand.

Requires xlwings Lite.

#### *async* get_hyperlink()

Fetch this cell’s hyperlink address on demand.

The async equivalent of `hyperlink`: same result, including for cells
that use a `HYPERLINK()` formula. Raises if the cell has no hyperlink.

Requires xlwings Lite.

#### *async* get_left()

Fetch the distance from the sheet’s left edge, in points, on demand.

Requires xlwings Lite.

#### *async* get_merge_area()

Fetch the merged range containing this cell, on demand.

Returns this range itself if it isn’t part of a merged range.

Requires xlwings Lite.

#### *async* get_merge_cells()

Fetch whether this range contains merged cells, on demand.

`True` if the whole range is merged, `False` if none of it is, and
`None` if it’s only partly merged.

Requires xlwings Lite.

#### *async* get_number_format()

Fetch the number format on demand.

A single format string, or `None` if the range’s cells don’t share one.

Requires xlwings Lite.

#### *async* get_row_height()

Fetch the row height on demand, in points.

Returns `None` if the range’s rows aren’t all the same height.

Requires xlwings Lite.

#### *async* get_table()

Fetch the Table this range is part of, on demand.

Returns `None` if the range isn’t part of a table.

Requires xlwings Lite.

#### *async* get_top()

Fetch the distance from the sheet’s top edge, in points, on demand.

Requires xlwings Lite.

#### *async* get_value()

Fetch values on demand.

Requires xlwings Lite.

#### *async* get_width()

Fetch the range’s width in points, on demand.

Requires xlwings Lite.

#### *async* get_wrap_text()

Fetch the wrap text setting on demand.

`None` if the range doesn’t have a uniform wrap setting.

Requires xlwings Lite.

#### group(by=None)

Group rows or columns.

* **Parameters:**
  **by** (*str* *|* *None*) – “columns” or “rows”. Figured out automatically if the range is defined as
  ‘1:3’ or ‘A:C’, respectively.

#### *property* has_array *: bool*

`True` if the range is part of a legacy CSE Array formula
and `False` otherwise.

#### *property* height *: float*

Returns the height, in points, of a Range. Read-only.

#### Versionadded
Added in version 0.4.0.

#### *property* hyperlink *: str*

Returns the hyperlink address of the specified Range (single Cell only)

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = wb.sheets[0]
>>> sheet1.range('A1').value
'www.xlwings.org'
>>> sheet1.range('A1').hyperlink
'http://www.xlwings.org'
```

#### Versionadded
Added in version 0.3.0.

#### insert(shift, copy_origin='format_from_left_or_above')

Insert a cell or range of cells into the sheet.

* **Parameters:**
  * **shift** (*str*) – Use `right` or `down`.
  * **copy_origin** (*str*) – Use `format_from_left_or_above` or `format_from_right_or_below`.
    Note that copy_origin is only supported on Windows.

#### Versionchanged
Changed in version 0.30.3: `shift` is now a required argument.

#### *property* last_cell *: [Range](#xlwings.Range)*

Returns the bottom right cell of the specified range. Read-only.

### Example

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = wb.sheets[0]
>>> myrange = sheet1.range('A1:E4')
>>> myrange.last_cell.row, myrange.last_cell.column
(4, 5)
```

#### Versionadded
Added in version 0.3.5.

#### *property* left *: float*

Returns the distance, in points, from the left edge of column A to the left
edge of the range. Read-only.

#### Versionadded
Added in version 0.6.0.

#### merge(across=False)

Creates a merged cell from the specified Range object.

* **Parameters:**
  **across** (*bool*) – True to merge cells in each row of the specified Range as separate
  merged cells.

#### *property* merge_area *: [Range](#xlwings.Range)*

Returns a Range object that represents the merged Range containing the
specified cell. If the specified cell isn’t in a merged range, this property
returns the specified cell.

#### *property* merge_cells *: bool*

Returns `True` if the Range contains merged cells, otherwise `False`

#### *property* name *: [Name](name.md#xlwings.Name) | None*

Sets or gets the name of a Range.

#### Versionadded
Added in version 0.4.0.

#### *property* note *: [Note](note.md#xlwings.main.Note) | None*

Returns a Note object.
Before the introduction of threaded comments, a Note was called a Comment.

#### Versionadded
Added in version 0.24.2.

#### *property* number_format *: str*

Gets and sets the number_format of a Range.

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> sheet1 = wb.sheets[0]
>>> sheet1.range('A1').number_format
'General'
>>> sheet1.range('A1:C3').number_format = '0.00%'
>>> sheet1.range('A1:C3').number_format
'0.00%'
```

#### Versionadded
Added in version 0.2.3.

#### offset(row_offset=0, column_offset=0)

Returns a Range object that represents a Range that’s offset from the
specified range.

#### Versionadded
Added in version 0.3.0.

#### options(convert=None, \*\*options)

Allows you to set a converter and their options. Converters define how Excel
Ranges and their values are being converted both during reading and writing
operations. If no explicit converter is specified, the base converter is being
applied, see `converters`.

* **Parameters:**
  **convert** (*Any*) – A converter, e.g. `dict`, `np.array`, `pd.DataFrame`,
  `pd.Series`, defaults to default converter.

:keyword ndim: Number of dimensions.
:keyword numbers: Type of numbers, e.g. `int`.
:keyword dates: E.g. `datetime.date`, defaults to `datetime.datetime`.
:keyword empty: Transformation of empty cells.
:keyword transpose: Transpose values.
:keyword expand: One of `'table'`, `'down'`, `'right'`.
:keyword chunksize: Use a chunksize, e.g. `10000` to prevent timeout or
memory issues when reading or writing large amounts of data.
Works with all formats, including DataFrames, NumPy arrays,
and list of lists.
:keyword err_to_str: If `True`, will include cell errors such as `#N/A` as
strings. By default, they will be converted to `None`.
*New in version 0.28.0.*

For converter-specific options, see `converters`.

#### paste(paste=None, operation=None, skip_blanks=False, transpose=False)

Pastes a range from the clipboard into the specified range.

* **Parameters:**
  * **paste** (*str* *|* *None*) – One of `all_merging_conditional_formats`, `all`, `all_except_borders`,
    `all_using_source_theme`, `column_widths`, `comments`, `formats`,
    `formulas`, `formulas_and_number_formats`, `validation`, `values`,
    `values_and_number_formats`.
  * **operation** (*str* *|* *None*) – One of “add”, “divide”, “multiply”, “subtract”.
  * **skip_blanks** (*bool*) – Set to `True` to skip over blank cells
  * **transpose** (*bool*) – Set to `True` to transpose rows and columns.

#### *property* raw_value *: Any*

Gets and sets the values directly as delivered from/accepted by the engine that
s being used (`pywin32` or `appscript`) without going through any of
xlwings’ data cleaning/converting. This can be helpful if speed is an issue but
naturally will be engine specific, i.e. might remove the cross-platform
compatibility.

#### resize(row_size=None, column_size=None)

Resizes the specified Range

* **Parameters:**
  * **row_size** (*int* *|* *None*) – The number of rows in the new range (if None, the number of rows
    in the range is unchanged).
  * **column_size** (*int* *|* *None*) – The number of columns in the new range (if None, the number of
    columns in the range is unchanged).

#### Versionadded
Added in version 0.3.0.

#### *property* row *: int*

Returns the number of the first row in the specified range. Read-only.

#### Versionadded
Added in version 0.3.5.

#### *property* row_height *: float | None*

Gets or sets the height, in points, of a Range.
If all rows in the Range have the same height, returns the height.
If rows in the Range have different heights, returns None.

row_height must be in the range:
0 <= row_height <= 409.5

Note: If the Range is outside the used range of the Worksheet, and rows in the
Range have different heights, returns the height of the first row.

#### Versionadded
Added in version 0.4.0.

#### *property* rows *: [RangeRows](range_rows.md#xlwings.RangeRows)*

Returns a `RangeRows` object that represents the rows in the specified
range.

#### Versionadded
Added in version 0.9.0.

#### select()

Selects the range. Select only works on the active book.

#### Versionadded
Added in version 0.9.0.

#### *property* shape *: tuple[int, int]*

Tuple of Range dimensions.

#### Versionadded
Added in version 0.3.0.

#### *property* sheet *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the Sheet object to which the Range belongs.

#### Versionadded
Added in version 0.9.0.

#### *property* size *: int*

Number of elements in the Range.

#### Versionadded
Added in version 0.3.0.

#### *property* table *: [Table](table.md#xlwings.main.Table) | None*

Returns a Table object if the range is part of one, otherwise `None`.

#### Versionadded
Added in version 0.21.0.

#### to_pdf(path=None, layout=None, show=None, quality='standard')

Exports the range as PDF.

* **Parameters:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path where you want to store the pdf. Defaults to the address of the range
    in the same directory as the Excel file if the Excel file is stored and to
    the current working directory otherwise.
  * **layout** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – This argument requires xlwings `PRO`.
    Path to a PDF file on which the report will be printed. This is ideal for
    headers and footers as well as borderless printing of graphics/artwork. The
    PDF file either needs to have only 1 page (every report page uses the same
    layout) or otherwise needs the same amount of pages as the report (each
    report page is printed on the respective page in the layout PDF).
  * **show** (*bool* *|* *None*) – Once created, open the PDF file with the default application.
  * **quality** (*str*) – Quality of the PDF file. Can either be `'standard'` or `'minimum'`.

#### Versionadded
Added in version 0.26.2.

#### to_png(path=None)

Exports the range as PNG picture.

* **Parameters:**
  **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path where you want to store the picture. Defaults to the name of the range
  in the same directory as the Excel file if the Excel file is stored and to
  the current working directory otherwise.

#### Versionadded
Added in version 0.24.8.

#### *property* top *: float*

Returns the distance, in points, from the top edge of row 1 to the top edge of
the range. Read-only.

#### Versionadded
Added in version 0.6.0.

#### ungroup(by=None)

Ungroup rows or columns

* **Parameters:**
  **by** (*str* *|* *None*) – “columns” or “rows”. Figured out automatically if the range is defined as
  ‘1:3’ or ‘A:C’, respectively.

#### unmerge()

Separates a merged area into individual cells.

#### *property* value *: Any*

Gets and sets the values for the given Range. See `xlwings.Range.options`
about how to set options, e.g., to transform it into a DataFrame or how to set
a chunksize.

#### *property* width *: float*

Returns the width, in points, of a Range. Read-only.

#### Versionadded
Added in version 0.4.0.

#### *property* wrap_text *: bool | None*

Returns `True` if the wrap_text property is enabled and `False` if it’s
disabled. If not all cells have the same value in a range, on Windows it returns
`None` and on macOS `False`.

#### Versionadded
Added in version 0.23.2.
