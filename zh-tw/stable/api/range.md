# Range

### *class* Range(cell1=None, cell2=None, \*\*options)

Returns a Range object that represents a cell or a range of cells.

* **參數:**
  * **cell1** (*str* *|* *tuple* *[**int* *,* *int* *]*  *|* [*Range*](#xlwings.Range) *|* *None*) -- Name of the range in the upper-left corner in A1 notation or as index-tuple or
    as name or as xw.Range object. It can also specify a range using the range
    operator (a colon), .e.g. 'A1:B2'
  * **cell2** (*str* *|* *tuple* *[**int* *,* *int* *]*  *|* [*Range*](#xlwings.Range) *|* *None*) -- Name of the range in the lower-right corner in A1 notation or as index-tuple or
    as name or as xw.Range object.

### 範例

```python
# Assumes an existing workbook object named book.
sheet1 = book.sheets[0]

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

#### add_comment(text)

Add a plain-text threaded comment to one cell and return its handle.

Raises `TypeError` for non-text content and `ValueError` for empty text or a multi-cell range. Excel rejects a second comment on the same cell.

* **回傳型別:**
  [*Comment*](comment.md#xlwings.main.Comment)

#### add_hyperlink(address, text_to_display=None, screen_tip=None)

Adds a hyperlink to the specified Range (single Cell)

* **參數:**
  * **address** (*str*) -- The address of the hyperlink.
  * **text_to_display** (*str* *|* *None*) -- The text to be displayed for the hyperlink. Defaults to the hyperlink
    address.
  * **screen_tip** (*str* *|* *None*) -- The screen tip to be displayed when the mouse pointer is paused over the
    hyperlink. Default is set to '<address> - Click once to follow. Click and
    hold to select this cell.'

#### Versionadded
在 0.3.0 版被加入.

#### add_note(text)

Add a note to this single cell and return it.

Raises `TypeError` for non-text content and `ValueError` for empty text, a multi-cell range, or a cell that already has a note.

* **回傳型別:**
  [*Note*](note.md#xlwings.main.Note)

#### *property* address *: str*

Returns a string value that represents the range reference.
Use `get_address()` to be able to provide parameters.

#### Versionadded
在 0.9.0 版被加入.

#### adjust_indent(amount)

Adjusts the indentation in a Range.

* **參數:**
  **amount** (*int*) -- Number of spaces by which the indent is adjusted.
  Can be positive or negative.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.9.0 版被加入.

#### autofill(destination, type_='fill_default')

Autofills the destination Range. Note that the destination Range must include
the origin Range.

* **參數:**
  * **destination** ([*Range*](#xlwings.Range)) -- The origin.
  * **type** -- One of the following strings: `"fill_copy"`, `"fill_days"`,
    `"fill_default"`, `"fill_formats"`, `"fill_months"`,
    `"fill_series"`, `"fill_values"`, `"fill_weekdays"`, `"fill_years"`,
    `"growth_trend"`, `"linear_trend"`, `"flash_fill`

#### Versionadded
在 0.30.1 版被加入.

#### *property* autofilter *: [AutoFilter](autofilter.md#xlwings.AutoFilter)*

Returns the AutoFilter for this range.

The first row is treated as the header row and `field` arguments are one-based column positions relative to the range.

### 範例

```python
data = sheet["A1:C100"]
data.autofilter.apply_values(1, ["East", "West"])
data.autofilter.apply_comparison(3, "greater_than_or_equal", 10)
data.autofilter.clear(1)
data.autofilter.clear()
```

#### Versionadded
在 0.37.5 版被加入.

#### autofit()

Autofits the width and height of all cells in the range.

* To autofit only the width of the columns use
  `myrange.columns.autofit()`
* To autofit only the height of the rows use
  `myrange.rows.autofit()`

#### Versionchanged
在 0.9.0 版的變更.

#### *property* borders *: [Borders](borders.md#xlwings.main.Borders)*

Returns the [`Borders`](borders.md#xlwings.main.Borders) collection of the
range, which gives access to the eight individual
[`Border`](border.md#xlwings.main.Border) sides.

### 範例

```pycon
>>> myrange = sheet["A1:D10"]
>>> myrange.borders.line_style = "continuous"  # edges + inside borders
>>> myrange.borders["edge_bottom"].weight = "thick"
>>> myrange.borders.set("outside", line_style="double", color="#ff0000")
>>> myrange.borders.clear()
```

#### Versionadded
在 0.37.1 版被加入.

#### *property* characters *: [Characters](characters.md#xlwings.main.Characters)*

Returns the [`Characters`](characters.md#xlwings.main.Characters) object for the range, used to format part of the text.

#### clear()

Clears the content and the formatting of a Range.

#### clear_contents()

Clears the content of a Range but leaves the formatting.

#### clear_formats()

Clears the format of a Range but leaves the content.

#### Versionadded
在 0.26.2 版被加入.

#### *property* color *: tuple[int, int, int] | None*

Gets and sets the background color of the specified Range.

To set the color, either use an RGB tuple `(0, 0, 0)` or a hex string
like `#efefef` or an Excel color constant.
To remove the background, set the color to `None`, see Examples.

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
>>> sheet1.range('A1').color = (255, 255, 255)  # or '#ffffff'
>>> sheet1.range('A2').color
(255, 255, 255)
>>> sheet1.range('A2').color = None
>>> sheet1.range('A2').color is None
True
```

* **Setter type:**
  tuple[int, int, int] | int | str | None

#### Versionadded
在 0.3.0 版被加入.

#### *property* column *: int*

Returns the number of the first column in the in the specified range. Read-only.

#### Versionadded
在 0.3.5 版被加入.

#### *property* column_width *: float | None*

Gets or sets the width of a Range.

The unit depends on the engine: **characters** with classic xlwings, where one unit is the width of one character in the Normal style (for proportional fonts, the character 0), and **points** with xlwings Server and xlwings Lite.

If all columns in the Range have the same width, returns the width.
If columns in the Range have different widths, returns None.

In characters, column_width must be in the range 0 <= column_width <=
255. In points it only has to be positive.

Note: If the Range is outside the used range of the Worksheet, and columns in
the Range have different widths, returns the width of the first column.

* **Setter type:**
  float

#### Versionadded
在 0.4.0 版被加入.

#### *property* columns *: [RangeColumns](range_columns.md#xlwings.RangeColumns)*

Returns a `RangeColumns` object that represents the columns in the
specified range.

#### Versionadded
在 0.9.0 版被加入.

#### *property* comment *: [Comment](comment.md#xlwings.main.Comment) | None*

The cell's threaded comment. In xlwings Lite, use `await get_comment()`.

#### *property* conditional_formats *: [ConditionalFormats](conditional_formats.md#xlwings.ConditionalFormats)*

Returns the conditional-format rules for this range.

In xlwings Lite, use
`await sheet["A1:D10"].get_conditional_formats()` for inspection;
`sheet["A1:D10"].conditional_formats.clear()` remains available.

#### Versionadded
在 0.37.5 版被加入.

#### copy(destination=None)

Copy a range to a destination range or clipboard.

* **參數:**
  **destination** ([*Range*](#xlwings.Range) *|* *None*) -- xlwings Range to which the specified range will be copied.
  If omitted, the range is copied to the clipboard.

#### copy_from(source_range, copy_type='all', skip_blanks=False, transpose=False)

A newer variant of copy that replaces copy/paste.

* **參數:**
  * **source_range** ([*Range*](#xlwings.Range))
  * **copy_type** (*str*) -- One of "all", "formats", "formulas", "link", "values"
  * **skip_blanks** (*bool*)
  * **transpose** (*bool*)

#### copy_picture(appearance='screen', format='picture')

Copies the range to the clipboard as picture.

* **參數:**
  * **appearance** (*str*) -- Either 'screen' or 'printer'.
  * **format** (*str*) -- Either 'picture' or 'bitmap'.

#### Versionadded
在 0.24.8 版被加入.

#### *property* count *: int*

Returns the number of cells.

#### *property* current_region *: [Range](#xlwings.Range)*

This property returns a Range object representing a range bounded by (but not
including) any combination of blank rows and blank columns or the edges of the
worksheet. It corresponds to `Ctrl-*` on Windows and `Shift-Ctrl-Space` on
Mac.

#### *property* data_validation *: [DataValidation](data_validation.md#xlwings.main.DataValidation)*

Returns the data validation object for the range.

Use it to create, replace, or remove a validation rule. On xlwings Lite,
use [`Range.get_data_validation()`](#xlwings.Range.get_data_validation) instead.

### 範例

```python
sheet["A1:A10"].data_validation.set_list(["Open", "Closed"])
sheet["B1:B10"].data_validation.set_list(sheet["D1:D3"])
sheet["C1:C10"].data_validation.set_list(book.names["Statuses"])
sheet["E1:E10"].data_validation.set_whole_number("between", 1, 10)
sheet["F1:F10"].data_validation.set_custom("=F1<>""")
sheet["A1:A10"].data_validation.delete()
```

#### Versionadded
在 0.37.5 版被加入.

#### delete(shift=None)

Deletes a cell or range of cells.

* **參數:**
  **shift** (*str* *|* *None*) -- Use `left` or `up`. If omitted, Excel decides based on the shape of
  the range.

#### end(direction)

Returns a Range object that represents the cell at the end of the region that
contains the source range. Equivalent to pressing Ctrl+Up, Ctrl+down,
Ctrl+left, or Ctrl+right.

* **參數:**
  **direction** (*str*)
* **回傳型別:**
  [*Range*](#xlwings.Range)

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
>>> sheet1.range('A1:B2').value = 1
>>> sheet1.range('A1').end('down')
<Range [Book1]Sheet1!$A$2>
>>> sheet1.range('B2').end('right')
<Range [Book1]Sheet1!$B$2>
```

#### Versionadded
在 0.9.0 版被加入.

#### expand(mode='table')

Expands the range according to the mode provided. Ignores empty top-left cells
(unlike `Range.end()`).

For async workbook reads in xlwings Lite, use [`get_expanded_range`](#xlwings.Range.get_expanded_range).

* **參數:**
  **mode** (*str*) -- One of `'table'` (=down and right), `'down'`, `'right'`.
* **回傳型別:**
  [*Range*](#xlwings.Range)

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
>>> sheet1.range('A1').value = [[None, 1], [2, 3]]
>>> sheet1.range('A1').expand().address
$A$1:$B$2
>>> sheet1.range('A1').expand('right').address
$A$1:$B$1
```

#### Versionadded
在 0.9.0 版被加入.

#### find(text, , whole=False, direction='forward', order='rows', match_case=False)

Find the first matching cell or return `None`; in xlwings Lite, use `await rng.find(...)`.

Classic xlwings returns the result directly. The search starts at the first cell in the requested traversal direction and is restricted to this range.

* **參數:**
  * **text** (*str*) -- Text to find. An empty string is not allowed.
  * **whole** (*bool*) -- Match the entire cell rather than part of it.
  * **direction** (*str*) -- `"forward"` 或 `"backward"`。
  * **order** (*str*) -- Search by `"rows"` or `"columns"`.
  * **match_case** (*bool*) -- Whether matching is case-sensitive.
* **回傳型別:**
  [*Range*](#xlwings.Range) | *None* | *Awaitable*[[*Range*](#xlwings.Range) | *None*]

### 範例

With classic xlwings:

```python
# Assumes an existing workbook object named book.
sheet = book.sheets[0]
sheet["A1:A3"].value = [["North"], ["South"], ["North"]]
found = sheet["A1:A3"].find("South", whole=True)
print(found.address if found else None)  # $A$2
```

In xlwings Lite, await the same search:

```python
# Assumes an existing workbook object named book.
sheet = book.sheets["Sheet1"]
sheet["A1:A3"].value = [["North"], ["South"], ["North"]]
found = await sheet["A1:A3"].find("South", whole=True)
print(found.address if found else None)  # $A$2
```

#### *property* font *: [Font](font.md#xlwings.main.Font)*

Returns the [`Font`](font.md#xlwings.main.Font) object for the range, used to read or change font formatting.

#### *property* formula *: str | list[str] | list[list[str]]*

Gets or sets formulas using Excel's implicit-intersection formula semantics with classic xlwings.

Assign a formula string to fill the range, or a list or list of lists to write formulas to individual cells. Writing accepts the converters and shape options supported by [`Range.options`](#xlwings.Range.options).

For dynamic-array formulas with classic xlwings, use [`formula2`](#xlwings.Range.formula2). In xlwings Lite, both properties write through the same Office.js formula API, which supports dynamic arrays. Synchronous formula reads are not supported in xlwings Lite; use `await rng.get_formula()`.

* **Setter type:**
  Any

#### *property* formula2 *: str | list[str] | list[list[str]]*

Gets or sets formulas using Excel's dynamic-array formula semantics with classic xlwings.

Assign a formula string to fill the range, or a list or list of lists to write formulas to individual cells. Writing accepts the converters and shape options supported by [`Range.options`](#xlwings.Range.options).

In xlwings Lite, this property writes through the same Office.js formula API as [`formula`](#xlwings.Range.formula). Read formulas with `await rng.get_formula()`; there is no separate `get_formula2()` method.

* **Setter type:**
  Any

#### *property* formula_array *: str | None*

Gets or sets an  array formula for the given Range.

* **Setter type:**
  str

#### Versionadded
在 0.7.1 版被加入.

#### get_address(row_absolute=True, column_absolute=True, include_sheetname=False, external=False)

Returns the address of the range in the specified format. `address` can be
used instead if none of the defaults need to be changed.

* **參數:**
  * **row_absolute** (*bool*) -- Set to True to return the row part of the reference as an absolute
    reference.
  * **column_absolute** (*bool*) -- Set to True to return the column part of the reference as an absolute
    reference.
  * **include_sheetname** (*bool*) -- Set to True to include the Sheet name in the address. Ignored if
    external=True.
  * **external** (*bool*) -- Set to True to return an external reference with workbook and worksheet
    name.
* **回傳型別:**
  *str*

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
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
在 0.2.3 版被加入.

#### *async* get_color()

Fetch the fill color on demand, as an RGB tuple.

Returns `None` if the range has no fill.

Requires xlwings Lite.

* **回傳型別:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_colors()

Returns the fill color of every cell as a two-dimensional list.

Each entry is an RGB tuple, or `None` for no fill color. The result always has the range's row and column dimensions, including `[[color]]` for a single cell; conversion options such as `ndim` and `transpose` do not affect it.

Reads direct cell fills, excluding colors supplied by conditional formatting or table styles. For patterned fills, returns the background fill color, not the pattern color or rendered appearance.

A call supports at most 100,000 cells.

### 範例

```python
colors = await sheet["A1:B2"].get_colors()
# [[(255, 0, 0), None], [(255, 255, 255), (0, 0, 255)]]
```

See [`get_color`](#xlwings.Range.get_color) for a single range-level fill color.

Requires xlwings Lite.

#### Versionadded
在 0.37.5 版被加入.

* **回傳型別:**
  *list*[*list*[*tuple*[*int*, *int*, *int*] | *None*]]

#### *async* get_column_width()

Fetch the column width on demand, in points.

Returns `None` if the range's columns aren't all the same width.

Requires xlwings Lite.

* **回傳型別:**
  *float* | *None*

#### *async* get_comment()

Fetch the cell's threaded comment, or `None`.

Requires xlwings Lite.

* **回傳型別:**
  [*Comment*](comment.md#xlwings.main.Comment) | *None*

#### *async* get_conditional_formats()

Fetch the ordered conditional-format rules on demand.

Rules are returned from highest to lowest evaluation priority. Requires
xlwings Lite.

#### Versionadded
在 0.37.5 版被加入.

* **回傳型別:**
  [*ConditionalFormats*](conditional_formats.md#xlwings.ConditionalFormats)

#### *async* get_current_region()

Fetch the current region on demand.

The region bounded by blank rows and columns around this range, i.e.
`Ctrl-*`.

Requires xlwings Lite.

* **回傳型別:**
  [*Range*](#xlwings.Range)

#### *async* get_data_validation()

Fetch this range's data-validation rule for the range.

Returns `"none"` when the range has no validation,
`"mixed_criteria"` when only some cells have validation, and
`"inconsistent"` when cells have different rules.

Requires xlwings Lite.

#### Versionadded
在 0.37.5 版被加入.

* **回傳型別:**
  [*DataValidation*](data_validation.md#xlwings.main.DataValidation)

#### *async* get_expanded_range(mode='table')

Return the range expanded down, right, or in both directions.

Like [`expand`](#xlwings.Range.expand), ignores an empty top-left cell. `mode` is `"table"` (down and right), `"down"`, or `"right"`; the aliases accepted by `expand()` are also supported.

Requires xlwings Lite. Like other asynchronous getters, it flushes pending actions first, so expansion sees earlier writes in the same script. Captured writes in a dry run remain unapplied. When the intended dimensions are already known, construct the range directly instead.

### 範例

```python
header = await sheet["A6"].get_expanded_range("right")
header.font.bold = True
```

* **回傳型別:**
  [*Range*](#xlwings.Range)

#### *async* get_formula()

Fetch formulas on demand.

The returned shape follows the same rules as reading `value`: a single
cell gives a string, a 1-by-n or n-by-1 range a flat list, and anything
else a nested list. `options(ndim=...)` applies as usual.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *list*[*str*] | *list*[*list*[*str*]]

#### *async* get_formula_array()

Fetch the array formula for this range on demand.

A single string, or `None` if the range holds no array formula. Unlike
`get_formula()`, this isn't a value per cell.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *async* get_height()

Fetch the range's height in points, on demand.

Requires xlwings Lite.

* **回傳型別:**
  *float*

#### *async* get_horizontal_alignment()

Fetch the horizontal alignment on demand.

`None` if the cells in the range don't all have the same alignment.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *async* get_hyperlink()

Fetch this cell's hyperlink address on demand.

The async equivalent of `hyperlink`: same result, including for cells
that use a `HYPERLINK()` formula. Raises if the cell has no hyperlink.

Requires xlwings Lite.

* **回傳型別:**
  *str*

#### *async* get_left()

Fetch the distance from the sheet's left edge, in points, on demand.

Requires xlwings Lite.

* **回傳型別:**
  *float*

#### *async* get_merge_area()

Fetch the merged range containing this cell, on demand.

Returns this range itself if it isn't part of a merged range.

Requires xlwings Lite.

* **回傳型別:**
  [*Range*](#xlwings.Range)

#### *async* get_merge_cells()

Fetch whether this range contains merged cells, on demand.

`True` if the whole range is merged, `False` if none of it is, and
`None` if it's only partly merged.

Requires xlwings Lite.

* **回傳型別:**
  *bool* | *None*

#### *async* get_number_format()

Fetch the number format on demand.

A single format string, or `None` if the range's cells don't share one.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *async* get_row_height()

Fetch the row height on demand, in points.

Returns `None` if the range's rows aren't all the same height.

Requires xlwings Lite.

* **回傳型別:**
  *float* | *None*

#### get_special_cells(cell_type, value_type=None)

Return matching rectangular areas; in xlwings Lite, use `await rng.get_special_cells(...)`.

`cell_type` is `"blanks"`, `"constants"`, `"formulas"`, or `"visible"`. For constants or formulas, `value_type` may be `"numbers"`, `"text"`, `"logical"`, or `"errors"`; omitting it selects all value types. Returns an empty list when no cells match. Classic xlwings returns the list directly.

* **回傳型別:**
  *list*[[*Range*](#xlwings.Range)] | *Awaitable*[*list*[[*Range*](#xlwings.Range)]]

#### *async* get_table()

Fetch the Table this range is part of, on demand.

Returns `None` if the range isn't part of a table.

Requires xlwings Lite.

* **回傳型別:**
  [*Table*](table.md#xlwings.main.Table) | *None*

#### *async* get_top()

Fetch the distance from the sheet's top edge, in points, on demand.

Requires xlwings Lite.

* **回傳型別:**
  *float*

#### *async* get_value()

Reads the range's current values from Excel asynchronously.

Supports the same converters and options as [`value`](#xlwings.Range.value), including DataFrames, `ndim`, `transpose`, and `chunksize`. Without an explicit converter, a single cell returns a scalar, a single row or column returns a flat list, and a two-dimensional range returns a list of lists. Use `rng.options(ndim=2)` to always receive a list of rows.

### 範例

Inside an asynchronous script with a supplied `book` parameter:

```python
sheet = book.sheets["Sheet1"]
rows = await sheet["A1:B10"].options(ndim=2).get_value()
```

Requires xlwings Lite.

* **回傳型別:**
  *Any*

#### *async* get_vertical_alignment()

Fetch the vertical alignment on demand.

`None` if the cells in the range don't all have the same alignment.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *async* get_width()

Fetch the range's width in points, on demand.

Requires xlwings Lite.

* **回傳型別:**
  *float*

#### *async* get_wrap_text()

Fetch the wrap text setting on demand.

`None` if the range doesn't have a uniform wrap setting.

Requires xlwings Lite.

* **回傳型別:**
  *bool* | *None*

#### group(by=None)

Group rows or columns.

* **參數:**
  **by** (*str* *|* *None*) -- "columns" or "rows". Figured out automatically if the range is defined as
  '1:3' or 'A:C', respectively.

#### *property* has_array *: bool*

`True` if the range is part of a legacy CSE Array formula
and `False` otherwise.

#### *property* height *: float*

Returns the height, in points, of a Range. Read-only.

#### Versionadded
在 0.4.0 版被加入.

#### *property* horizontal_alignment *: Literal['general', 'left', 'center', 'right', 'fill', 'justify', 'center_across_selection', 'distributed'] | None*

Returns or sets the horizontal alignment of the range.

One of `'general'`, `'left'`, `'center'`, `'right'`, `'fill'`,
`'justify'`, `'center_across_selection'` or `'distributed'`. Returns
`None` if the cells in the range don't all have the same alignment.
The default is `'general'`, which right-aligns numbers and dates and
left-aligns text.

Reading this property synchronously requires classic xlwings. Setting it is also supported with xlwings Lite and xlwings Server; in xlwings Lite, read it via [`get_horizontal_alignment()`](#xlwings.Range.get_horizontal_alignment).

### 範例

```pycon
>>> sheet["A1"].horizontal_alignment = "center"
>>> sheet["A1"].horizontal_alignment
'center'
```

* **Setter type:**
  str

#### Versionadded
在 0.37.3 版被加入.

#### *property* hyperlink *: str*

Returns the hyperlink address of the specified Range (single Cell only)

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
>>> sheet1.range('A1').value
'www.xlwings.org'
>>> sheet1.range('A1').hyperlink
'http://www.xlwings.org'
```

#### Versionadded
在 0.3.0 版被加入.

#### insert(shift, copy_origin='format_from_left_or_above')

Insert a cell or range of cells into the sheet.

* **參數:**
  * **shift** (*str*) -- 使用 `right` 或 `down`。
  * **copy_origin** (*str*) -- Use `format_from_left_or_above` or `format_from_right_or_below`.
    Note that copy_origin is only supported on Windows.

#### Versionchanged
在 0.30.3 版的變更: `shift` is now a required argument.

#### *property* last_cell *: [Range](#xlwings.Range)*

Returns the bottom right cell of the specified range. Read-only.

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
>>> myrange = sheet1.range('A1:E4')
>>> myrange.last_cell.row, myrange.last_cell.column
(4, 5)
```

#### Versionadded
在 0.3.5 版被加入.

#### *property* left *: float*

Returns the distance, in points, from the left edge of column A to the left
edge of the range. Read-only.

#### Versionadded
在 0.6.0 版被加入.

#### merge(across=False)

Creates a merged cell from the specified Range object.

* **參數:**
  **across** (*bool*) -- True to merge cells in each row of the specified Range as separate
  merged cells.

#### *property* merge_area *: [Range](#xlwings.Range)*

Returns a Range object that represents the merged Range containing the
specified cell. If the specified cell isn't in a merged range, this property
returns the specified cell.

#### *property* merge_cells *: bool*

Returns `True` if the Range contains merged cells, otherwise `False`

#### move(destination)

Move this range to a new location, like cut and paste in Excel.

Values, formulas, and formatting move together, and formulas that reference the moved cells are updated to point to the new location. The source and destination may overlap, e.g., to shift a block a few columns to the left. Source cells that the moved block doesn't cover afterwards are left empty.

* **參數:**
  **destination** ([*Range*](#xlwings.Range)) -- The range whose top-left cell becomes the top-left cell of the moved block. The moved block keeps this range's size. The destination may be on another sheet of the same book.

### 範例

```python
# Shift the block O1:S145 three columns to the left, to L1:P145
sheet["O1:S145"].move(sheet["L1"])
```

#### Versionadded
在 0.37.7 版被加入.

#### *property* name *: [Name](name.md#xlwings.Name) | None*

Sets or gets the name of a Range.

* **Setter type:**
  str

#### Versionadded
在 0.4.0 版被加入.

#### *property* note *: [Note](note.md#xlwings.main.Note) | None*

Returns a Note object.
Before the introduction of threaded comments, a Note was called a Comment.

#### Versionadded
在 0.24.2 版被加入.

#### *property* number_format *: str*

Gets and sets the number_format of a Range.

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet1 = book.sheets[0]
>>> sheet1.range('A1').number_format
'General'
>>> sheet1.range('A1:C3').number_format = '0.00%'
>>> sheet1.range('A1:C3').number_format
'0.00%'
```

#### Versionadded
在 0.2.3 版被加入.

#### offset(row_offset=0, column_offset=0)

Returns a Range object that represents a Range that's offset from the
specified range.

#### Versionadded
在 0.3.0 版被加入.

* **回傳型別:**
  [*Range*](#xlwings.Range)

#### options(convert=None, \*\*options)

Allows you to set a converter and their options. Converters define how Excel
Ranges and their values are being converted both during reading and writing
operations. If no explicit converter is specified, the base converter is being
applied, see `converters`.

* **參數:**
  **convert** (*Any*) -- A converter, e.g. `dict`, `np.array`, `pd.DataFrame`,
  `pd.Series`, defaults to default converter.
* **關鍵字引數:**
  * **ndim** -- Number of dimensions.
  * **numbers** -- Type of numbers, e.g. `int`.
  * **dates** -- E.g. `datetime.date`, defaults to `datetime.datetime`.
  * **empty** -- Transformation of empty cells.
  * **transpose** -- Transpose values.
  * **expand** -- One of `'table'`, `'down'`, `'right'`.
  * **chunksize** -- Number of rows per chunk when reading or writing large
    amounts of data, e.g. `10000`. Large ranges are chunked
    automatically (reads above 4,000,000 cells, writes above
    100,000 cells with classic xlwings and the remote engines); set
    `chunksize` explicitly to tune the row count, e.g. to prevent
    timeout, memory or payload-size issues, or `chunksize=None` to
    disable chunking. Works with all formats, including DataFrames,
    NumPy arrays, and list of lists.
  * **err_to_str** -- If `True`, will include cell errors such as `#N/A` as
    strings. By default, they will be converted to `None`.
    *New in version 0.28.0.*
* **回傳型別:**
  [*Range*](#xlwings.Range)

Without an explicit converter, a single cell returns a scalar, a single row or column returns a flat list, and a two-dimensional range returns a list of lists. Use `ndim=2` when code needs a list of rows regardless of the range shape: a single cell becomes `[[value]]`, one row becomes `[[a, b]]`, and one column becomes `[[a], [b]]`.

Options apply to both [`value`](#xlwings.Range.value) and, in xlwings Lite, [`get_value`](#xlwings.Range.get_value). For example, `rng.options(ndim=2).value` and `await rng.options(ndim=2).get_value()` both request a two-dimensional result. Converter-specific options are described in `converters`.

#### paste(paste=None, operation=None, skip_blanks=False, transpose=False)

Pastes a range from the clipboard into the specified range.

* **參數:**
  * **paste** (*str* *|* *None*) -- One of `all_merging_conditional_formats`, `all`, `all_except_borders`,
    `all_using_source_theme`, `column_widths`, `comments`, `formats`,
    `formulas`, `formulas_and_number_formats`, `validation`, `values`,
    `values_and_number_formats`.
  * **operation** (*str* *|* *None*) -- One of "add", "divide", "multiply", "subtract".
  * **skip_blanks** (*bool*) -- Set to `True` to skip over blank cells
  * **transpose** (*bool*) -- Set to `True` to transpose rows and columns.

#### *property* raw_value *: Any*

Gets and sets the values directly as delivered from/accepted by the engine that
s being used (`pywin32` or `appscript`) without going through any of
xlwings' data cleaning/converting. This can be helpful if speed is an issue but
naturally will be engine specific, i.e. might remove the cross-platform
compatibility.

#### remove_duplicates(columns, has_headers=False)

Remove duplicate rows within this range, keeping the first occurrence.

* **參數:**
  * **columns** (*int* *|* *Sequence* *[**int* *]*) -- One-based column positions within this range used to identify duplicates.
  * **has_headers** (*bool*) -- Keep the first row as a header.

Only cells inside this range are shifted. Ranges intersecting an Excel table are not supported. On macOS this method isn't supported and raises `NotImplementedError`.

#### replace_all(old, new, , whole=False, match_case=False)

Replace matching text within this range.

An empty replacement string is allowed; an empty search string is not.

### 範例

```python
# Assumes an existing workbook object named book.
sheet = book.sheets[0]
sheet["A1:A2"].value = [["Draft"], ["Draft report"]]
sheet["A1:A2"].replace_all("Draft", "Final")
print(sheet["A2"].value)  # Final report
```

In xlwings Lite, inside an asynchronous script with a supplied `book` parameter:

```python
# Assumes an existing workbook object named book.
sheet = book.sheets["Sheet1"]
sheet["A1:A2"].value = [["Draft"], ["Draft report"]]
sheet["A1:A2"].replace_all("Draft", "Final")
print(await sheet["A2"].get_value())  # Final report
```

#### resize(row_size=None, column_size=None)

Resizes the specified Range

* **參數:**
  * **row_size** (*int* *|* *None*) -- The number of rows in the new range (if None, the number of rows
    in the range is unchanged).
  * **column_size** (*int* *|* *None*) -- The number of columns in the new range (if None, the number of
    columns in the range is unchanged).
* **回傳型別:**
  [*Range*](#xlwings.Range)

#### Versionadded
在 0.3.0 版被加入.

#### *property* row *: int*

Returns the number of the first row in the specified range. Read-only.

#### Versionadded
在 0.3.5 版被加入.

#### *property* row_height *: float | None*

Gets or sets the height, in points, of a Range.
If all rows in the Range have the same height, returns the height.
If rows in the Range have different heights, returns None.

row_height must be in the range:
0 <= row_height <= 409.5

Note: If the Range is outside the used range of the Worksheet, and rows in the
Range have different heights, returns the height of the first row.

* **Setter type:**
  float

#### Versionadded
在 0.4.0 版被加入.

#### *property* rows *: [RangeRows](range_rows.md#xlwings.RangeRows)*

Returns a `RangeRows` object that represents the rows in the specified
range.

#### Versionadded
在 0.9.0 版被加入.

#### select()

Selects the range. Select only works on the active book.

#### Versionadded
在 0.9.0 版被加入.

#### set_colors(colors)

Set direct fill colors cell by cell without changing cell contents.

`colors` must be a two-dimensional matrix exactly matching the range's shape, including `[[color]]` for one cell. An RGB tuple or list or a `#RRGGBB` hex string. `None` removes the fill; `...` leaves that cell's existing fill unchanged. Conversion options do not change the required matrix shape.

### 範例

```python
sheet["A1:C1"].set_colors([[(0, 128, 0), ..., None]])
```

#### Versionadded
在 0.37.5 版被加入.

#### *property* shape *: tuple[int, int]*

Tuple of Range dimensions.

#### Versionadded
在 0.3.0 版被加入.

#### *property* sheet *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the Sheet object to which the Range belongs.

#### Versionadded
在 0.9.0 版被加入.

#### *property* size *: int*

Number of elements in the Range.

#### Versionadded
在 0.3.0 版被加入.

#### sort(keys, ascending=True, has_headers=False)

Sort the rows of this rectangular range by one or more columns.

* **參數:**
  * **keys** (*int* *|* *Sequence* *[**int* *]*) -- One-based column positions within this range, in priority order.
  * **ascending** (*bool* *|* *Sequence* *[**bool* *]*) -- One direction for every key, or one boolean per key.
  * **has_headers** (*bool*) -- Whether to keep the first row in place as a header.

The range is not expanded to adjacent data. Ranges intersecting an Excel table are not supported; use [`Table.sort`](table.md#xlwings.main.Table.sort) instead.

### 範例

```python
sheet["A1:D20"].sort([2, 1], [False, True], has_headers=True)
```

#### *property* table *: [Table](table.md#xlwings.main.Table) | None*

Returns a Table object if the range is part of one, otherwise `None`.

#### Versionadded
在 0.21.0 版被加入.

#### to_pdf(path=None, layout=None, show=None, quality='standard')

Exports the range as PDF.

* **參數:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Path where you want to store the pdf. Defaults to the address of the range
    in the same directory as the Excel file if the Excel file is stored and to
    the current working directory otherwise.
  * **layout** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- This argument requires xlwings `PRO`.
    Path to a PDF file on which the report will be printed. This is ideal for
    headers and footers as well as borderless printing of graphics/artwork. The
    PDF file either needs to have only 1 page (every report page uses the same
    layout) or otherwise needs the same amount of pages as the report (each
    report page is printed on the respective page in the layout PDF).
  * **show** (*bool* *|* *None*) -- Once created, open the PDF file with the default application.
  * **quality** (*str*) -- Quality of the PDF file. Can either be `'standard'` or `'minimum'`.
* **回傳型別:**
  *str*

#### Versionadded
在 0.26.2 版被加入.

#### to_png(path=None)

Exports the range as PNG picture.

In xlwings Lite, `path` is required and the export is queued like other workbook actions: the file exists only after the script returns or you call `await book.flush()`. Flush before reading the file within the same script.

* **參數:**
  **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Path where you want to store the picture. Defaults to the name of the range
  in the same directory as the Excel file if the Excel file is stored and to
  the current working directory otherwise.

### 範例

Inside an asynchronous xlwings Lite script with a supplied `book` parameter:

```python
book.sheets[0]["A1:D10"].to_png("/tmp/range.png")
await book.flush()
png_bytes = Path("/tmp/range.png").read_bytes()
```

#### Versionadded
在 0.24.8 版被加入.

#### *property* top *: float*

Returns the distance, in points, from the top edge of row 1 to the top edge of
the range. Read-only.

#### Versionadded
在 0.6.0 版被加入.

#### ungroup(by=None)

Ungroup rows or columns

* **參數:**
  **by** (*str* *|* *None*) -- "columns" or "rows". Figured out automatically if the range is defined as
  '1:3' or 'A:C', respectively.

#### unmerge()

Separates a merged area into individual cells.

#### *property* value *: Any*

Gets and sets the values for the given range.

Without an explicit converter, a single cell returns a scalar, a single row or column returns a flat list, and a two-dimensional range returns a list of lists. Use [`options`](#xlwings.Range.options) to control dimensionality, convert values to a DataFrame, or set a chunk size.

In xlwings Lite, reads use the cell values loaded for the script. A script annotated with [`BookAsync`](book_async.md#xlwings.BookAsync) does not preload cell values: use `await rng.get_value()` to read them.

#### *property* vertical_alignment *: Literal['top', 'center', 'bottom', 'justify', 'distributed'] | None*

Returns or sets the vertical alignment of the range.

One of `'top'`, `'center'`, `'bottom'`, `'justify'` or `'distributed'`.
Returns `None` if the cells in the range don't all have the same
alignment. The default is `'bottom'`.

Reading this property synchronously requires classic xlwings. Setting it is also supported with xlwings Lite and xlwings Server; in xlwings Lite, read it via [`get_vertical_alignment()`](#xlwings.Range.get_vertical_alignment).

### 範例

```pycon
>>> sheet["A1"].vertical_alignment = "top"
>>> sheet["A1"].vertical_alignment
'top'
```

* **Setter type:**
  str

#### Versionadded
在 0.37.3 版被加入.

#### *property* width *: float*

Returns the width, in points, of a Range. Read-only.

#### Versionadded
在 0.4.0 版被加入.

#### *property* wrap_text *: bool | None*

Returns `True` if the wrap_text property is enabled and `False` if it's
disabled. If not all cells have the same value in a range, on Windows it returns
`None` and on macOS `False`.

* **Setter type:**
  bool

#### Versionadded
在 0.23.2 版被加入.
