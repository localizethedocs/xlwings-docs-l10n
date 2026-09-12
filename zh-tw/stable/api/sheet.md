# Sheet

### *class* Sheet(sheet=None, impl=None)

A sheet object is a member of the `sheets` collection:

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> wb.sheets[0]
<Sheet [Book1]Sheet1>
>>> wb.sheets['Sheet1']
<Sheet [Book1]Sheet1>
>>> wb.sheets.add()
<Sheet [Book1]Sheet2>
```

#### Versionchanged
Changed in version 0.9.0.

#### activate()

Activates the Sheet and returns it.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.9.0.

#### autofit(axis=None)

Autofits the width of either columns, rows or both on a whole Sheet.

* **Parameters:**
  **axis** (*str* *|* *None*) – To autofit rows, use `"rows"` or `"r"`. To autofit columns,
  use `"columns"` or `"c"`. To autofit rows and columns, provide
  no arguments.

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> wb.sheets['Sheet1'].autofit('c')
>>> wb.sheets['Sheet1'].autofit('r')
>>> wb.sheets['Sheet1'].autofit()
```

#### Versionadded
Added in version 0.2.3.

#### *property* book *: [Book](book.md#xlwings.Book)*

Returns the Book of the specified Sheet. Read-only.

#### *property* cells *: [Range](range.md#xlwings.Range)*

Returns a Range object that represents all the cells on the Sheet
(not just the cells that are currently in use).

#### Versionadded
Added in version 0.9.0.

#### *property* charts *: [Charts](charts.md#xlwings.main.Charts)*

See `Charts`

#### Versionadded
Added in version 0.9.0.

#### clear()

Clears the content and formatting of the whole sheet.

#### clear_contents()

Clears the content of the whole sheet but leaves the formatting.

#### clear_formats()

Clears the format of the whole sheet but leaves the content.

#### Versionadded
Added in version 0.26.2.

#### copy(before=None, after=None, name=None)

Copy a sheet to the current or a new Book. By default, it places the copied
sheet after all existing sheets in the current Book. Returns the copied sheet.

#### Versionadded
Added in version 0.22.0.

* **Parameters:**
  * **before** ([*Sheet*](#xlwings.Sheet) *|* *None*) – The sheet object before which you want to place the sheet
  * **after** ([*Sheet*](#xlwings.Sheet) *|* *None*) – The sheet object after which you want to place the sheet,
    by default it is placed after all existing sheets
  * **name** (*str* *|* *None*) – The sheet name of the copy

* **Returns:**
  The copied sheet
* **Return type:**
  [*Sheet*](#xlwings.Sheet)

### Examples

```python
# Create two books and add a value to the first sheet of the first book
first_book = xw.Book()
second_book = xw.Book()
first_book.sheets[0]['A1'].value = 'some value'

# Copy to same Book with the default location and name
first_book.sheets[0].copy()

# Copy to same Book with custom sheet name
first_book.sheets[0].copy(name='copied')

# Copy to second Book requires to use before or after
first_book.sheets[0].copy(after=second_book.sheets[0])
```

#### delete()

Deletes the Sheet.

#### Versionadded
Added in version 0.6.0.

#### *property* freeze_panes *: FreezePanes*

Interface to freeze/unfreeze panes.

### Examples

```pycon
>>> mysheet.freeze_panes.freeze_at("A1")
>>> mysheet.freeze_panes.freeze_at(mysheet["A1"])
>>> mysheet.freeze_panes.freeze_at("A:A")
>>> mysheet.freeze_panes.freeze_at("1:1")
>>> mysheet.freeze_panes.unfreeze()
```

#### *property* index *: int*

Returns the index of the Sheet (1-based as in Excel).

#### *async* load(values=None)

(Re)loads the sheet’s data from Excel on demand.

Like `Book.load`, this loads only *metadata* by default on an async book
and everything (including values) on a regular book. Pass `values=True`
to also load this sheet’s cell values.

Requires xlwings Lite.

#### *property* name *: str*

Gets or sets the name of the Sheet.

#### *property* names *: [Names](names.md#xlwings.main.Names)*

Returns a names collection that represents all the sheet-specific names
(names defined with the “SheetName!” prefix).

#### Versionadded
Added in version 0.9.0.

#### *property* page_setup *: [PageSetup](page_setup.md#xlwings.main.PageSetup)*

Returns a PageSetup object.

#### Versionadded
Added in version 0.24.2.

#### *property* pictures *: [Pictures](pictures.md#xlwings.main.Pictures)*

See `Pictures`

#### Versionadded
Added in version 0.9.0.

#### range(cell1, cell2=None)

Returns a Range object from the active sheet of the active book,
see `Range`.

#### Versionadded
Added in version 0.9.0.

#### render_template(\*\*data)

This method requires xlwings `PRO`.

Replaces all Jinja variables (e.g `{{ myvar }}`) in the sheet with the keyword
argument that has the same name. Following variable types are supported:

strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames,
PIL Image objects that have a filename and Matplotlib figures.

#### Versionadded
Added in version 0.22.0.

* **Parameters:**
  **data** (*Any*) – All key/value pairs that are used in the template.

### Examples

```pycon
>>> import xlwings as xw
>>> book = xw.Book()
>>> book.sheets[0]['A1:A2'].value = '{{ myvar }}'
>>> book.sheets[0].render_template(myvar='test')
```

#### select()

Selects the Sheet. Activates the book if it isn’t the active one.

#### Versionadded
Added in version 0.9.0.

#### *property* shapes *: [Shapes](shapes.md#xlwings.main.Shapes)*

See `Shapes`

#### Versionadded
Added in version 0.9.0.

#### *property* tables *: [Tables](tables.md#xlwings.main.Tables)*

See `Tables`

#### Versionadded
Added in version 0.21.0.

#### to_html(path=None)

Export a Sheet as HTML page.

* **Parameters:**
  **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path where you want to save the HTML file. Defaults to Sheet name in the
  current working directory.

#### Versionadded
Added in version 0.28.1.

#### to_pdf(path=None, layout=None, show=False, quality='standard')

Exports the sheet to a PDF file.

* **Parameters:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path to the PDF file, defaults to the name of the sheet in the same
    directory of the workbook. For unsaved workbooks, it defaults to the current
    working directory instead.
  * **layout** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – This argument requires xlwings `PRO`.
    Path to a PDF file on which the report will be printed. This is ideal for
    headers and footers as well as borderless printing of graphics/artwork. The
    PDF file either needs to have only 1 page (every report page uses the same
    layout) or otherwise needs the same amount of pages as the report (each
    report page is printed on the respective page in the layout PDF).
    *New in version 0.24.3.*
  * **show** (*bool*) – Once created, open the PDF file with the default application.
    *New in version 0.24.6.*
  * **quality** (*str*) – Quality of the PDF file. Can either be `'standard'` or `'minimum'`.
    *New in version 0.26.2.*

### Examples

```pycon
>>> wb = xw.Book()
>>> sheet = wb.sheets[0]
>>> sheet['A1'].value = 'PDF'
>>> sheet.to_pdf()
```

See also `xlwings.Book.to_pdf`

#### Versionadded
Added in version 0.22.3.

#### *property* used_range *: [Range](range.md#xlwings.Range)*

Used Range of Sheet.

#### Versionadded
Added in version 0.13.0.

#### *property* visible *: bool*

Gets or sets the visibility of the Sheet (bool).

#### Versionadded
Added in version 0.21.1.
