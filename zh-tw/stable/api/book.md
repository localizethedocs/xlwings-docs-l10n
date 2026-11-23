# Book

### *class* Book(fullname=None, update_links=None, read_only=None, format=None, password=None, write_res_password=None, ignore_read_only_recommended=None, origin=None, delimiter=None, editable=None, notify=None, converter=None, add_to_mru=None, local=None, corrupt_load=None, json=None, mode=None, engine=None, \*\*kwargs)

A book object is a member of the `books` collection:

```pycon
>>> import xlwings as xw
>>> xw.books[0]
<Book [Book1]>
```

The easiest way to connect to a book is offered by `xw.Book`: it looks for the
book in all app instances and returns an error, should the same book be open in
multiple instances. To connect to a book in the active app instance, use
`xw.books` and to refer to a specific app, use:

```pycon
>>> app = xw.App()  # or xw.apps[10559] (get the PIDs via xw.apps.keys())
>>> app.books['Book1']
```

|                    | xw.Book                            | xw.books                                 |
|--------------------|------------------------------------|------------------------------------------|
| New book           | `xw.Book()`                        | `xw.books.add()`                         |
| Unsaved book       | `xw.Book('Book1')`                 | `xw.books['Book1']`                      |
| Book by (full)name | `xw.Book(r'C:/path/to/file.xlsx')` | `xw.books.open(r'C:/path/to/file.xlsx')` |
* **參數:**
  * **fullname** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Full path or name (incl. xlsx, xlsm etc.) of existing workbook or name of an
    unsaved workbook. Without a full path, it looks for the file in the current
    working directory.
  * **update_links** (*bool* *|* *None*) -- If this argument is omitted, the user is prompted to specify how links will be
    updated
  * **read_only** (*bool* *|* *None*) -- True to open workbook in read-only mode
  * **format** (*str* *|* *None*) -- If opening a text file, this specifies the delimiter character
  * **password** (*str* *|* *None*) -- Password to open a protected workbook
  * **write_res_password** (*str* *|* *None*) -- Password to write to a write-reserved workbook
  * **ignore_read_only_recommended** (*bool* *|* *None*) -- Set to `True` to mute the read-only recommended message
  * **origin** (*int* *|* *None*) -- For text files only. Specifies where it originated. Use Platform constants.
  * **delimiter** (*str* *|* *None*) -- If format argument is 6, this specifies the delimiter.
  * **editable** (*bool* *|* *None*) -- This option is only for legacy Microsoft Excel 4.0 addins.
  * **notify** (*bool* *|* *None*) -- Notify the user when a file becomes available If the file cannot be opened in
    read/write mode.
  * **converter** (*int* *|* *None*) -- The index of the first file converter to try when opening the file.
  * **add_to_mru** (*bool* *|* *None*) -- Add this workbook to the list of recently added workbooks.
  * **local** (*bool* *|* *None*) -- If `True`, saves files against the language of Excel, otherwise against the
    language of VBA. Not supported on macOS.
  * **corrupt_load** (*int* *|* *None*) -- Can be one of xlNormalLoad, xlRepairFile or xlExtractData.
    Not supported on macOS.
  * **json** (*dict* *[**str* *,* *Any* *]*  *|* *None*) -- A JSON object as delivered by the MS Office Scripts or Google Apps Script
    xlwings module but in a deserialized form, i.e., as dictionary.
    *New in version 0.26.0.*
  * **mode** (*str* *|* *None*) -- Either `"i"` (interactive (default)) or `"r"` (read). In interactive mode,
    xlwings opens the workbook in Excel, i.e., Excel needs to be installed. In read
    mode, xlwings reads from the file directly, without requiring Excel to be
    installed. Read mode requires xlwings `PRO`.
    *New in version 0.28.0.*

#### activate(steal_focus=False)

Activates the book.

* **參數:**
  **steal_focus** (*bool*) -- If True, make frontmost window and hand over focus from Python to Excel.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj) of the engine being
used.

In xlwings Lite, this is the remote backend's internal workbook data dictionary. It does not expose the Office.js API; use the public xlwings methods and properties for workbook operations.

#### Versionadded
在 0.9.0 版被加入.

#### *property* app *: [App](app.md#xlwings.App)*

Returns an app object that represents the creator of the book.

#### Versionadded
在 0.9.0 版被加入.

#### *classmethod* caller()

References the calling book when the Python function is called from Excel via
`RunPython`. Pack it into the function being called from Excel, e.g.:

```python
import xlwings as xw

 def my_macro():
    wb = xw.Book.caller()
    wb.sheets[0].range('A1').value = 1
```

To be able to easily invoke such code from Python for debugging, use
`xw.Book.set_mock_caller()`.

#### Versionadded
在 0.3.0 版被加入.

* **回傳型別:**
  [*Book*](#xlwings.Book)

#### close()

Closes the book without saving it.

#### Versionadded
在 0.1.1 版被加入.

#### *property* comments *: [Comments](comments.md#xlwings.main.Comments)*

Threaded comments in this workbook. In xlwings Lite, use `await get_comments()`.

#### *async* flush()

Flushes all pending actions to Excel and lets the Output pane update.

In xlwings Lite, writes and other workbook changes are queued until the script returns or you call `await book.flush()`. Asynchronous getters such as `await rng.get_value()` flush pending actions automatically before they read from Excel, so they always see the script's earlier changes, including a sheet that was just added. The same applies to `await book.load()` and `await sheet.load()`, e.g., to access the columns of a table that was just added. You only need to call `await book.flush()` explicitly in the following cases:

- **Showing progress during a long-running script:** a change that no read depends on, such as a status message written to a cell, only appears in Excel after the script returns unless you flush it. Flushing also lets the Output pane show what you printed so far.
- **Reading a file that an action creates:** actions such as [`Range.to_png`](range.md#xlwings.Range.to_png) write their file only when they are dispatched, and reading a file doesn't flush pending actions. Flush before reading the file within the same script.
- **Choosing what's applied if the script fails:** if a script raises an exception, its pending actions are discarded. Everything flushed before the exception, explicitly or by an asynchronous getter, stays applied. Flush to apply the changes up to a certain point, for example, before a step that may fail.

Flushing does not refresh cell values previously loaded for synchronous [`Range.value`](range.md#xlwings.Range.value) reads. Use `await rng.get_value()` to read the changed cells, or `await book.load(values=True)` to reload values before using `.value`.

### 範例

Inside an asynchronous script with a supplied `book` parameter:

```python
status = book.sheets["Sheet1"]["A1"]
status.value = "Calculating..."
await book.flush()  # Shows the status before the slow step
run_slow_calculation()
status.value = "Done"
```

Requires xlwings Lite.

#### Versionadded
在 0.35.0 版被加入.

#### *property* fullname *: str*

Returns the name of the object, including its path on disk, as a string.
Read-only String.

#### *async* get_comments()

Fetch the workbook's threaded comments.

Requires xlwings Lite.

* **回傳型別:**
  [*Comments*](comments.md#xlwings.main.Comments)

#### *async* get_selection()

Returns the selected cells as Range, fetched live from Excel.

Requires xlwings Lite.

#### Versionadded
在 0.35.0 版被加入.

* **回傳型別:**
  [*Range*](range.md#xlwings.Range) | *None*

#### json()

Returns a JSON serializable object as expected by the MS Office Scripts or
Google Apps Script xlwings module. Only available with book objects that have
been instantiated via `xw.Book(json=...)`.

#### Versionadded
在 0.26.0 版被加入.

* **回傳型別:**
  *dict*[*str*, *Any*]

#### *async* load(values=None)

(Re)loads the book's data from Excel on demand.

On an async book (see `xw.BookAsync`), this loads only *metadata*
(tables, pictures, names) by default — not cell values — since
bulk-loading values would defeat the point of the async API. Pass
`values=True` to also snapshot all cell values, after which sync
`.value` reads work again. On a regular book, everything including
values is loaded regardless.

Requires xlwings Lite.

* **回傳型別:**
  [*Book*](#xlwings.Book)

#### macro(name)

Runs a Sub or Function in Excel VBA.

* **參數:**
  **name** (*str*) -- `'Module1.MyMacro'` or `'MyMacro'`
* **回傳型別:**
  *Macro*

### 範例

```vb.net
Function MySum(x, y)
    MySum = x + y
End Function
```

can be accessed like this:

```pycon
>>> import xlwings as xw
>>> wb = xw.books.active
>>> my_sum = wb.macro('MySum')
>>> my_sum(1, 2)
3
```

See also: `App.macro`

#### Versionadded
在 0.7.1 版被加入.

#### *property* name *: str*

Returns the name of the book as str.

#### *property* names *: [Names](names.md#xlwings.main.Names)*

Returns a names collection that represents all the names in the specified book
(including all sheet-specific names).

#### Versionchanged
在 0.9.0 版的變更.

#### *property* notes *: [Notes](notes.md#xlwings.main.Notes)*

Notes in this workbook. xlwings Lite includes the worksheets loaded for this script.

#### render_template(\*\*data)

This method requires xlwings `PRO`.

Replaces all Jinja variables (e.g `{{ myvar }}`) in the book
with the keyword argument of the same name.

#### Versionadded
在 0.25.0 版被加入.

* **參數:**
  **data** (*Any*) -- All key/value pairs that are used in the template.

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.sheets[0]['A1:A2'].value = '{{ myvar }}'
>>> book.render_template(myvar='test')
```

#### save(path=None, password=None)

Saves the Workbook. If a path is provided, this works like SaveAs() in
Excel. If no path is specified and if the file hasn't been saved previously,
it's saved in the current working directory with the current filename.
Existing files are overwritten without prompting. To change the file type,
provide the appropriate extension, e.g. to save `myfile.xlsx` in the `xlsb`
format, provide `myfile.xlsb` as path.

* **參數:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Path where you want to save the Book.
  * **password** (*str* *|* *None*) -- Protection password with max. 15 characters
    *New in version 0.25.1.*

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.save()
>>> book.save(r'C:\path\to\new_file_name.xlsx')
```

#### NOTE
On xlwings Server and xlwings Lite, the workbook is saved in place
after your script returns rather than while it runs. `path` and
`password` aren't supported there: the book can only be saved in
place, and there's no way to set a password.

#### Versionadded
在 0.3.1 版被加入.

#### *property* selection *: [Range](range.md#xlwings.Range) | None*

Returns the selected cells as Range.

#### Versionadded
在 0.9.0 版被加入.

#### set_mock_caller()

Sets the Excel file which is used to mock `xw.Book.caller()` when the code is
called from Python and not from Excel via `RunPython`.

### 範例

```python
# This code runs unchanged from Excel via RunPython and from Python directly
import os
import xlwings as xw

def my_macro():
sheet = xw.Book.caller().sheets[0]
sheet.range('A1').value = 'Hello xlwings!'

if __name__ == '__main__':
xw.Book('file.xlsm').set_mock_caller()
my_macro()
```

#### Versionadded
在 0.3.1 版被加入.

#### *property* sheet_names *: list[str]*

* **回傳:**
  List of sheet names in order of appearance.

#### Versionadded
在 0.28.1 版被加入.

#### *property* sheets *: [Sheets](sheets.md#xlwings.main.Sheets)*

Returns a sheets collection that represents all the sheets in the book.

#### Versionadded
在 0.9.0 版被加入.

#### *async* sync()

#### Deprecated
在 0.35.0 版之後被棄用: 請改用 `flush`。

#### to_pdf(path=None, include=None, exclude=None, layout=None, exclude_start_string='#', show=False, quality='standard')

Exports the whole Excel workbook or a subset of the sheets to a PDF file.
If you want to print hidden sheets, you will need to list them explicitely
under `include`.

* **參數:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Path to the PDF file, defaults to the same name as the workbook, in the same
    directory. For unsaved workbooks, it defaults to the current working
    directory instead.
  * **include** (*int* *|* *str* *|* *list* *[**int* *|* *str* *]*  *|* *None*) -- Which sheets to include: provide a selection of sheets in the form of sheet
    indices (1-based like in Excel) or sheet names. Can be an int/str for a
    single sheet or a list of int/str for multiple sheets.
  * **exclude** (*int* *|* *str* *|* *list* *[**int* *|* *str* *]*  *|* *None*) -- Which sheets to exclude: provide a selection of sheets in the form of sheet
    indices (1-based like in Excel) or sheet names. Can be an int/str for a
    single sheet or a list of int/str for multiple sheets.
  * **layout** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- This argument requires xlwings `PRO`.
    Path to a PDF file on which the report will be printed. This is ideal for
    headers and footers as well as borderless printing of graphics/artwork. The
    PDF file either needs to have only 1 page (every report page uses the same
    layout) or otherwise needs the same amount of pages as the report (each
    report page is printed on the respective page in the layout PDF).
    *New in version 0.24.3.*
  * **exclude_start_string** (*str*) -- Sheet names that start with this character/string will
    not be printed. *New in version 0.24.4.*
  * **show** (*bool*) -- Once created, open the PDF file with the default application.
    *New in version 0.24.6.*
  * **quality** (*str*) -- Quality of the PDF file. Can either be `'standard'` or `'minimum'`.
    *New in version 0.26.2.*
* **回傳型別:**
  *str*

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.sheets[0]['A1'].value = 'PDF'
>>> book.to_pdf()
```

另請參見 `xlwings.Sheet.to_pdf`

#### Versionadded
在 0.21.1 版被加入.
