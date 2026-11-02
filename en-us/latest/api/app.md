# App

### *class* App(visible=None, spec=None, add_book=True, impl=None)

An app corresponds to an Excel instance and should normally be used as context
manager to make sure that everything is properly cleaned up again and to prevent
zombie processes. New Excel instances can be fired up like so:

```python
import xlwings as xw

with xw.App() as app:
    print(app.books)
```

An app object is a member of the `apps` collection:

```pycon
>>> xw.apps
Apps([<Excel App 1668>, <Excel App 1644>])
>>> xw.apps[1668]  # get the available PIDs via xw.apps.keys()
<Excel App 1668>
>>> xw.apps.active
<Excel App 1668>
```

* **Parameters:**
  * **visible** (*bool* *|* *None*) – Returns or sets a boolean value that determines whether the app is visible. The
    default leaves the state unchanged or sets visible=True if the object doesn’t
    exist yet.
  * **spec** (*str* *|* *None*) – Mac-only, use the full path to the Excel application,
    e.g., `/Applications/Microsoft Office 2011/Microsoft Excel` or
    `/Applications/Microsoft Excel`.
    On Windows, if you want to change the version of Excel that xlwings talks to, go
    to `Control Panel > Programs and Features` and `Repair` the Office version
    that you want as default.

#### NOTE
On Mac, while xlwings allows you to run multiple instances of Excel, it’s a
feature that is not officially supported by Excel for Mac: Unlike on Windows,
Excel will not ask you to open a read-only version of a file if it is already
open in another instance. This means that you need to watch out yourself so
that the same file is not being overwritten from different instances.

#### activate(steal_focus=False)

Activates the Excel app.

* **Parameters:**
  **steal_focus** (*bool*) – If True, make frontmost application
  and hand over focus from Python to Excel.

#### Versionadded
Added in version 0.9.0.

#### alert(prompt, title=None, buttons='ok', mode=None, callback=None)

This corresponds to `MsgBox` in VBA, shows an alert/message box and
returns the value of the pressed button. For xlwings Server, instead of
returning a value, the function accepts the name of a callback to which
it will supply the value of the pressed button.

* **Parameters:**
  * **prompt** (*str*) – The message to be displayed.
  * **title** (*str* *|* *None*) – The title of the alert.
  * **buttons** (*str*) – Can be either `"ok"`, `"ok_cancel"`, `"yes_no"`, or
    `"yes_no_cancel"`.
  * **mode** (*str* *|* *None*) – Can be `"info"` or `"critical"`. Not supported by Google
    Sheets.
  * **callback** (*str* *|* *None*) – Only used by xlwings Server: you can provide the name of
    a function that will be called with the value of the pressed
    button as argument. The function has to exist on the client
    side, i.e., in VBA or JavaScript.

* **Returns:**
  `None` when used with xlwings Server, otherwise the value
  of the pressed button in lowercase: `"ok"`, `"cancel"`, `"yes"`,
  `"no"`.
* **Return type:**
  str | None

#### Versionadded
Added in version 0.27.13.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.9.0.

#### *property* books *: [Books](books.md#xlwings.main.Books)*

A collection of all Book objects that are currently open.

#### Versionadded
Added in version 0.9.0.

#### calculate()

Calculates all open books.

#### Versionadded
Added in version 0.3.6.

#### *property* calculation *: str*

Returns or sets a calculation value that represents the calculation mode.
Modes: `'manual'`, `'automatic'`, `'semiautomatic'`

### Examples

```pycon
>>> import xlwings as xw
>>> wb = xw.Book()
>>> wb.app.calculation = 'manual'
```

#### Versionchanged
Changed in version 0.9.0.

#### *property* cut_copy_mode *: str | None*

Gets or sets the status of the cut or copy mode.
Accepts `False` for setting and returns `None`,
`copy` or `cut` when getting the status.

#### Versionadded
Added in version 0.24.0.

#### *property* display_alerts *: bool*

The default value is True. Set this property to False to suppress prompts and
alert messages while code is running; when a message requires a response, Excel
chooses the default response.

#### Versionadded
Added in version 0.9.0.

#### *property* enable_events *: bool*

`True` if events are enabled. Read/write boolean.

#### Versionadded
Added in version 0.24.4.

#### *async* get_selection()

Returns the selected cells as Range, fetched live from Excel.

Requires xlwings Lite.

#### Versionadded
Added in version 0.35.0.

#### *property* hwnd *: int | None*

Returns the Window handle (Windows-only).

#### Versionadded
Added in version 0.9.0.

#### *property* interactive *: bool*

`True` if Excel is in interactive mode. If you set this property to `False`,
Excel blocks all input from the keyboard and mouse (except input to dialog boxes
that are displayed by your code). Read/write Boolean.
NOTE: Not supported on macOS.

#### Versionadded
Added in version 0.24.4.

#### kill()

Forces the Excel app to quit by killing its process.

#### Versionadded
Added in version 0.9.0.

#### macro(name)

Runs a Sub or Function in Excel VBA that are not part of a specific workbook
but e.g. are part of an add-in.

* **Parameters:**
  **name** (*str*) – e.g., `'Module1.MyMacro'` or `'MyMacro'`

### Examples

```vb.net
Function MySum(x, y)
    MySum = x + y
End Function
```

can be accessed like this:

```pycon
>>> import xlwings as xw
>>> app = xw.App()
>>> my_sum = app.macro('MySum')
>>> my_sum(1, 2)
3
```

Types are supported too:

```vb.net
Function MySum(x as integer, y as integer)
    MySum = x + y
End Function
```

```pycon
>>> import xlwings as xw
>>> app = xw.App()
>>> my_sum = app.macro('MySum')
>>> my_sum(1, 2)
3
```

However typed arrays are not supported. So the following won’t work

```vb.net
Function MySum(arr() as integer)
    ' code here
End Function
```

See also: `Book.macro`

#### Versionadded
Added in version 0.9.0.

#### *property* path *: str*

Returns the path to where the App is installed.

#### Versionadded
Added in version 0.28.4.

#### *property* pid *: int*

Returns the PID of the app.

#### Versionadded
Added in version 0.9.0.

#### properties(\*\*kwargs)

Context manager that allows you to easily change the app’s properties
temporarily. Once the code leaves the with block, the properties are changed
back to their previous state.
Note: Must be used as context manager or else will have no effect. Also, you can
only use app properties that you can both read and write.

### Examples

```python
import xlwings as xw
app = App()

# Sets app.display_alerts = False
with app.properties(display_alerts=False):
# do stuff

# Sets app.calculation = 'manual' and app.enable_events = True
with app.properties(calculation='manual', enable_events=True):
# do stuff

# Makes sure the status bar is reset even if an error happens in the with block
with app.properties(status_bar='Calculating...'):
# do stuff
```

#### Versionadded
Added in version 0.24.4.

#### quit()

Quits the application without saving any workbooks.

#### Versionadded
Added in version 0.3.3.

#### range(cell1, cell2=None)

Range object from the active sheet of the active book, see `Range`.

#### Versionadded
Added in version 0.9.0.

#### render_template(template=None, output=None, book_settings=None, \*\*data)

This function requires xlwings `PRO`.

This is a convenience wrapper around `mysheet.render_template`

Writes the values of all key word arguments to the `output` file according to
the `template` and the variables contained in there (Jinja variable syntax).
Following variable types are supported:

strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames, pictures
and Matplotlib/Plotly figures.

* **Parameters:**
  * **template** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path to your Excel template, e.g. `r'C:\Path\to\my_template.xlsx'`
  * **output** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path to your Report, e.g. `r'C:\Path\to\my_report.xlsx'`
  * **book_settings** (*dict* *[**str* *,* *Any* *]*  *|* *None*) – A dictionary of `xlwings.Book` parameters, for details see:
    `xlwings.Book`.
    For example: `book_settings={'update_links': False}`.
  * **data** (*Any*) – All key/value pairs that are used in the template.

#### Versionadded
Added in version 0.24.4.

#### *property* screen_updating *: bool*

Turn screen updating off to speed up your script. You won’t be able to see
what the script is doing, but it will run faster. Remember to set the
screen_updating property back to True when your script ends.

#### NOTE
On xlwings Server and xlwings Lite, screen updating can only be
suspended until the next sync rather than turned off indefinitely, so
the effect ends on its own before your script does. Setting it back to
`True` is a no-op there, and reading the property isn’t supported.

#### Versionadded
Added in version 0.3.3.

#### *property* selection *: [Range](range.md#xlwings.Range) | None*

Returns the selected cells as Range.

#### Versionadded
Added in version 0.9.0.

#### *property* startup_path *: str*

Returns the path to `XLSTART` which is where the xlwings add-in gets
copied to by doing `xlwings addin install`.

#### Versionadded
Added in version 0.19.4.

#### *property* status_bar *: str | bool*

Gets or sets the value of the status bar.
Returns `False` if Excel has control of it.

#### Versionadded
Added in version 0.20.0.

#### *property* version *: VersionNumber*

Returns the Excel version number object.

### Examples

```pycon
>>> import xlwings as xw
>>> xw.App().version
VersionNumber('15.24')
>>> xw.apps[10559].version.major
15
```

#### Versionchanged
Changed in version 0.9.0.

#### *property* visible *: bool*

Gets or sets the visibility of Excel to `True` or  `False`.

#### Versionadded
Added in version 0.3.3.
