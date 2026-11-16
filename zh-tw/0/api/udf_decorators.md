# UDF 裝飾器

Each parameter states which integrations support it. **All integrations** means xlwings Lite, xlwings Server, and classic xlwings on Windows (VBA integration).

### func(f=None, \*\*kwargs)

Registers a Python function as an Excel custom function.

Use `@xw.func` or `@xw.func(...)`. Function annotations describe argument and return converters; [`arg`](#xlwings.arg) and [`ret`](#xlwings.ret) provide explicit converter options. The function's docstring supplies its description in Excel. Classic xlwings registers functions with **Import Python UDFs**; xlwings Lite and xlwings Server register them when the project is loaded.

* **參數:**
  **f** (*callable* *|* *None*) -- **All integrations.** Function to decorate when calling `func(my_function)` directly. Leave it out when using `@func` or `@func(...)`.
* **關鍵字引數:**
  * **name** -- **xlwings Lite and xlwings Server only.** Excel-facing name, preserving the supplied spelling. By default, Excel uses the uppercase Python function name.
  * **namespace** -- **xlwings Lite and xlwings Server only.** Function namespace, overriding the module-level `__xlwings_func_namespace__`. An empty string suppresses that module-level namespace. A namespace configured in the add-in manifest is prepended separately.
  * **volatile** -- **All integrations.** Recalculate whenever Excel calculates. Defaults to `False`.
  * **cache** -- **xlwings Lite and xlwings Server only.** Cache custom-function results in the add-in. Defaults to `False`; cannot be combined with `volatile=True` or a streaming function.
  * **help_url** -- **xlwings Lite and xlwings Server only.** URL shown as additional help in Excel.
  * **required_roles** -- **xlwings Server only.** Roles required for a server-authenticated user.
  * **category** -- **Classic xlwings only.** Excel function category: a built-in category number from 1 to 14 or a custom string. Defaults to `"xlwings"`.
  * **call_in_wizard** -- **Classic xlwings only.** Whether to evaluate the function in Excel's function wizard. Defaults to `True`.
  * **async_mode** -- **Classic xlwings only.** Set to `"threading"` to run an I/O-bound function asynchronously. Defaults to `None`.
* **回傳:**
  The decorated function, or a decorator when called with options.
* **回傳型別:**
  callable

### 範例

```python
import xlwings as xw

@xw.func
def add_one(value: float) -> float:
    return value + 1
```

### script(f=None, name=None, required_roles=None, include=None, exclude=None, button=None, show_taskpane=None, \*\*kwargs)

Registers a Python function as a runnable script.

Use `@xw.script` or `@xw.script(...)`. In xlwings Lite and xlwings Server, exactly one parameter must be annotated with [`Book`](book.md#xlwings.Book) or [`BookAsync`](book_async.md#xlwings.BookAsync) so the runtime can supply the workbook. Use that parameter instead of constructing `xw.Book()`. In xlwings Lite, `BookAsync` skips preloading cell values; define an `async def` script and read values with `await rng.get_value()`.

With classic xlwings, **Import Python UDFs** registers the script as a VBA `Sub` procedure. The injected workbook parameter described above applies to xlwings Lite and xlwings Server.

* **參數:**
  * **f** (*callable* *|* *None*) -- **All integrations.** Function to decorate when calling `script(my_function)` directly. Leave it out when using `@script` or `@script(...)`.
  * **name** (*str* *|* *None*) -- **xlwings Lite and xlwings Server only.** Display name of the script. Defaults to the Python function name.
  * **required_roles** (*list* *[**str* *]*  *|* *None*) -- **xlwings Server only.** Roles required for a server-authenticated user.
  * **include** (*str* *|* *list* *[**str* *]*  *|* *None*) -- **xlwings Lite and xlwings Server only.** Sheet name or list of names whose content is loaded for the script.
  * **exclude** (*str* *|* *list* *[**str* *]*  *|* *None*) -- **xlwings Lite and xlwings Server only.** Sheet name or list of names whose content is excluded from loading.
  * **button** (*str* *|* *None*) -- **xlwings Lite and xlwings Server only.** Sheet-button reference, for example `"[mybutton]Sheet1!A1"`.
  * **show_taskpane** (*bool* *|* *None*) -- **xlwings Lite and xlwings Server only.** Whether clicking the sheet button opens the task pane.
* **關鍵字引數:**
  **kwargs** -- **All integrations.** Further options depend on the integration. Classic xlwings supports `category`, `call_in_wizard`, `volatile`, and `async_mode`, as described under [`func`](#xlwings.func); xlwings Lite and xlwings Server use them as additional script metadata. In xlwings Lite, the legacy `lazy` option is deprecated; annotate the workbook parameter with `BookAsync` instead.
* **回傳:**
  The decorated function, or a decorator when called with options.
* **回傳型別:**
  callable

### 範例

In xlwings Lite:

```python
import xlwings as xw

@xw.script(name="Read sales")
async def read_sales(book: xw.BookAsync):
    rows = await book.sheets["Sales"]["A1:B10"].options(ndim=2).get_value()
    print(rows)
```

### arg(arg, convert=None, \*\*kwargs)

Sets the converter and options for a named custom-function argument.

Place `@xw.arg(...)` below `@xw.func`. The argument name must match a Python parameter. Converter options follow [`Range.options`](range.md#xlwings.Range.options); `doc` supplies the argument description shown in Excel.

* **參數:**
  * **arg** (*str*) -- **All integrations.** Name of the Python parameter to configure.
  * **convert** (*Any*) -- **All integrations.** Converter, such as `numpy.array` or `pandas.DataFrame`. Omit it to retain the argument's existing converter, including one selected by a type annotation.
* **關鍵字引數:**
  * **kwargs** -- **All integrations.** Converter options, such as `ndim=2`, `index=False`, or `header=False`.
  * **doc** -- **All integrations.** Argument description shown in Excel.
  * **vba** -- **Classic xlwings only.** VBA expression that supplies a Python argument, for example `vba="Sheet1"`. Excel supplies this argument by evaluating the VBA expression, so users omit it from the worksheet formula.
* **回傳:**
  A decorator that applies the configured conversion options to the function.
* **回傳型別:**
  callable

### 範例

```python
import xlwings as xw

@xw.func
@xw.arg("values", ndim=2)
def total(values):
    return sum(value for row in values for value in row if value is not None)
```

### ret(convert=None, \*\*kwargs)

Sets the converter and options for a custom function's return value.

Place `@xw.ret(...)` below `@xw.func`. Converter options follow [`Range.options`](range.md#xlwings.Range.options) and control how the Python result becomes Excel cell values.

* **參數:**
  **convert** (*Any*) -- **All integrations.** Return converter. Omit it to retain the existing converter, including one selected by a type annotation.
* **關鍵字引數:**
  * **kwargs** -- **All integrations.** Converter options, such as `index=False` and `header=False` for a DataFrame.
  * **expand** -- **Classic xlwings only.** Set to `"table"` for legacy dynamic-array expansion on Excel versions without native dynamic arrays. This overwrites existing data and cannot take volatile formulas such as `TODAY()` as arguments. Excel versions with native dynamic arrays, xlwings Lite, and xlwings Server spill array results without this option.
* **回傳:**
  A decorator that applies the configured conversion options to the function.
* **回傳型別:**
  callable

### 範例

```python
import xlwings as xw

@xw.func
@xw.ret(transpose=True)
def vertical_values():
    return [1, 2, 3]
```
