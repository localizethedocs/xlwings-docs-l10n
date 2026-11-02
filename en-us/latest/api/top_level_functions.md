# Top-level functions

### load(index=1, header=1, chunksize=5000)

Loads the selected cell(s) of the active workbook into a pandas DataFrame. If you
select a single cell that has adjacent cells, the range is auto-expanded (via
current region) and turned into a pandas DataFrame. If you don’t have pandas
installed, it returns the values as nested lists.

#### NOTE
Only use this in an interactive context like e.g. a Jupyter notebook! Don’t use
this in a script as it depends on the active book.

* **Parameters:**
  * **index** (*bool* *|* *int*) – Defines the number of columns on the left that will be turned into the
    DataFrame’s index
  * **header** (*bool* *|* *int*) – Defines the number of rows at the top that will be turned into the DataFrame’s
    columns
  * **chunksize** (*int*) – Chunks the loading of big arrays.

### Examples

```pycon
>>> import xlwings as xw
>>> xw.load()
```

See also: `view`

#### Versionchanged
Changed in version 0.23.1.

### view(obj, sheet=None, table=True, chunksize=5000)

Opens a new workbook and displays an object on its first sheet by default. If you
provide a sheet object, it will clear the sheet before displaying the object on the
existing sheet.

#### NOTE
Only use this in an interactive context like e.g., a Jupyter notebook! Don’t use
this in a script as it depends on the active book.

* **Parameters:**
  * **obj** (*Any*) – the object to display, e.g. numbers, strings, lists, numpy arrays, pandas
    DataFrames
  * **sheet** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) – Sheet object. If none provided, the first sheet of a new workbook is used.
  * **table** (*bool*) – If your object is a pandas DataFrame, by default it is formatted as an
    Excel
    Table
  * **chunksize** (*int*) – Chunks the loading of big arrays.

### Examples

```pycon
>>> import xlwings as xw
>>> import pandas as pd
>>> import numpy as np
>>> df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
>>> xw.view(df)
```

See also: `load`

#### Versionchanged
Changed in version 0.22.0.
