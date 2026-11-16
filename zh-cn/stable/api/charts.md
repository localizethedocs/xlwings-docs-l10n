# Charts

### *class* Charts

A collection of all `chart` objects on the specified sheet:

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.sheets[0].charts
Charts([<Chart 'Chart 1' in Sheet1>,
        <Chart 'Chart 1' in Sheet1>])
```

#### Versionadded
在 0.9.0 版本加入.

#### add(left=0, top=0, width=355, height=211, chart_type=None, source=None, plot_by=None, name=None, anchor=None, style=227)

Creates a new chart on the specified sheet.

* **参数:**
  * **left** (*float*) -- left position in points. If you use `top`/`left`, you must not
    provide a value for `anchor`.
  * **top** (*float*) -- top position in points. If you use `top`/`left`, you must not
    provide a value for `anchor`.
  * **width** (*float*) -- width in points
  * **height** (*float*) -- height in points
  * **chart_type** (*str* *|* *None*) -- Chart type, see
    [`Chart.chart_type`](chart.md#xlwings.Chart.chart_type).
    *New in version 0.37.3.*
  * **source** ([*Range*](range.md#xlwings.Range) *|* *None*) -- Source data range, see
    [`Chart.set_source_data`](chart.md#xlwings.Chart.set_source_data).
    *New in version 0.37.3.*
  * **plot_by** (*str* *|* *None*) -- `"rows"` or `"columns"`, requires `source`.
    *New in version 0.37.3.*
  * **name** (*str* *|* *None*) -- Excel chart name. Defaults to Excel standard name if not provided,
    e.g., 'Chart 1'. *New in version 0.37.3.*
  * **anchor** ([*Range*](range.md#xlwings.Range) *|* *None*) -- The xlwings Range object of where you want to insert the chart.
    If you use `anchor`, you must not provide values for `top`/`left`.
    *New in version 0.37.3.*
  * **style** (*int* *|* *None*) -- Built-in chart style. Defaults to the modern Excel style 227.
    Use `None` to retain the host application's native creation
    default. See [`Chart.style`](chart.md#xlwings.Chart.style).
* **返回类型:**
  [*Chart*](chart.md#xlwings.Chart)

### 示例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet.range('A1').value = [['Foo1', 'Foo2'], [1, 2]]
>>> chart = sheet.charts.add(
...     source=sheet.range('A1').expand(),
...     chart_type='line',
...     anchor=sheet.range('D1'),
... )
>>> chart.title = 'My chart'
```

The same in steps:

```pycon
>>> chart = sheet.charts.add()
>>> chart.set_source_data(sheet.range('A1').expand())
>>> chart.chart_type = 'line'
>>> chart.name
'Chart 1'
```

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.
