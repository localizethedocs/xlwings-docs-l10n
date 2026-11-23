# ChartAxis

### *class* ChartAxis

A chart axis, accessed through [`Chart.category_axis`](chart.md#xlwings.Chart.category_axis) and [`Chart.value_axis`](chart.md#xlwings.Chart.value_axis) for the primary axes, or [`Chart.secondary_category_axis`](chart.md#xlwings.Chart.secondary_category_axis) and [`Chart.secondary_value_axis`](chart.md#xlwings.Chart.secondary_value_axis) for the secondary axes.

Use [`set()`](#xlwings.main.ChartAxis.set) to change several attributes in one operation. On xlwings Lite, use the asynchronous getters to fetch the current values from Excel.

```pycon
>>> # Assumes an existing workbook object named book.
>>> chart = book.sheets[0].charts[0]
>>> chart.value_axis.set(
...     title="Revenue", minimum_scale=0, major_unit=10_000
... )
```

#### Versionadded
Added in version 0.37.5.

#### *property* api *: Any*

Returns the native `pywin32` object on Windows.

#### *property* crosses *: str*

`"automatic"`, `"minimum"`, or `"maximum"`. On a category axis, `"maximum"` means the value axis crosses at the last category. Reads return `"custom"` if the crossing point was set to a specific value in Excel.

* **Type:**
  Returns or sets where the other axis crosses this axis

#### *async* get_crosses()

Fetches where the other axis crosses this axis: `"automatic"`, `"minimum"`, `"maximum"`, or `"custom"`.

Requires xlwings Lite.

* **Return type:**
  *str*

#### *async* get_major_unit()

Fetches the resolved major-unit interval.

Requires xlwings Lite.

* **Return type:**
  *float*

#### *async* get_maximum_scale()

Fetches the resolved maximum scale.

Requires xlwings Lite.

* **Return type:**
  *float*

#### *async* get_minimum_scale()

Fetches the resolved minimum scale.

Requires xlwings Lite.

* **Return type:**
  *float*

#### *async* get_number_format()

Fetches the tick-label number format.

Requires xlwings Lite.

* **Return type:**
  *str*

#### *async* get_reverse_order()

Fetches whether the axis is plotted in reverse order.

Requires xlwings Lite.

* **Return type:**
  *bool*

#### *async* get_title()

Fetches the axis title. Returns `None` if the axis or title is hidden.

Requires xlwings Lite.

* **Return type:**
  *str* | *None*

#### *async* get_visible()

Fetches whether the axis is shown.

Requires xlwings Lite.

* **Return type:**
  *bool*

#### *property* major_unit *: float*

Returns or sets the interval between major tick marks. Set to `None` to restore Excel’s automatic interval; reads return the resolved value.

* **Setter type:**
  float | None

#### *property* maximum_scale *: float*

Returns or sets the maximum scale. Set to `None` to restore Excel’s automatic scale; reads return the resolved numeric value.

* **Setter type:**
  float | None

#### *property* minimum_scale *: float*

Returns or sets the minimum scale. Set to `None` to restore Excel’s automatic scale; reads return the resolved numeric value.

* **Setter type:**
  float | None

#### *property* number_format *: str*

Returns or sets the format code for the axis tick labels.

#### *property* reverse_order *: bool*

Returns or sets whether the axis plots its categories or values in reverse order (“Categories in reverse order” / “Values in reverse order” in Excel).

Bar charts draw the first category at the bottom; reversing the category axis puts it at the top. This only flips the display order, it doesn’t sort: sort the source data to order the bars by value. Excel then also moves the value axis to the other side; set [`crosses`](#xlwings.main.ChartAxis.crosses) to `"maximum"` to keep it where it was:

```pycon
>>> chart.category_axis.set(reverse_order=True, crosses="maximum")
```

#### set(, title=..., minimum_scale=..., maximum_scale=..., major_unit=..., number_format=..., visible=..., reverse_order=..., crosses=...)

Sets one or more axis attributes in one operation.

Only supplied attributes are changed, and all arguments are validated before anything is written. `visible=True` is applied before the other attributes; `visible=False` is applied last so it describes the final state.

* **Parameters:**
  * **title** (*str* *|* *None*) – Axis title, or `None` to hide it.
  * **minimum_scale** (*float* *|* *None*) – Minimum scale, or `None` for automatic scaling.
  * **maximum_scale** (*float* *|* *None*) – Maximum scale, or `None` for automatic scaling.
  * **major_unit** (*float* *|* *None*) – Positive major-unit interval, or `None` for automatic.
  * **number_format** (*str*) – Format code for the tick labels.
  * **visible** (*bool*) – Whether the axis is shown.
  * **reverse_order** (*bool*) – Whether the axis is plotted in reverse order.
  * **crosses** (*str*) – Where the other axis crosses this axis: `"automatic"`, `"minimum"`, or `"maximum"`.

#### *property* title *: str | None*

Returns or sets the axis title. Setting it to `None` hides the title; setting it to a string shows the axis and its title.

#### *property* visible *: bool*

Returns or sets whether the axis is shown.
