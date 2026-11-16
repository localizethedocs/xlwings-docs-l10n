# ChartSeries

### *class* ChartSeries

A chart data series.

[`Chart.series`](chart.md#xlwings.Chart.series) and [`Chart.get_series`](chart.md#xlwings.Chart.get_series) return a [`ChartSeriesCollection`](chart_series_collection.md#xlwings.main.ChartSeriesCollection). Indexing or iterating that collection returns `ChartSeries` objects.

Use [`set()`](#xlwings.main.ChartSeries.set) to change several attributes in one operation. On xlwings Lite, use the asynchronous getters to fetch current values from Excel.

```pycon
>>> # Assumes an existing workbook object named book.
>>> chart = book.sheets[0].charts[0]
>>> series = chart.series[0]
>>> series.set(name="Revenue", marker_style="circle", marker_size=8)
```

#### Versionadded
在 0.37.5 版被加入.

#### *property* api *: Any*

Returns the native `pywin32` or `appscript` series object.

#### *property* axis_group *: Literal['primary', 'secondary']*

Returns or sets whether the series is plotted on the `"primary"` or `"secondary"` axes. Format the secondary axes via [`Chart.secondary_value_axis`](chart.md#xlwings.Chart.secondary_value_axis) and [`Chart.secondary_category_axis`](chart.md#xlwings.Chart.secondary_category_axis).

* **Setter type:**
  str

#### Versionadded
在 0.37.6 版被加入.

#### *property* chart_type *: str*

Returns or sets the series chart type, which turns the chart into a combination chart, e.g., columns and a line. See [`Chart.chart_type`](chart.md#xlwings.Chart.chart_type) for the possible values. Excel rejects combinations it doesn't support.

#### Versionadded
在 0.37.6 版被加入.

#### *property* fill_color *: tuple[int, int, int] | None*

Returns or sets the solid series fill color.

* **Setter type:**
  tuple[int, int, int] | str | int

#### *async* get_axis_group()

Fetches whether the series is plotted on the primary or secondary axes.

Requires xlwings Lite.

#### Versionadded
在 0.37.6 版被加入.

* **回傳型別:**
  *str*

#### *async* get_chart_type()

Fetches the series chart type.

Requires xlwings Lite.

#### Versionadded
在 0.37.6 版被加入.

* **回傳型別:**
  *str*

#### *async* get_fill_color()

Fetches the solid series fill color.

Requires xlwings Lite.

* **回傳型別:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_line_color()

Fetches the series line color.

Requires xlwings Lite.

* **回傳型別:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_marker_background_color()

Fetches the marker background color.

Requires xlwings Lite.

* **回傳型別:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_marker_foreground_color()

Fetches the marker foreground color.

Requires xlwings Lite.

* **回傳型別:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_marker_size()

Fetches the marker size.

Requires xlwings Lite.

* **回傳型別:**
  *int*

#### *async* get_marker_style()

Fetches the marker style.

Requires xlwings Lite.

* **回傳型別:**
  *str*

#### *async* get_name()

Fetches the displayed series name.

Requires xlwings Lite.

* **回傳型別:**
  *str*

#### *property* line_color *: tuple[int, int, int] | None*

Returns or sets the series line color.

* **Setter type:**
  tuple[int, int, int] | str | int

#### *property* marker_background_color *: tuple[int, int, int] | None*

Returns or sets the marker background color.

* **Setter type:**
  tuple[int, int, int] | str | int

#### *property* marker_foreground_color *: tuple[int, int, int] | None*

Returns or sets the marker foreground color.

* **Setter type:**
  tuple[int, int, int] | str | int

#### *property* marker_size *: int*

Returns or sets the marker size in points, from 2 through 72.

#### *property* marker_style *: Literal['automatic', 'none', 'square', 'diamond', 'triangle', 'x', 'star', 'dot', 'dash', 'circle', 'plus']*

Returns or sets the marker style.

* **Setter type:**
  str

#### *property* name *: str*

Returns or sets the displayed series name.

#### set(, name=..., marker_style=..., marker_size=..., marker_foreground_color=..., marker_background_color=..., line_color=..., fill_color=..., chart_type=..., axis_group=...)

Sets one or more series attributes in one operation.

Only supplied attributes are changed, and all arguments are validated
before anything is written. `chart_type` and `axis_group` are applied
before the formatting attributes, whatever the argument order, as
changing them can make Excel reformat the series.

```pycon
>>> series.set(chart_type="line_markers", axis_group="secondary")
```
