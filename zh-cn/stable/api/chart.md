# Chart

### *class* Chart(name_or_index=None)

The chart object is a member of the `charts` collection:

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet.charts[0]  # or sheet.charts['ChartName']
<Chart 'Chart 1' in Sheet1>
```

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.9.0 版本加入.

#### *property* category_axis *: [ChartAxis](chart_axis.md#xlwings.main.ChartAxis)*

Returns the chart's primary category axis.

#### Versionadded
在 0.37.5 版本加入.

#### *property* chart_type *: str*

Returns and sets the chart type of the chart.
The following chart types are available:

`3d_area`,
`3d_area_stacked`,
`3d_area_stacked_100`,
`3d_bar_clustered`,
`3d_bar_stacked`,
`3d_bar_stacked_100`,
`3d_column`,
`3d_column_clustered`,
`3d_column_stacked`,
`3d_column_stacked_100`,
`3d_line`,
`3d_pie`,
`3d_pie_exploded`,
`area`,
`area_stacked`,
`area_stacked_100`,
`bar_clustered`,
`bar_of_pie`,
`bar_stacked`,
`bar_stacked_100`,
`bubble`,
`bubble_3d_effect`,
`column_clustered`,
`column_stacked`,
`column_stacked_100`,
`combination`,
`cone_bar_clustered`,
`cone_bar_stacked`,
`cone_bar_stacked_100`,
`cone_col`,
`cone_col_clustered`,
`cone_col_stacked`,
`cone_col_stacked_100`,
`cylinder_bar_clustered`,
`cylinder_bar_stacked`,
`cylinder_bar_stacked_100`,
`cylinder_col`,
`cylinder_col_clustered`,
`cylinder_col_stacked`,
`cylinder_col_stacked_100`,
`doughnut`,
`doughnut_exploded`,
`line`,
`line_markers`,
`line_markers_stacked`,
`line_markers_stacked_100`,
`line_stacked`,
`line_stacked_100`,
`pie`,
`pie_exploded`,
`pie_of_pie`,
`pyramid_bar_clustered`,
`pyramid_bar_stacked`,
`pyramid_bar_stacked_100`,
`pyramid_col`,
`pyramid_col_clustered`,
`pyramid_col_stacked`,
`pyramid_col_stacked_100`,
`radar`,
`radar_filled`,
`radar_markers`,
`stock_hlc`,
`stock_ohlc`,
`stock_vhlc`,
`stock_vohlc`,
`surface`,
`surface_top_view`,
`surface_top_view_wireframe`,
`surface_wireframe`,
`xy_scatter`,
`xy_scatter_lines`,
`xy_scatter_lines_no_markers`,
`xy_scatter_smooth`,
`xy_scatter_smooth_no_markers`

#### Versionadded
在 0.1.1 版本加入.

#### delete()

Deletes the chart.

#### *async* get_png()

Fetch the chart as a base64-encoded PNG, on demand.

Returns the image data rather than writing a file, which is what
`to_png()` does.

Requires xlwings Lite.

* **返回类型:**
  *str*

#### *async* get_series()

Fetches the chart's current ordered series collection.

A newly created chart needs its source data first, see [`set_source_data()`](#xlwings.Chart.set_source_data). Like other asynchronous getters, this method flushes pending actions first, so Excel creates the chart's series before they are read.

Requires xlwings Lite.

#### Versionadded
在 0.37.5 版本加入.

* **返回类型:**
  [*ChartSeriesCollection*](chart_series_collection.md#xlwings.main.ChartSeriesCollection)

#### *async* get_title()

Fetches the chart title. Returns `None` if the title is hidden.

Requires xlwings Lite.

#### Versionadded
在 0.37.8 版本加入.

* **返回类型:**
  *str* | *None*

#### *property* height *: float*

Returns or sets the number of points that represent the height of the
chart.

#### *property* left *: float*

Returns or sets the number of points that represent the horizontal position
of the chart.

#### *property* legend *: [ChartLegend](chart_legend.md#xlwings.main.ChartLegend)*

Returns the [`ChartLegend`](chart_legend.md#xlwings.main.ChartLegend) of the
chart.

```pycon
>>> chart.legend.visible = True
>>> chart.legend.position = "bottom"
```

#### Versionadded
在 0.37.3 版本加入.

#### *property* name *: str*

Returns or sets the name of the chart.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the parent of the chart.

#### Versionadded
在 0.9.0 版本加入.

#### *property* plot_by *: str*

Returns or sets whether the data series come from the rows or from the
columns of the source data: either `"rows"` or `"columns"`.

On xlwings Lite and xlwings Server, reading it only works after it has
been set in the same script, either via this property or via
`set_source_data(plot_by=...)`.

#### Versionadded
在 0.37.3 版本加入.

#### *property* secondary_category_axis *: [ChartAxis](chart_axis.md#xlwings.main.ChartAxis)*

Returns the chart's secondary category axis.

A chart only has secondary axes while at least one series is plotted on them, see [`ChartSeries.axis_group`](chart_series.md#xlwings.main.ChartSeries.axis_group). Excel hides the secondary category axis by default; set `visible=True` to show it.

#### Versionadded
在 0.37.6 版本加入.

#### *property* secondary_value_axis *: [ChartAxis](chart_axis.md#xlwings.main.ChartAxis)*

Returns the chart's secondary value axis.

A chart only has secondary axes while at least one series is plotted on them, see [`ChartSeries.axis_group`](chart_series.md#xlwings.main.ChartSeries.axis_group).

```pycon
>>> chart.series[1].set(chart_type="line_markers", axis_group="secondary")
>>> chart.secondary_value_axis.set(title="Margin", number_format="0%")
```

#### Versionadded
在 0.37.6 版本加入.

#### *property* series *: [ChartSeriesCollection](chart_series_collection.md#xlwings.main.ChartSeriesCollection)*

Returns the chart's ordered series collection.

In xlwings Lite, use [`get_series()`](#xlwings.Chart.get_series) before inspecting or indexing
the collection.

#### Versionadded
在 0.37.5 版本加入.

#### set_source_data(source, plot_by=None)

Sets the source data range for the chart.

* **参数:**
  * **source** ([*Range*](range.md#xlwings.Range)) -- Range object, e.g. `book.sheets[0].range('A1')`
  * **plot_by** (*str* *|* *None*) -- Whether the data series come from the `"rows"` or from the
    `"columns"` of the source range. Defaults to letting Excel decide.
    *New in version 0.37.3.*

#### set_x_axis_values(source)

Sets the x-axis values (category labels) for every series in the chart.

This is useful when Excel would otherwise interpret a numeric category
column as another data series. Create the chart from the value columns,
then assign the category column separately:

```python
chart = sheet.charts.add(
    source=sheet["B1:B11"],
    chart_type="line",
    plot_by="columns",
)
chart.set_x_axis_values(sheet["A2:A11"])
```

* **参数:**
  **source** ([*Range*](range.md#xlwings.Range)) -- Range containing one x-axis value or category label per data
  point. Do not include the header cell.

#### *property* style *: int*

Returns or sets the built-in chart style.

Valid values are the legacy styles from 1 to 48 and the modern styles
from 201 to 248. The style numbers shown by Excel's chart gallery can
differ from the values exposed by the object model; use the value
produced by Excel's macro recorder to reproduce a gallery style.

On xlwings Lite and xlwings Server, reading it only works after it has
been set in the same script.

#### Versionadded
在 0.37.3 版本加入.

#### *property* title *: str | None*

Returns or sets the chart title. Setting it to `None` hides the title,
setting it to a string shows it.

```pycon
>>> chart.title = "Sales 2026"
>>> chart.title = None  # hides the title
```

On xlwings Lite and xlwings Server, reading the title only works after it
has been set in the same script. In xlwings Lite, use [`get_title()`](#xlwings.Chart.get_title) to
fetch the title of any chart.

#### Versionadded
在 0.37.3 版本加入.

#### to_pdf(path=None, show=None, quality='standard')

Exports the chart as PDF.

* **参数:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Path where you want to store the pdf. Defaults to the name of the
    chart in the same directory as the Excel file if the Excel file is
    stored and to the current working directory otherwise.
  * **show** (*bool* *|* *None*) -- Once created, open the PDF file with the default application.
  * **quality** (*str*) -- Quality of the PDF file. Can either be `'standard'` or `'minimum'`.
* **返回类型:**
  *str*

#### Versionadded
在 0.26.2 版本加入.

#### to_png(path=None)

Exports the chart as PNG picture.

In xlwings Lite, the export is queued like other workbook actions: the file exists only after the script returns or you call `await book.flush()`. Flush before reading the file within the same script.

* **参数:**
  **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) -- Path where you want to store the picture. Defaults to the name of the
  chart in the same directory as the Excel file if the Excel file is
  stored and to the current working directory otherwise.

#### Versionadded
在 0.24.8 版本加入.

#### *property* top *: float*

Returns or sets the number of points that represent the vertical position
of the chart.

#### *property* value_axis *: [ChartAxis](chart_axis.md#xlwings.main.ChartAxis)*

Returns the chart's primary value axis.

#### Versionadded
在 0.37.5 版本加入.

#### *property* width *: float*

Returns or sets the number of points that represent the width of the
chart.
