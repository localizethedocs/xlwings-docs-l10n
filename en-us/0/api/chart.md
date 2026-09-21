# Chart

### *class* Chart(name_or_index=None, impl=None)

The chart object is a member of the `charts` collection:

```pycon
>>> import xlwings as xw
>>> sht = xw.books['Book1'].sheets[0]
>>> sht.charts[0]  # or sht.charts['ChartName']
<Chart 'Chart 1' in <Sheet [Book1]Sheet1>>
```

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.9.0.

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
Added in version 0.1.1.

#### delete()

Deletes the chart.

#### *async* get_png()

Fetch the chart as a base64-encoded PNG, on demand.

Returns the image data rather than writing a file, which is what
`to_png()` does.

Requires xlwings Lite.

#### *property* height *: float*

Returns or sets the number of points that represent the height of the
chart.

#### *property* left *: float*

Returns or sets the number of points that represent the horizontal position
of the chart.

#### *property* name *: str*

Returns or sets the name of the chart.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the parent of the chart.

#### Versionadded
Added in version 0.9.0.

#### set_source_data(source)

Sets the source data range for the chart.

* **Parameters:**
  **source** ([*Range*](range.md#xlwings.Range)) – Range object, e.g. `xw.books['Book1'].sheets[0].range('A1')`

#### to_pdf(path=None, show=None, quality='standard')

Exports the chart as PDF.

* **Parameters:**
  * **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path where you want to store the pdf. Defaults to the name of the
    chart in the same directory as the Excel file if the Excel file is
    stored and to the current working directory otherwise.
  * **show** (*bool* *|* *None*) – Once created, open the PDF file with the default application.
  * **quality** (*str*) – Quality of the PDF file. Can either be `'standard'` or `'minimum'`.

#### Versionadded
Added in version 0.26.2.

#### to_png(path=None)

Exports the chart as PNG picture.

* **Parameters:**
  **path** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – Path where you want to store the picture. Defaults to the name of the
  chart in the same directory as the Excel file if the Excel file is
  stored and to the current working directory otherwise.

#### Versionadded
Added in version 0.24.8.

#### *property* top *: float*

Returns or sets the number of points that represent the vertical position
of the chart.

#### *property* width *: float*

Returns or sets the number of points that represent the width of the
chart.
