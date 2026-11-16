# ChartLegend

### *class* ChartLegend

The legend of a chart, accessed via `mychart.legend`:

```pycon
>>> # Assumes an existing workbook object named book.
>>> chart = book.sheets[0].charts[0]
>>> chart.legend.position = "bottom"
>>> chart.legend.visible = False
```

#### Versionadded
在 0.37.3 版被加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* position *: str | None*

`"top"`, `"bottom"`,
`"left"`, `"right"` or `"corner"`. Returns `None` if the legend is hidden.
Setting a position shows the legend.

On xlwings Lite and xlwings Server, reading it only works after it has
been set in the same script.

* **Type:**
  Returns or sets the position of the legend

* **Setter type:**
  str

#### *property* visible *: bool*

Returns or sets whether the legend is shown.

On xlwings Lite and xlwings Server, reading it only works after it has
been set in the same script.
