# ChartSeriesCollection

### *class* ChartSeriesCollection

An ordered, integer-indexed collection of chart series.

Series names aren't unique in Excel, so string lookup isn't supported.

```pycon
>>> # Assumes an existing workbook object named book.
>>> chart = book.sheets[0].charts[0]
>>> series = chart.series
>>> len(series)
2
>>> series[0].set(name="Revenue", marker_style="circle", marker_size=8)
```

In xlwings Lite, use `series = await chart.get_series()` instead of `chart.series`.

#### Versionadded
在 0.37.5 版本加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.
