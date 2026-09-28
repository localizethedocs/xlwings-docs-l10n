# Charts

### *class* Charts(impl)

A collection of all `chart` objects on the specified sheet:

```pycon
>>> import xlwings as xw
>>> xw.books['Book1'].sheets[0].charts
Charts([<Chart 'Chart 1' in <Sheet [Book1]Sheet1>>,
        <Chart 'Chart 1' in <Sheet [Book1]Sheet1>>])
```

#### Versionadded
Added in version 0.9.0.

#### add(left=0, top=0, width=355, height=211)

Creates a new chart on the specified sheet.

* **Parameters:**
  * **left** (*float*) – left position in points
  * **top** (*float*) – top position in points
  * **width** (*float*) – width in points
  * **height** (*float*) – height in points

### Examples

```pycon
>>> import xlwings as xw
>>> sht = xw.Book().sheets[0]
>>> sht.range('A1').value = [['Foo1', 'Foo2'], [1, 2]]
>>> chart = sht.charts.add()
>>> chart.set_source_data(sht.range('A1').expand())
>>> chart.chart_type = 'line'
>>> chart.name
'Chart1'
```
