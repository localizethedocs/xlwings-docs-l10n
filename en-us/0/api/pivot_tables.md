# PivotTables

### *class* PivotTables

A collection of all [`PivotTable`](pivot_table.md#xlwings.PivotTable)
objects on the specified sheet:

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.sheets[0].pivot_tables
PivotTables([<PivotTable 'PivotTable1' in Sheet1>])
```

#### Versionadded
Added in version 0.37.3.

#### add(source, destination, name=None, rows=None, columns=None, filters=None, values=None, layout=None)

Creates a pivot table on the sheet of this collection.

On macOS, only the first pivot table on a sheet can be created; a
second one on the same sheet raises `NotImplementedError`. Existing
pivot tables can be modified without that restriction.

* **Parameters:**
  * **source** ([*Range*](range.md#xlwings.Range) *|* [*Table*](table.md#xlwings.main.Table)) – The source data, either a range including the header row
    or a [`Table`](table.md#xlwings.main.Table). Can be on another sheet.
    On xlwings Lite, `expand()` doesn’t see values written in the
    same script until you run `await sheet.load()`; on xlwings
    Server, spell out the range.
  * **destination** ([*Range*](range.md#xlwings.Range)) – The cell where the top-left corner of the pivot
    table goes. Must be on the sheet of this collection.
  * **name** (*str* *|* *None*) – Name of the pivot table. Defaults to Excel’s standard
    name, e.g. `"PivotTable1"`.
  * **rows** (*str* *|* *list* *[**str* *]*  *|* *None*) – Field name(s) for the *Rows* area.
  * **columns** (*str* *|* *list* *[**str* *]*  *|* *None*) – Field name(s) for the *Columns* area.
  * **filters** (*str* *|* *list* *[**str* *]*  *|* *None*) – Field name(s) for the *Filters* area.
  * **values** (*str* *|* *list* *[**str* *|* *tuple* *[**str* *,* *str* *|* *None* *]* *]*  *|* *dict* *[**str* *,* *str* *|* *None* *]*  *|* *None*) – Field name(s) for the *Values* area. Use a mapping
    (`{"Sales": "sum"}`) or `(field, function)` tuples to pick the
    summary function; `None` keeps Excel’s default. The same field
    can only be listed twice via tuples, see
    [`PivotValueFields.add`](pivot_value_fields.md#xlwings.PivotValueFields.add).
  * **layout** (*str* *|* *None*) – `"compact"`, `"outline"` or `"tabular"`.
* **Return type:**
  [*PivotTable*](pivot_table.md#xlwings.PivotTable)

### Examples

```pycon
>>> # Assumes an existing workbook object named book.
>>> data = book.sheets[0]
>>> report = book.sheets.add('Report', after=data)
>>> data['A1'].value = [['Region', 'Year', 'Sales'],
...                     ['North', 2023, 100], ['South', 2024, 200]]
>>> pt = report.pivot_tables.add(
...     source=data['A1:C3'],
...     destination=report['A3'],
...     rows='Region',
...     columns='Year',
...     values={'Sales': 'sum'},
... )
>>> pt.values[0].number_format = '#,##0'
```

The same in steps:

```pycon
>>> pt = report.pivot_tables.add(data['A1:C3'], report['A3'])
>>> pt.rows.add('Region')
>>> pt.columns.add('Year')
>>> pt.values.add('Sales', function='sum', number_format='#,##0')
```

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the sheet the collection belongs to.
