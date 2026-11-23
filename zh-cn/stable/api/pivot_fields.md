# PivotFields

### *class* PivotFields

The fields in one area of a pivot table, i.e. `pt.rows`, `pt.columns`
or `pt.filters`. Iteration follows the field order.

```pycon
>>> # Assumes an existing workbook object named book.
>>> pt = book.sheets[0].pivot_tables[0]
>>> pt.rows
PivotFields([<PivotField 'Region' in 'PivotTable1'>])
>>> pt.rows[0]  # or pt.rows['Region']
<PivotField 'Region' in 'PivotTable1'>
```

#### Versionadded
在 0.37.3 版本加入.

#### add(name)

Places a source field in this area, after the existing fields.

A field that is already in this area stays where it is; a field in
another of the rows/columns/filters areas is moved here.

* **参数:**
  **name** (*str*) -- Name of the source field, see
  [`PivotTable.field_names`](pivot_table.md#xlwings.PivotTable.field_names).
* **返回类型:**
  [*PivotField*](pivot_field.md#xlwings.PivotField)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.

#### *property* parent *: [PivotTable](pivot_table.md#xlwings.PivotTable)*

Returns the pivot table the area belongs to.
