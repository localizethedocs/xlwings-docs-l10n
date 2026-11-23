# PivotValueFields

### *class* PivotValueFields

The fields in the *Values* area of a pivot table, i.e. `pt.values`.
Iteration follows the field order.

```pycon
>>> # Assumes an existing workbook object named book.
>>> pt = book.sheets[0].pivot_tables[0]
>>> pt.values
PivotValueFields([<PivotValueField 'Sum of Sales' in 'PivotTable1'>])
>>> pt.values[0]  # or pt.values['Sum of Sales']
<PivotValueField 'Sum of Sales' in 'PivotTable1'>
```

#### Versionadded
Added in version 0.37.3.

#### add(field, function=None, name=None, number_format=None)

Adds a source field to the values area, after the existing value
fields. The same source field can be added more than once, e.g. as a
sum and as a count.

* **Parameters:**
  * **field** (*str*) – Name of the source field, see
    [`PivotTable.field_names`](pivot_table.md#xlwings.PivotTable.field_names).
  * **function** (*str* *|* *None*) – Summary function, see
    [`PivotValueField.function`](pivot_value_field.md#xlwings.PivotValueField.function).
    Defaults to Excel’s choice: `"sum"` for numeric fields,
    `"count"` otherwise.
  * **name** (*str* *|* *None*) – Caption, e.g. `"Total Sales"`. Defaults to Excel’s caption,
    e.g. `"Sum of Sales"`. Excel rejects a name that equals the
    name of a source field.
  * **number_format** (*str* *|* *None*) – Number format, e.g. `"#,##0.00"`.
* **Return type:**
  [*PivotValueField*](pivot_value_field.md#xlwings.PivotValueField)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.

#### *property* parent *: [PivotTable](pivot_table.md#xlwings.PivotTable)*

Returns the pivot table the area belongs to.
