# PivotField

### *class* PivotField

A source field placed in the *Rows*, *Columns* or *Filters* area of a
pivot table, accessed via the [`PivotFields`](pivot_fields.md#xlwings.PivotFields)
collections:

```pycon
>>> # Assumes an existing workbook object named book.
>>> pt = book.sheets[0].pivot_tables[0]
>>> field = pt.rows.add('Region')
>>> field.name
'Region'
>>> field.parent
<PivotTable 'PivotTable1' in Sheet1>
>>> field.remove()
```

#### Versionadded
Added in version 0.37.3.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* name *: str*

The name of the source field. Read-only; this property has no setter.

#### *property* parent *: [PivotTable](pivot_table.md#xlwings.PivotTable)*

Returns the pivot table the field belongs to.

#### remove()

Removes the field from its area.
