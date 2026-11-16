# PivotValueField

### *class* PivotValueField

A field in the *Values* area of a pivot table, accessed via [`PivotValueFields`](pivot_value_fields.md#xlwings.PivotValueFields):

```pycon
>>> # Assumes an existing workbook object named book.
>>> pt = book.sheets[0].pivot_tables[0]
>>> value_field = pt.values.add('Sales')
>>> value_field.name, value_field.function
('Sum of Sales', 'sum')
>>> value_field.function = 'average'
>>> value_field.number_format = '#,##0.00'
>>> value_field.name = 'Average Sales'
>>> pt.values['Average Sales'].source_field
'Sales'
```

#### Versionadded
Added in version 0.37.3.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* function *: Literal['sum', 'count', 'average', 'max', 'min', 'product', 'count_numbers', 'stdev', 'stdevp', 'var', 'varp'] | None*

`"sum"`, `"count"`,
`"average"`, `"max"`, `"min"`, `"product"`, `"count_numbers"`,
`"stdev"`, `"stdevp"`, `"var"` or `"varp"`. Returns `None` if the
engine only knows that Excel picked the default.

* **Type:**
  Returns or sets the summary function

* **Setter type:**
  str

#### *property* name *: str*

Returns or sets the caption, e.g. `"Sum of Sales"`. Excel rejects a
name that equals the name of a source field.

On xlwings Lite and xlwings Server, reading it only works for value
fields that already existed when the script started or if the name
was set in the same script. The automatic caption of a value field
added in the same script can’t be read.

#### *property* number_format *: str*

Returns or sets the number format of the value field.

On xlwings Lite and xlwings Server, reading it only works for value
fields that already existed when the script started or if the format
was set in the same script.

#### *property* parent *: [PivotTable](pivot_table.md#xlwings.PivotTable)*

Returns the pivot table the field belongs to.

#### remove()

Removes the field from the values area.

#### *property* source_field *: str*

The name of the source field this value field summarizes.
