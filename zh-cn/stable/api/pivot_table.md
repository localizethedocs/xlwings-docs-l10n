# PivotTable

### *class* PivotTable

The pivot table object is a member of the
[`pivot_tables`](sheet.md#xlwings.Sheet.pivot_tables) collection:

```pycon
>>> # Assumes an existing workbook object named book.
>>> mysheet = book.sheets[0]
>>> mysheet.pivot_tables[0]  # or mysheet.pivot_tables['PivotTable1']
<PivotTable 'PivotTable1' in Sheet1>
```

Fields are placed via the four areas of Excel's field list, see
[`rows`](#xlwings.PivotTable.rows),
[`columns`](#xlwings.PivotTable.columns),
[`filters`](#xlwings.PivotTable.filters) and
[`values`](#xlwings.PivotTable.values):

```pycon
>>> pt = mysheet.pivot_tables['PivotTable1']
>>> pt.rows.add('Region')
>>> pt.values.add('Sales', function='sum', number_format='#,##0')
>>> pt.layout = 'tabular'
>>> pt.refresh()
```

#### Versionadded
在 0.37.3 版本加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* columns *: [PivotFields](pivot_fields.md#xlwings.PivotFields)*

The fields in the *Columns* area, see
[`PivotFields`](pivot_fields.md#xlwings.PivotFields).

#### *property* data_body_range *: [Range](range.md#xlwings.Range) | None*

The range of the values area, or `None` if the pivot table has no value fields.

In xlwings Lite, use `await get_data_body_range()`.

#### delete()

Deletes the pivot table.

#### *property* field_names *: list[str]*

The names of the source fields (the column headers of the source
data), i.e. what can be passed to `rows.add()`, `columns.add()`,
`filters.add()` and `values.add()`.

#### *property* filters *: [PivotFields](pivot_fields.md#xlwings.PivotFields)*

The fields in the *Filters* area, see
[`PivotFields`](pivot_fields.md#xlwings.PivotFields).

#### *async* get_data_body_range()

Fetch the current values area, or `None` if the pivot table has no value fields.

Not available on xlwings Server.

* **返回类型:**
  [*Range*](range.md#xlwings.Range) | *None*

#### *async* get_range()

Fetch the current pivot table report range, excluding the filters area.

Not available on xlwings Server.

* **返回类型:**
  [*Range*](range.md#xlwings.Range)

#### *property* layout *: Literal['compact', 'outline', 'tabular'] | None*

`"compact"`, `"outline"` or
`"tabular"`. Returns `None` if the row fields use mixed layouts.

On xlwings Lite and xlwings Server, reading it returns the layout as it
was when the script started, or the value set in the same script.

* **Type:**
  Returns or sets the report layout

* **Setter type:**
  str

#### *property* name *: str*

Returns or sets the name of the pivot table.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the sheet the pivot table is on.

#### *property* range *: [Range](range.md#xlwings.Range)*

The range of the pivot table report, excluding the filters area.

In xlwings Lite, use `await get_range()`.

#### refresh()

Refreshes the pivot table from its source data.

#### *property* rows *: [PivotFields](pivot_fields.md#xlwings.PivotFields)*

The fields in the *Rows* area, see
[`PivotFields`](pivot_fields.md#xlwings.PivotFields).

#### *property* show_column_grand_totals *: bool*

Returns or sets whether the grand totals for columns are shown.

#### *property* show_row_grand_totals *: bool*

Returns or sets whether the grand totals for rows are shown.

#### *property* values *: [PivotValueFields](pivot_value_fields.md#xlwings.PivotValueFields)*

The fields in the *Values* area, see
[`PivotValueFields`](pivot_value_fields.md#xlwings.PivotValueFields).
