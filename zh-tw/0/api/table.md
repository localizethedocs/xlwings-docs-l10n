# Table

See [TableRows](table_rows.md) for row insertion and deletion. See [TableColumns](table_columns.md) for column insertion and deletion. Positions passed to `add()` are one-based, while square-bracket lookup uses Python indexing.

### *class* Table(\*args, \*\*options)

The table object is a member of the `tables` collection:

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet.tables[0]  # or sheet.tables['TableName']
<Table 'Table 1' in Sheet1>
```

#### Versionadded
在 0.21.0 版被加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* autofilter *: [AutoFilter](autofilter.md#xlwings.AutoFilter)*

Returns the AutoFilter for this table.

`field` arguments are one-based column positions relative to the table.

#### Versionadded
在 0.37.5 版被加入.

#### *property* columns *: [TableColumns](table_columns.md#xlwings.main.TableColumns)*

The columns belonging to this table. Integer collection lookup is zero-based.

#### *property* data_body_range *: [Range](range.md#xlwings.Range) | None*

Returns an xlwings range object that represents the range of values,
excluding the header row

#### *property* display_name *: str*

Returns or sets the display name for the specified Table object

#### *property* header_row_range *: [Range](range.md#xlwings.Range) | None*

Returns an xlwings range object that represents the range of the header
row

#### *property* insert_row_range *: [Range](range.md#xlwings.Range) | None*

Returns an xlwings range object representing the row where data is going to
be inserted. This is only available for empty tables, otherwise it'll return
`None`

#### *property* name *: str*

Returns or sets the name of the Table.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the parent of the table.

#### *property* range *: [Range](range.md#xlwings.Range)*

Returns an xlwings range object of the table.

#### resize(range)

Resize a Table by providing an xlwings range object

#### Versionadded
在 0.24.4 版被加入.

#### *property* rows *: [TableRows](table_rows.md#xlwings.main.TableRows)*

Data rows of this table, excluding headers and totals.

#### *property* show_autofilter *: bool*

Turn the autofilter on or off by setting it to `True` or `False`
(read/write boolean)

#### *property* show_headers *: bool*

Show or hide the header (read/write)

#### *property* show_table_style_column_stripes *: bool*

Returns or sets if the Column Stripes table style is used for
(read/write boolean)

#### *property* show_table_style_first_column *: bool*

Returns or sets if the first column is formatted (read/write boolean)

#### *property* show_table_style_last_column *: bool*

Returns or sets if the last column is displayed (read/write boolean)

#### *property* show_table_style_row_stripes *: bool*

Returns or sets if the Row Stripes table style is used
(read/write boolean)

#### *property* show_totals *: bool*

Gets or sets a boolean to show/hide the Total row.

#### sort(keys, ascending=True)

Sort the table's data rows by one or more columns using Excel's native table sort.

* **參數:**
  * **keys** (*int* *|* *Sequence* *[**int* *]*) -- One-based column positions within the table, in priority order.
  * **ascending** (*bool* *|* *Sequence* *[**bool* *]*) -- One direction for every key, or one boolean per key.

The header and totals rows stay in place. Calculated columns, row formatting, and the table's sort state are kept. In xlwings Lite, the sort is queued until the next sync.

#### Versionadded
在 0.37.6 版被加入.

### 範例

```python
table = sheet.tables["Products"]
table.sort([3, 1], [False, True])
```

#### *property* table_style *: str*

Gets or sets the table style.
See `Tables.add` for possible values.

#### *property* totals_row_range *: [Range](range.md#xlwings.Range) | None*

Returns an xlwings range object representing the Total row

#### update(data, index=True)

Updates the Excel table with the provided data.
Currently restricted to DataFrames.

#### Versionchanged
在 0.24.0 版的變更.

* **參數:**
  * **data** (*Any*) -- Currently restricted to pandas DataFrames.
  * **index** (*bool*) -- Whether or not the index of a pandas DataFrame should be written to the
    Excel table.
* **回傳型別:**
  [*Table*](#xlwings.main.Table)

### 範例

```python
import pandas as pd

# Assumes an existing workbook object named book.
sheet = book.sheets[0]
table_name = 'mytable'

# Sample DataFrame
nrows, ncols = 3, 3
df = pd.DataFrame(data=nrows * [ncols * ['test']],
              columns=['col ' + str(i) for i in range(ncols)])

# Hide the index, then insert a new table if it doesn't exist yet,
# otherwise update the existing one
df = df.set_index('col 0')
if table_name in [table.name for table in sheet.tables]:
sheet.tables[table_name].update(df)
else:
mytable = sheet.tables.add(source=sheet['A1'],
                           name=table_name).update(df)
```
