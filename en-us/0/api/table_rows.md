# TableRows

The collection belongs to a [Table](table.md). Use [TableRow](table_row.md) to delete an indexed row.

### *class* TableRows

Collection of data rows in a table. Row indexes are one-based; square-bracket lookup is zero-based. In xlwings Lite, `len(rows)` uses loaded metadata; use `await rows.get_count()` for the current count.

### Examples

```python
# Assumes an existing workbook object named book.
rows = book.sheets[0].tables["Table1"].rows
rows.add(["Pencil", 2])  # Append
rows.add(["Pen", 3], index=1)  # Insert before the first data row
print(rows[0].range.value)  # ["Pen", 3]
```

#### add(values=None, index=None)

Add one row before one-based `index`, or append when it is `None`.

`values` must contain exactly one value per table column. Omit it to let Excel create a blank row and fill calculated columns. Excel may shift cells below the table downward.

* **Return type:**
  [*TableRow*](table_row.md#xlwings.main.TableRow)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.

#### *async* get_count()

Fetch the current row count. Use this method for xlwings Lite readback.

* **Return type:**
  *int*
