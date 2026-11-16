# TableRow

### *class* TableRow

A data row in an Excel table. Row objects refer to a position, so obtain them again after sorting or changing the table structure.

### 示例

```python
# Assumes an existing workbook object named book.
table = book.sheets[0].tables["Table1"]
row = table.rows[0]  # First data row; row.index is 1
print(row.range.value)
row.delete()
```

#### delete()

Delete this table row. Excel may shift cells below the table upward.

#### *async* get_range()

Fetch this row's cells. Use this method for xlwings Lite readback.

* **返回类型:**
  [*Range*](range.md#xlwings.Range)

#### *property* index *: int*

One-based position within the table's data rows.

#### *property* range *: [Range](range.md#xlwings.Range)*

Cells in this row. In xlwings Lite, use `await get_range()`.
