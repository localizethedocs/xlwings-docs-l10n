# TableColumn

### *class* TableColumn

A column in an Excel table. Obtain it again after changing table structure.

### 範例

```python
# Assumes an existing workbook object named book.
table = book.sheets[0].tables["Table1"]
column = table.columns["MyColumn"]
print(column.index)
print(column.data_body_range.value)
column.delete()
```

#### *property* data_body_range *: [Range](range.md#xlwings.Range) | None*

Data cells, excluding header and totals. In xlwings Lite, use `await get_data_body_range()`.

#### delete()

Delete the column from its table. Excel may move cells beside the table.

#### *async* get_data_body_range()

Fetch this column's current data cells in xlwings Lite.

* **回傳型別:**
  [*Range*](range.md#xlwings.Range) | *None*

#### *async* get_range()

Fetch this column's current range in xlwings Lite.

* **回傳型別:**
  [*Range*](range.md#xlwings.Range)

#### *property* index *: int*

One-based position within the table.

#### *property* name *: str*

The column header.

#### *property* range *: [Range](range.md#xlwings.Range)*

All cells in this column. In xlwings Lite, use `await get_range()`.
