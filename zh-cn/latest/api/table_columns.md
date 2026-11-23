# TableColumns

### *class* TableColumns

Table columns. Collection lookup is zero-based; column indexes are one-based. For a newly created table in xlwings Lite, call `await sheet.load()` before accessing its columns, then obtain the table again.

### 示例

```python
# Assumes an existing workbook object named book.
columns = book.sheets[0].tables["Table1"].columns
columns.add("Tax", index=2)
print(columns[1].name)
```

#### add(name, index=None)

Insert a named column before one-based `index`, or append when omitted. Excel may move cells beside the table.

* **返回类型:**
  [*TableColumn*](table_column.md#xlwings.main.TableColumn)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.

#### *async* get_count()

Fetch the current column count in xlwings Lite.

* **返回类型:**
  *int*
