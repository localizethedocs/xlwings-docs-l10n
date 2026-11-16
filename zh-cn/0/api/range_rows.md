# RangeRows

### *class* RangeRows(rng)

Represents the rows of a range. Do not construct this class directly, use
`Range.rows` instead.

### 示例

```python
# Assumes an existing workbook object named book.
sheet1 = book.sheets[0]
myrange = sheet1.range('A1:C4')

assert len(myrange.rows) == 4  # or myrange.rows.count

myrange.rows[0].value = 'a'

assert myrange.rows[2] == sheet1.range('A3:C3')
assert myrange.rows(2) == sheet1.range('A2:C2')

for r in myrange.rows:
    print(r.address)
```

#### autofit()

Autofits the height of the rows.

#### *property* count *: int*

Returns the number of rows.

#### Versionadded
在 0.9.0 版本加入.

#### *async* get_hidden()

Fetch row visibility in xlwings Lite; return `None` for a mixed selection.

* **返回类型:**
  *bool* | *None*

#### *property* hidden *: bool | None*

Whether all represented worksheet rows are hidden.

Setting this property hides or shows the entire worksheet rows represented by the range. Returns `None` when some rows are hidden and others are visible. In xlwings Lite, use `await rng.rows.get_hidden()` to read this state.

* **Setter type:**
  bool
