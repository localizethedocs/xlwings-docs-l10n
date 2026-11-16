# RangeColumns

### *class* RangeColumns(rng)

Represents the columns of a range. Do not construct this class directly, use
`Range.columns` instead.

### Example

```python
# Assumes an existing workbook object named book.
sheet1 = book.sheets[0]
myrange = sheet1.range('A1:C4')

assert len(myrange.columns) == 3  # or myrange.columns.count

myrange.columns[0].value = 'a'

assert myrange.columns[2] == sheet1.range('C1:C4')
assert myrange.columns(2) == sheet1.range('B1:B4')

for c in myrange.columns:
    print(c.address)
```

#### autofit()

Autofits the width of the columns.

#### *property* count *: int*

Returns the number of columns.

#### Versionadded
Added in version 0.9.0.

#### *async* get_hidden()

Fetch column visibility in xlwings Lite; return `None` for a mixed selection.

* **Return type:**
  *bool* | *None*

#### *property* hidden *: bool | None*

Whether all represented worksheet columns are hidden.

Setting this property hides or shows the entire worksheet columns represented by the range. Returns `None` when some columns are hidden and others are visible. In xlwings Lite, use `await rng.columns.get_hidden()` to read this state.

* **Setter type:**
  bool
