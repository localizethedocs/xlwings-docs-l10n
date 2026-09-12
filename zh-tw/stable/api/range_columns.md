# RangeColumns

### *class* RangeColumns(rng)

Represents the columns of a range. Do not construct this class directly, use
`Range.columns` instead.

### Example

```python
import xlwings as xw

wb = xw.Book("MyFile.xlsx")
sheet1 = wb.sheets[0]
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
