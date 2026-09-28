# RangeRows

### *class* RangeRows(rng)

Represents the rows of a range. Do not construct this class directly, use
`Range.rows` instead.

### Example

```python
import xlwings as xw

wb = xw.Book("MyBook.xlsx")
sheet1 = wb.sheets[0]
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
Added in version 0.9.0.
