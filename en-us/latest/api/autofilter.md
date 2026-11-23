# AutoFilter

### *class* AutoFilter(parent)

An AutoFilter belonging to a range or table.

Do not construct this class directly; access it through [`Range.autofilter`](range.md#xlwings.Range.autofilter) or [`Table.autofilter`](table.md#xlwings.main.Table.autofilter).

Fields are one-based column positions relative to the range or table. Value filters accept strings, finite numbers, and booleans. Comparison filters additionally accept Python dates and timezone-naive datetimes. Date and datetime values are not supported by `apply_values()`; to filter for a single exact date or datetime, use `apply_comparison()` with `"equal_to"`. Comparison operators are `"between"`, `"not_between"`, `"equal_to"`, `"not_equal_to"`, `"greater_than"`, `"less_than"`, `"greater_than_or_equal"`, and `"less_than_or_equal"`.

### Examples

```python
from datetime import date

# Assumes an existing workbook object named book.
sheet = book.sheets[0]
myrange = sheet["A1:C6"]
myrange.value = [
    ["Region", "Order date", "Amount"],
    ["East", date(2025, 1, 1), 10],
    ["West", date(2025, 1, 2), 20],
    ["East", date(2025, 1, 3), 30],
    ["North", date(2025, 1, 1), 40],
    ["West", date(2025, 1, 4), 50],
]

myrange.autofilter.apply_values(1, ["East", "West"])
myrange.autofilter.apply_comparison(3, "between", 10, 20)
# Use a comparison for an exact date instead of apply_values().
myrange.autofilter.apply_comparison(2, "equal_to", date(2025, 1, 1))
myrange.autofilter.apply_top_items(3, 2)
myrange.autofilter.apply_comparison(2, "equal_to", None)
myrange.autofilter.clear(2)
myrange.autofilter.clear()
```

#### Versionadded
Added in version 0.37.5.

#### apply_bottom_items(field, count)

Shows the lowest-valued items in a field.

#### apply_bottom_percent(field, percent)

Shows the lowest-valued percentage of items in a field.

#### apply_comparison(field, operator, value1, value2=None)

Filters a field using a comparison.

`None` represents blanks with `"equal_to"` and nonblanks with `"not_equal_to"`. `"between"` and `"not_between"` require a second value; all other operators reject one.

* **Parameters:**
  * **field** (*int*) – One-based column position relative to the range or table.
  * **operator** (*str*) – The comparison to apply.
  * **value1** (*str* *|* *int* *|* *float* *|* *bool* *|* *date* *|* *datetime* *|* *None*) – A string, finite number, boolean, date, timezone-naive datetime, or `None`.
  * **value2** (*str* *|* *int* *|* *float* *|* *bool* *|* *date* *|* *datetime* *|* *None*) – The upper bound for `"between"` and `"not_between"`; omit it for every other operator.
* **Raises:**
  * **TypeError** – If a value has an unsupported type.
  * **ValueError** – If a field, operator, operand combination, or number is invalid.

#### apply_top_items(field, count)

Shows the highest-valued items in a field.

#### apply_top_percent(field, percent)

Shows the highest-valued percentage of items in a field.

#### apply_values(field, values)

Filters a field to rows matching any of the supplied values.

Dates and datetimes are not supported. Use `apply_comparison(field, "equal_to", value)` to filter for a single exact date or datetime.

* **Parameters:**
  * **field** (*int*) – One-based column position relative to the range or table.
  * **values** (*Sequence* *[**str* *|* *int* *|* *float* *|* *bool* *]*) – One or more exact values to include.
* **Raises:**
  * **TypeError** – If `field` isn’t an integer or `values` isn’t a sequence of supported scalar values.
  * **ValueError** – If `field` is outside the target, `values` is empty, or a number isn’t finite.

#### clear(field=None)

Clears one field’s criteria, or all criteria when `field` is omitted.

* **Parameters:**
  **field** (*int* *|* *None*) – Optional one-based column position relative to the range or table.

#### *property* criteria *: list[[AutoFilterCriteria](#xlwings.AutoFilterCriteria)]*

Returns the criteria for all fields.

The list contains one entry per field, including entries whose type is `"none"`. In xlwings Lite, use [`get_criteria()`](#xlwings.AutoFilter.get_criteria) instead.

#### *async* get_criteria()

Returns the criteria for all fields from Excel.

The list contains one entry per field, including entries whose type is `"none"`. Unsupported native criteria are reported as `"unknown"`.

Requires xlwings Lite.

#### Versionadded
Added in version 0.37.5.

* **Return type:**
  *list*[[*AutoFilterCriteria*](#xlwings.AutoFilterCriteria)]

### *class* AutoFilterCriteria

The applied AutoFilter criterion for one field.

Instances are read-only entries returned by [`AutoFilter.criteria`](#xlwings.AutoFilter.criteria) and [`AutoFilter.get_criteria()`](#xlwings.AutoFilter.get_criteria).

#### Versionadded
Added in version 0.37.5.

#### *property* count *: int | None*

The number of items for a top/bottom items filter, otherwise `None`.

This may be `None` when the classic xlwings engine only exposes the calculated cutoff value rather than the requested item count.

#### *property* field *: int*

The one-based field position relative to the filtered range or table.

#### *property* operator *: Literal['between', 'not_between', 'equal_to', 'not_equal_to', 'greater_than', 'less_than', 'greater_than_or_equal', 'less_than_or_equal'] | None*

The comparison operator, otherwise `None`.

#### *property* percent *: float | None*

The percentage for a top/bottom percent filter, otherwise `None`.

This may be `None` when the classic xlwings engine only exposes the calculated cutoff value rather than the requested percentage.

#### *property* type *: Literal['none', 'values', 'comparison', 'top_items', 'bottom_items', 'top_percent', 'bottom_percent', 'unknown']*

The normalized criterion type.

#### *property* value1 *: str | None*

The first normalized comparison operand, otherwise `None`.

#### *property* value2 *: str | None*

The second normalized comparison operand, otherwise `None`.

#### *property* values *: list[str] | None*

The included values for a values filter, otherwise `None`.
