# DataValidation

### *class* DataValidation(parent)

Data validation for a range.

Do not construct this class directly; access it through
[`Range.data_validation`](range.md#xlwings.Range.data_validation).

Comparison setters accept these operators: `"between"`, `"not_between"`,
`"equal_to"`, `"not_equal_to"`, `"greater_than"`, `"less_than"`,
`"greater_than_or_equal"`, and `"less_than_or_equal"`.

### 範例

Create a dropdown from literal values:

```python
sheet["A1:A10"].data_validation.set_list(["Open", "Closed"])
```

Require whole numbers between 1 and 10:

```python
sheet["B1:B10"].data_validation.set_whole_number("between", 1, 10)
```

#### Versionadded
在 0.37.5 版被加入.

#### *property* alert_style *: Literal['stop', 'warning', 'information'] | None*

The error-alert style.

One of `"stop"`, `"warning"`, or `"information"`.

#### *property* api *: Any*

Returns the native data validation object of the engine being used.

#### delete()

Removes data validation from the range.

#### *property* error_message *: str | None*

The validation-error message.

#### *property* error_title *: str | None*

The validation-error title.

#### *property* formula *: str | None*

The formula for a custom validation rule.

#### *property* formula1 *: str | None*

The first comparison operand as an Excel formula string.

#### *property* formula2 *: str | None*

The second operand for between/not-between rules.

#### *property* ignore_blank *: bool | None*

Whether blank cells are ignored by the validation rule.

#### *property* in_cell_dropdown *: bool | None*

Whether a list rule displays an in-cell dropdown.

#### *property* input_message *: str | None*

The input-prompt message.

#### *property* input_title *: str | None*

The input-prompt title.

#### *property* operator *: Literal['between', 'not_between', 'equal_to', 'not_equal_to', 'greater_than', 'less_than', 'greater_than_or_equal', 'less_than_or_equal'] | None*

The comparison operator, or `None` for list and custom rules.

#### set_custom(formula)

Create or update a custom-formula validation rule.

The formula must be an A1-style formula starting with `=`. Existing
prompts and error alerts are preserved.

#### set_date(operator, formula1, formula2=None)

Create or update a date validation rule.

Operands may be Python dates, naive datetimes, Excel serial numbers, or Excel formula strings. Existing prompts and error alerts are preserved.

#### set_decimal(operator, formula1, formula2=None)

Create or update a decimal validation rule.

`formula2` is required for `"between"` and `"not_between"` and
rejected for every other operator. Existing prompts and error alerts are
preserved.

#### set_list(source, , in_cell_dropdown=True)

Creates or updates a list validation on the range.

Existing input prompts and error alerts are preserved. The source can be a
non-empty sequence of strings, numbers, or booleans, a one-dimensional Range,
or a Name that refers to a one-dimensional range in the same workbook.

* **參數:**
  * **source** (*Sequence* *[**str* *|* *int* *|* *float* *|* *bool* *]*  *|* [*Range*](range.md#xlwings.Range) *|* [*Name*](name.md#xlwings.Name)) -- Allowed list values, a worksheet range, or a named range.
  * **in_cell_dropdown** (*bool*) -- Whether Excel shows the list's in-cell dropdown.

#### set_text_length(operator, formula1, formula2=None)

Create or update a text-length validation rule.

`formula2` is required for `"between"` and `"not_between"` and
rejected for every other operator. Existing prompts and error alerts are
preserved.

#### set_time(operator, formula1, formula2=None)

Create or update a time validation rule.

Operands may be naive Python times, Excel day fractions, or Excel formula strings. Existing prompts and error alerts are preserved.

#### set_whole_number(operator, formula1, formula2=None)

Create or update a whole-number validation rule.

`formula2` is required for `"between"` and `"not_between"` and
rejected for every other operator. Existing prompts and error alerts are
preserved.

#### *property* show_error *: bool | None*

Whether invalid entries display an error alert.

#### *property* show_input *: bool | None*

Whether the input prompt is shown.

#### *property* source *: str | None*

The source string for a list validation rule.

#### *property* type *: Literal['none', 'whole_number', 'decimal', 'list', 'date', 'time', 'text_length', 'custom', 'inconsistent', 'mixed_criteria', 'unknown']*

The normalized validation type for this range.

`"none"` means that no cell has validation, `"mixed_criteria"` means
that only some cells have validation, and `"inconsistent"` means that
cells have different validation rules. Unsupported native rule types are
reported as `"unknown"`.
