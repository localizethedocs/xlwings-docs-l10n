# ConditionalFormats

### *class* ConditionalFormats

An ordered collection of conditional-format rules for a range.

New rules are inserted at the top of Excel's conditional-formatting rule
order. Rules are evaluated from highest to lowest priority. If a matching
rule has `stop_if_true=True`, Excel skips lower-priority rules.

In xlwings Lite, use
`await sheet["A1:D10"].get_conditional_formats()` instead of
`sheet["A1:D10"].conditional_formats`.

#### Versionadded
在 0.37.5 版被加入.

#### add_cell_value(operator, formula1, formula2=None, , fill_color=None, font_color=None, font_bold=None, font_italic=None, stop_if_true=False)

Add a cell-value rule.

`formula2` is required for `"between"` and `"not_between"` and
rejected for the other operators. Colors accept the same RGB tuple,
hex string, or Excel color integer forms as other xlwings color APIs.

### 範例

```python
sheet["B2:B12"].conditional_formats.add_cell_value(
    "less_than", 60, fill_color="#ffff00", font_italic=True
)
```

* **回傳型別:**
  [*ConditionalFormat*](conditional_format.md#xlwings.ConditionalFormat)

#### add_color_scale(colors, , thresholds=None, threshold_type='number')

Add a two- or three-color scale.

`colors` contains two or three colors ordered from the minimum to
the maximum. Without `thresholds`, a two-color scale uses the lowest
and highest values, while a three-color scale adds the 50th percentile
as its midpoint. Custom thresholds must match the number of colors and
be strictly increasing. `threshold_type` can be `"number"`,
`"percent"` or `"percentile"`.

### 範例

```python
sheet["B2:B20"].conditional_formats.add_color_scale(
    ["#f8696b", "#ffeb84", "#63be7b"],
    thresholds=[0, 50, 100],
    threshold_type="number",
)
```

* **回傳型別:**
  [*ConditionalFormat*](conditional_format.md#xlwings.ConditionalFormat)

#### add_custom(formula, , fill_color=None, font_color=None, font_bold=None, font_italic=None, stop_if_true=False)

Add a custom-formula rule.

### 範例

```python
sheet["A2:D20"].conditional_formats.add_custom(
    '=$D2="Late"', fill_color="#ffc7ce"
)
```

* **回傳型別:**
  [*ConditionalFormat*](conditional_format.md#xlwings.ConditionalFormat)

#### add_data_bar(color, , minimum=None, maximum=None, threshold_type='number', gradient=True, show_value=True)

Add a data bar.

Omitted bounds are automatic. Supplied bounds use `threshold_type`,
which can be `"number"`, `"percent"` or `"percentile"`.

* **回傳型別:**
  [*ConditionalFormat*](conditional_format.md#xlwings.ConditionalFormat)

#### add_icon_set(icon_set, , thresholds=None, threshold_type='number', show_value=True, reverse_order=False)

Add a built-in icon set.

Custom thresholds contain one fewer value than the number of icons and
must be strictly increasing. Without them, the icons use equal percent
bands (for example, 33 and 67 for a three-icon set).

Valid styles are `"3_arrows"`, `"3_arrows_gray"`, `"3_flags"`,
`"3_traffic_lights_1"`, `"3_traffic_lights_2"`, `"3_signs"`,
`"3_symbols"`, `"3_symbols_2"`, `"4_arrows"`,
`"4_arrows_gray"`, `"4_red_to_black"`, `"4_rating"`,
`"4_traffic_lights"`, `"5_arrows"`, `"5_arrows_gray"`,
`"5_rating"`, `"5_quarters"`, `"3_stars"`,
`"3_triangles"` and `"5_boxes"`.

### 範例

```python
sheet["C2:C20"].conditional_formats.add_icon_set(
    "3_traffic_lights_1", thresholds=[60, 80]
)
```

* **回傳型別:**
  [*ConditionalFormat*](conditional_format.md#xlwings.ConditionalFormat)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### clear()

Clear all conditional formats active on the represented range.

Rules that also apply outside the range remain active there.

#### *property* count *: int*

Returns the number of objects in the collection.
