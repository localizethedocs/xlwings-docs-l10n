# ConditionalFormat

### *class* ConditionalFormat

Represents one conditional-format rule that applies to a range.

#### Versionadded
Added in version 0.37.5.

#### *property* bar_color *: tuple[int, int, int] | None*

The positive fill color of a data-bar rule, otherwise `None`.

#### *property* colors *: tuple[tuple[int, int, int], ...] | None*

The ordered colors of a color-scale rule, otherwise `None`.

Colors run from the minimum criterion to the maximum criterion.

#### delete()

Delete this complete rule from all ranges to which it applies.

#### *property* fill_color *: tuple[int, int, int] | None*

The rule’s fill color as an RGB tuple, or `None` if unset.

#### *property* font_bold *: bool | None*

The rule’s bold setting, or `None` if it doesn’t set bold.

#### *property* font_color *: tuple[int, int, int] | None*

The rule’s font color as an RGB tuple, or `None` if unset.

#### *property* font_italic *: bool | None*

The rule’s italic setting, or `None` if it doesn’t set italic.

#### *property* formula *: str | None*

The formula for a custom-formula rule, otherwise `None`.

#### *property* formula1 *: str | None*

The first operand for a cell-value rule, otherwise `None`.

#### *property* formula2 *: str | None*

The second operand for a between/not-between rule, otherwise `None`.

#### *property* gradient *: bool | None*

Whether a data bar uses a gradient fill, otherwise `None`.

#### *property* icon_set *: Literal['3_arrows', '3_arrows_gray', '3_flags', '3_traffic_lights_1', '3_traffic_lights_2', '3_signs', '3_symbols', '3_symbols_2', '4_arrows', '4_arrows_gray', '4_red_to_black', '4_rating', '4_traffic_lights', '5_arrows', '5_arrows_gray', '5_rating', '5_quarters', '3_stars', '3_triangles', '5_boxes'] | None*

The built-in style of an icon-set rule, otherwise `None`.

#### *property* operator *: Literal['between', 'not_between', 'equal_to', 'not_equal_to', 'greater_than', 'less_than', 'greater_than_or_equal', 'less_than_or_equal'] | None*

The comparison operator for a cell-value rule, otherwise `None`.

#### *property* reverse_order *: bool | None*

Whether an icon set’s icon order is reversed, otherwise `None`.

#### set(, operator=..., formula1=..., formula2=..., formula=..., fill_color=..., font_color=..., font_bold=..., font_italic=..., stop_if_true=...)

Change selected attributes of this rule in place.

Omitted attributes remain unchanged. Cell-value rules accept
`operator`, `formula1` and `formula2`; custom-formula rules accept
`formula`. Formatting and `stop_if_true` apply to either family.
Other rule types can’t be edited by this initial API.

### Examples

```python
formats = await sheet["B2:B12"].get_conditional_formats()
formats[0].set(formula1=70, stop_if_true=True)
```

#### *property* show_value *: bool | None*

Whether cell values remain visible for a data bar or icon set.

#### *property* stop_if_true *: bool | None*

Whether lower-priority rules stop when this rule matches.

`None` is returned for color scales, data bars and icon sets, which
don’t have stop-if-true behavior.

#### *property* threshold_types *: tuple[Literal['automatic', 'lowest_value', 'highest_value', 'number', 'percent', 'percentile', 'formula', 'unknown'], ...] | None*

The ordered criterion types for a visual rule, otherwise `None`.

Color scales include all criteria, data bars include the lower and
upper bounds, and icon sets include only the effective thresholds
between icons.

#### *property* thresholds *: tuple[int | float | str | None, ...] | None*

The values corresponding to
[`threshold_types`](#xlwings.ConditionalFormat.threshold_types).

Criteria such as `"automatic"`, `"lowest_value"` and
`"highest_value"` have a value of `None`.

#### *property* type *: Literal['cell_value', 'custom', 'color_scale', 'data_bar', 'icon_set', 'unknown']*

The normalized rule type.

Rule types outside the initially supported cell-value, custom-formula,
color-scale, data-bar and icon-set families are reported as
`"unknown"` rather than omitted.
