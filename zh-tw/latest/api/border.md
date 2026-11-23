# Border

### *class* Border

A single side of a range's borders, accessed via
[`Borders`](borders.md#xlwings.main.Borders):

`mysheet['A1:D10'].borders['edge_bottom']`

Side names are

* `"edge_top"`
* `"edge_bottom"`
* `"edge_left"`
* `"edge_right"`
* `"inside_vertical"`
* `"inside_horizontal"`
* `"diagonal_down"`
* `"diagonal_up"`

#### Versionadded
在 0.37.1 版被加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.37.1 版被加入.

#### *property* color *: tuple[int, int, int] | None*

Returns or sets the color (tuple).

To set the color, use an RGB tuple `(255, 0, 0)`, a hex string like
`'#ff0000'` or an Excel color constant. Unlike `Range.color`, `None`
isn't accepted: remove a border via `line_style = None`. Reads `None`
for a removed border. A range whose cells don't all agree reports one
of its values, see [`line_style`](#xlwings.main.Border.line_style),
and a multi-cell range reads `None` for its diagonals even when a
color was set. Setting the color of a removed border makes it
visible.

```pycon
>>> sheet['A1'].borders['edge_bottom'].color = (255, 0, 0)  # or '#ff0000'
>>> sheet['A1'].borders['edge_bottom'].color
(255, 0, 0)
```

* **Setter type:**
  tuple[int, int, int] | str | int

#### Versionadded
在 0.37.1 版被加入.

#### *async* get_color()

Fetch the color on demand, as an RGB tuple.

`None` for a removed border. xlwings Lite and xlwings Server report a
mixed side differently than the classic xlwings engines, see
[`get_line_style()`](#xlwings.main.Border.get_line_style).

Requires xlwings Lite.

* **回傳型別:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_line_style()

Fetch the line style on demand.

`"none"` for a missing border. No engine detects a side whose
segments differ from cell to cell, but xlwings Lite and xlwings Server
report a different value than the classic xlwings engines do: an edge reports
its first segment's value rather than the range's, and an inside
border of a range whose cells don't share the same border formatting
reads `"none"` even though every cell's own borders are intact.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *async* get_weight()

Fetch the weight on demand.

xlwings Lite and xlwings Server report a mixed side differently than
the classic xlwings engines, see
[`get_line_style()`](#xlwings.main.Border.get_line_style).

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *property* line_style *: str | None*

Returns or sets the line style (str). Available styles:

* `"continuous"`
* `"dash"`
* `"dash_dot"`
* `"dash_dot_dot"`
* `"dot"`
* `"double"`
* `"slant_dash_dot"`
* `"none"`

Setting it to `None` or `"none"` removes the border. Reads `"none"` for a
missing border. Excel doesn't flag a range whose cells don't all agree, so a
mixed range reports one of its values; read a single cell for an
unambiguous answer.

Excel's border attributes influence one another, so a line style may
change the weight (and vice versa), and setting a weight or color on a
removed border makes it visible again. Getters always report the
resulting native values.

```pycon
>>> sheet['A1'].borders['edge_bottom'].line_style = 'double'
>>> sheet['A1'].borders['edge_bottom'].line_style
'double'
>>> sheet['A1'].borders['edge_bottom'].line_style = None
>>> sheet['A1'].borders['edge_bottom'].line_style
'none'
```

#### Versionadded
在 0.37.1 版被加入.

#### *property* weight *: str | None*

Returns or sets the weight (str). Available weihts:

* `"hairline"`
* `"thin"`
* `"medium"`
* `"thick"`

A range whose cells don't all agree reports one of its values, see
[`line_style`](#xlwings.main.Border.line_style). Setting the
weight of a removed border makes it visible.

```pycon
>>> sheet['A1'].borders['edge_bottom'].weight = 'thick'
>>> sheet['A1'].borders['edge_bottom'].weight
'thick'
```

* **Setter type:**
  str

#### Versionadded
在 0.37.1 版被加入.
