# Borders

### *class* Borders

The borders of a range, accessed via `myrange.borders`.

Use `set()` to write several attributes in one go or to target a group of
sides, and `clear()` to remove borders.

### Examples

```pycon
>>> myrange = sheet["A1:D10"]
>>> myrange.borders.line_style = "continuous"
>>> myrange.borders.weight = "thin"
>>> myrange.borders.color = "#000000"
>>> myrange.borders["edge_bottom"].weight = "thick"
>>> myrange.borders.set("outside", line_style="double", color=(255, 0, 0))
>>> myrange.borders.clear("inside")
>>> myrange.borders.clear()
>>> [border.line_style for border in myrange.borders]  # 8 sides, in table order
```

#### Versionadded
Added in version 0.37.1.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.37.1.

#### clear(which='everything')

Removes borders. Same as `set(which, line_style=None)`.

* **Parameters:**
  **which** (*str* *|* *list* *[**str* *]*) – A side name, a list of side names or a group, see
  [`set()`](#xlwings.main.Borders.set). Defaults to
  `"everything"`, i.e. all eight sides including the diagonals;
  `clear("all")` leaves the diagonals alone.

### Examples

```pycon
>>> myrange.borders.clear()           # all eight sides
>>> myrange.borders.clear("inside")   # the two inside borders only
```

#### Versionadded
Added in version 0.37.1.

#### *property* color *: tuple[int, int, int] | None*

Returns or sets the color of the four edges and the two inside
borders, i.e. Excel’s “All Borders”, see
[`Border.color`](border.md#xlwings.main.Border.color). Reads the common
value, or `None` if they differ.

* **Setter type:**
  tuple[int, int, int] | str | int

#### Versionadded
Added in version 0.37.1.

#### *async* get_color()

Fetch the common color of the four edges and the two inside borders
on demand, as an RGB tuple.

`None` if they differ. Segments that differ from cell to cell
within a side aren’t detected on any engine, see
[`Border.get_line_style()`](border.md#xlwings.main.Border.get_line_style).

Requires xlwings Lite.

* **Return type:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_line_style()

Fetch the common line style of the four edges and the two inside
borders on demand.

`None` if they differ. Segments that differ from cell to cell
within a side aren’t detected on any engine, see
[`Border.get_line_style()`](border.md#xlwings.main.Border.get_line_style).

Requires xlwings Lite.

* **Return type:**
  *str* | *None*

#### *async* get_weight()

Fetch the common weight of the four edges and the two inside borders
on demand.

`None` if they differ. Segments that differ from cell to cell
within a side aren’t detected on any engine, see
[`Border.get_line_style()`](border.md#xlwings.main.Border.get_line_style).

Requires xlwings Lite.

* **Return type:**
  *str* | *None*

#### *property* line_style *: str | None*

Returns or sets the line style of the four edges and the two inside
borders, i.e. Excel’s “All Borders”, see
[`Border.line_style`](border.md#xlwings.main.Border.line_style). Reads the
common value, or `None` if they differ.

#### Versionadded
Added in version 0.37.1.

#### set(which='all', , line_style=..., weight=..., color=...)

Sets one or more attributes on one or more sides in one go.

Only the attributes you pass are written. For each side, they’re
applied in the order color, weight, line style, so that the line
style is what survives when Excel can’t represent a combination (for
instance, `"dash_dot_dot"` with `"thick"` reads back as `"thin"`).
Validation happens before anything is written; if the host rejects a
combination, the error propagates and earlier writes stay in place.

* **Parameters:**
  * **which** (*str* *|* *list* *[**str* *]*) – A side name (`"edge_top"`, …), a list of side names, or one
    of the groups `"outside"` (the four edges), `"inside"` (the two
    inside borders), `"all"` (outside + inside, the default) and
    `"everything"` (all eight sides, including the diagonals).
  * **line_style** (*str* *|* *None*) – See
    [`Border.line_style`](border.md#xlwings.main.Border.line_style).
    `None` or `"none"` removes the selected borders.
  * **weight** (*str*) – See [`Border.weight`](border.md#xlwings.main.Border.weight).
    `None` isn’t allowed.
  * **color** (*tuple* *[**int* *,* *int* *,* *int* *]*  *|* *str* *|* *int*) – See [`Border.color`](border.md#xlwings.main.Border.color).
    `None` isn’t allowed.

### Examples

```pycon
>>> myrange = sheet["A1:D10"]
>>> myrange.borders.set("outside", line_style="continuous", weight="thin")
>>> myrange.borders.set(["edge_top", "edge_bottom"], line_style="double")
>>> myrange.borders.set("inside", line_style=None)  # same as clear("inside")
```

#### Versionadded
Added in version 0.37.1.

#### *property* weight *: str | None*

Returns or sets the weight of the four edges and the two inside
borders, i.e. Excel’s “All Borders”, see
[`Border.weight`](border.md#xlwings.main.Border.weight). Reads the common
value, or `None` if they differ.

* **Setter type:**
  str

#### Versionadded
Added in version 0.37.1.
