# Font

### *class* Font

The font object can be accessed as an attribute of the range or shape object.

* `mysheet['A1'].font`
* `mysheet.shapes[0].font`

#### Versionadded
Added in version 0.23.0.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.23.0.

#### *property* bold *: bool | None*

Returns or sets the bold property (boolean).

```pycon
>>> sheet['A1'].font.bold = True
>>> sheet['A1'].font.bold
True
```

* **Setter type:**
  bool

#### Versionadded
Added in version 0.23.0.

#### *property* color *: tuple[int, int, int] | None*

Returns or sets the color property (tuple).

```pycon
>>> sheet['A1'].font.color = (255, 0, 0)  # or '#ff0000'
>>> sheet['A1'].font.color
(255, 0, 0)
```

* **Setter type:**
  tuple[int, int, int] | str

#### Versionadded
Added in version 0.23.0.

#### *async* get_bold()

Fetch the bold property on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

* **Return type:**
  *bool* | *None*

#### *async* get_color()

Fetch the font color on demand, as an RGB tuple.

`None` if unset or if the range’s cells don’t all agree.

Requires xlwings Lite.

* **Return type:**
  *tuple*[*int*, *int*, *int*] | *None*

#### *async* get_italic()

Fetch the italic property on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

* **Return type:**
  *bool* | *None*

#### *async* get_name()

Fetch the font name on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

* **Return type:**
  *str* | *None*

#### *async* get_size()

Fetch the font size on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

* **Return type:**
  *float* | *None*

#### *property* italic *: bool | None*

Returns or sets the italic property (boolean).

```pycon
>>> sheet['A1'].font.italic = True
>>> sheet['A1'].font.italic
True
```

* **Setter type:**
  bool

#### Versionadded
Added in version 0.23.0.

#### *property* name *: str | None*

Returns or sets the name of the font (str).

```pycon
>>> sheet['A1'].font.name = 'Calibri'
>>> sheet['A1'].font.name
Calibri
```

* **Setter type:**
  str

#### Versionadded
Added in version 0.23.0.

#### set(, bold=..., italic=..., size=..., color=..., name=...)

Sets one or more font attributes in one go.

Only the attributes you pass are written. All arguments are validated before anything is written.

* **Parameters:**
  * **bold** (*bool*) – See [`Font.bold`](#xlwings.main.Font.bold).
  * **italic** (*bool*) – See [`Font.italic`](#xlwings.main.Font.italic).
  * **size** (*float*) – A positive font size, see [`Font.size`](#xlwings.main.Font.size).
  * **color** (*tuple* *[**int* *,* *int* *,* *int* *]*  *|* *str* *|* *int*) – An RGB tuple, a hex string or an Excel color integer, see [`Font.color`](#xlwings.main.Font.color).
  * **name** (*str*) – See [`Font.name`](#xlwings.main.Font.name).

### Examples

```pycon
>>> sheet["A1:D1"].font.set(bold=True, size=14, color="#ff0000")
>>> sheet.shapes[0].font.set(italic=True, name="Arial")
```

#### Versionadded
Added in version 0.37.6.

#### *property* size *: float | None*

Returns or sets the size (float).

```pycon
>>> sheet['A1'].font.size = 13
>>> sheet['A1'].font.size
13
```

* **Setter type:**
  float

#### Versionadded
Added in version 0.23.0.
