# Font

### *class* Font(impl)

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

#### Versionadded
Added in version 0.23.0.

#### *property* color *: tuple[int, int, int] | None*

Returns or sets the color property (tuple).

```pycon
>>> sheet['A1'].font.color = (255, 0, 0)  # or '#ff0000'
>>> sheet['A1'].font.color
(255, 0, 0)
```

#### Versionadded
Added in version 0.23.0.

#### *async* get_bold()

Fetch the bold property on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

#### *async* get_color()

Fetch the font colour on demand, as an RGB tuple.

`None` if unset or if the range’s cells don’t all agree.

Requires xlwings Lite.

#### *async* get_italic()

Fetch the italic property on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

#### *async* get_name()

Fetch the font name on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

#### *async* get_size()

Fetch the font size on demand.

`None` if the range’s cells don’t all agree.

Requires xlwings Lite.

#### *property* italic *: bool | None*

Returns or sets the italic property (boolean).

```pycon
>>> sheet['A1'].font.italic = True
>>> sheet['A1'].font.italic
True
```

#### Versionadded
Added in version 0.23.0.

#### *property* name *: str | None*

Returns or sets the name of the font (str).

```pycon
>>> sheet['A1'].font.name = 'Calibri'
>>> sheet['A1'].font.name
Calibri
```

#### Versionadded
Added in version 0.23.0.

#### *property* size *: float | None*

Returns or sets the size (float).

```pycon
>>> sheet['A1'].font.size = 13
>>> sheet['A1'].font.size
13
```

#### Versionadded
Added in version 0.23.0.
