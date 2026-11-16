# Shape

### *class* Shape(\*args, \*\*options)

The shape object is a member of the `shapes` collection:

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet.shapes[0]  # or sheet.shapes['ShapeName']
<Shape 'Rectangle 1' in Sheet1>
```

#### Versionchanged
在 0.9.0 版本发生变更.

#### activate()

Activates the shape.

#### Versionadded
在 0.5.0 版本加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj) of the engine
being used.

#### Versionadded
在 0.19.2 版本加入.

#### *property* characters *: [Characters](characters.md#xlwings.main.Characters)*

Returns the [`Characters`](characters.md#xlwings.main.Characters) object for the shape text, used to format part of the text.

#### delete()

Deletes the shape.

#### Versionadded
在 0.5.0 版本加入.

#### *property* font *: [Font](font.md#xlwings.main.Font)*

Returns the [`Font`](font.md#xlwings.main.Font) object for the shape text, used to read or change font formatting.

#### *async* get_text()

Fetch the shape's text on demand.

`None` if the shape holds no text.

Requires xlwings Lite.

* **返回类型:**
  *str* | *None*

#### *property* height *: float*

Returns or sets the number of points that represent the height of the shape.

#### Versionadded
在 0.5.0 版本加入.

#### *property* left *: float*

Returns or sets the number of points that represent the horizontal position
of the shape.

#### Versionadded
在 0.5.0 版本加入.

#### *property* name *: str*

Returns or sets the name of the shape.

#### Versionadded
在 0.5.0 版本加入.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the parent of the shape.

#### Versionadded
在 0.9.0 版本加入.

#### scale_height(factor, relative_to_original_size=False, scale='scale_from_top_left')

* **参数:**
  * **factor** (*float*) -- For example 1.5 to scale it up to 150%
  * **relative_to_original_size** (*bool*) -- If `False`, it scales relative to current
    height (default). For `True` must be a picture or OLE object.
  * **scale** (*str*) -- One of `scale_from_top_left` (default), `scale_from_bottom_right`,
    `scale_from_middle`

#### Versionadded
在 0.19.2 版本加入.

#### scale_width(factor, relative_to_original_size=False, scale='scale_from_top_left')

* **参数:**
  * **factor** (*float*) -- For example 1.5 to scale it up to 150%
  * **relative_to_original_size** (*bool*) -- If `False`, it scales relative to current
    width (default). For `True` must be a picture or OLE object.
  * **scale** (*str*) -- One of `scale_from_top_left` (default), `scale_from_bottom_right`,
    `scale_from_middle`

#### Versionadded
在 0.19.2 版本加入.

#### *property* text *: str*

Returns or sets the text of a shape.

* **Setter type:**
  Any

#### Versionadded
在 0.21.4 版本加入.

#### *property* top *: float*

Returns or sets the number of points that represent the vertical position of
the shape.

#### Versionadded
在 0.5.0 版本加入.

#### *property* type *: str*

Returns the type of the shape.

#### Versionadded
在 0.9.0 版本加入.

#### *property* width *: float*

Returns or sets the number of points that represent the width of the shape.

#### Versionadded
在 0.5.0 版本加入.
