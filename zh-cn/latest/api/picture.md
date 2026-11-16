# Picture

### *class* Picture

The picture object is a member of the `pictures`
collection:

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet.pictures[0]  # or sheet.charts['PictureName']
<Picture 'Picture 1' in Sheet1>
```

#### Versionchanged
在 0.9.0 版本发生变更.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj) of the engine
being used.

#### Versionadded
在 0.9.0 版本加入.

#### delete()

Deletes the picture.

#### Versionadded
在 0.5.0 版本加入.

#### *property* height *: float*

Returns or sets the number of points that represent the height of the
picture.

#### Versionadded
在 0.5.0 版本加入.

#### *property* left *: float*

Returns or sets the number of points that represent the horizontal position
of the picture.

#### Versionadded
在 0.5.0 版本加入.

#### *property* lock_aspect_ratio *: bool*

`True` will keep the original proportion,
`False` will allow you to change height and width independently of each other
(read/write).

#### Versionadded
在 0.24.0 版本加入.

#### *property* name *: str*

Returns or sets the name of the picture.

#### Versionadded
在 0.5.0 版本加入.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the parent of the picture.

#### Versionadded
在 0.9.0 版本加入.

#### *property* top *: float*

Returns or sets the number of points that represent the vertical position
of the picture.

#### Versionadded
在 0.5.0 版本加入.

#### update(image, format=None, export_options=None)

Replaces an existing picture with a new one, taking over the attributes of
the existing picture.

* **参数:**
  * **image** (*str* *|* *PathLike* *[**str* *]*  *|* *Any*) -- Either a filepath or a Matplotlib figure object.
  * **format** (*str* *|* *None*) -- See under `Pictures.add()`
  * **export_options** (*dict* *[**str* *,* *Any* *]*  *|* *None*) -- See under `Pictures.add()`
* **返回类型:**
  [*Picture*](#xlwings.Picture)

#### Versionadded
在 0.5.0 版本加入.

#### *property* width *: float*

Returns or sets the number of points that represent the width of the picture.

#### Versionadded
在 0.5.0 版本加入.
