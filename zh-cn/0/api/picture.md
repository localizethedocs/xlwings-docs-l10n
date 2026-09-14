# Picture

### *class* Picture(impl=None)

The picture object is a member of the `pictures`
collection:

```pycon
>>> import xlwings as xw
>>> sht = xw.books['Book1'].sheets[0]
>>> sht.pictures[0]  # or sht.charts['PictureName']
<Picture 'Picture 1' in <Sheet [Book1]Sheet1>>
```

#### Versionchanged
Changed in version 0.9.0.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj) of the engine
being used.

#### Versionadded
Added in version 0.9.0.

#### delete()

Deletes the picture.

#### Versionadded
Added in version 0.5.0.

#### *property* height *: float*

Returns or sets the number of points that represent the height of the
picture.

#### Versionadded
Added in version 0.5.0.

#### *property* left *: float*

Returns or sets the number of points that represent the horizontal position
of the picture.

#### Versionadded
Added in version 0.5.0.

#### *property* lock_aspect_ratio *: bool*

`True` will keep the original proportion,
`False` will allow you to change height and width independently of each other
(read/write).

#### Versionadded
Added in version 0.24.0.

#### *property* name *: str*

Returns or sets the name of the picture.

#### Versionadded
Added in version 0.5.0.

#### *property* parent *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the parent of the picture.

#### Versionadded
Added in version 0.9.0.

#### *property* top *: float*

Returns or sets the number of points that represent the vertical position
of the picture.

#### Versionadded
Added in version 0.5.0.

#### update(image, format=None, export_options=None)

Replaces an existing picture with a new one, taking over the attributes of
the existing picture.

* **Parameters:**
  * **image** (*str* *|* *PathLike* *[**str* *]*  *|* *Any*) – Either a filepath or a Matplotlib figure object.
  * **format** (*str* *|* *None*) – See under `Pictures.add()`
  * **export_options** (*dict* *[**str* *,* *Any* *]*  *|* *None*) – See under `Pictures.add()`

#### Versionadded
Added in version 0.5.0.

#### *property* width *: float*

Returns or sets the number of points that represent the width of the picture.

#### Versionadded
Added in version 0.5.0.
