# Names

### *class* Names(impl)

A collection of all `name` objects in the workbook:

```pycon
>>> import xlwings as xw
>>> book = xw.books['Book1']  # book scope and sheet scope
>>> book.names
[<Name 'MyName': =Sheet1!$A$3>]
>>> book.sheets[0].names  # sheet scope only
```

#### Versionadded
在 0.9.0 版被加入.

#### add(name, refers_to)

Defines a new name for a range of cells.

* **參數:**
  * **name** (*str*) -- Specifies the text to use as the name. Names cannot include spaces and
    cannot be formatted as cell references.
  * **refers_to** (*str*) -- Describes what the name refers to, in English, using A1-style
    notation.

#### Versionadded
在 0.9.0 版被加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine beingused.

#### Versionadded
在 0.9.0 版被加入.

#### *property* count *: int*

Returns the number of objects in the collection.
