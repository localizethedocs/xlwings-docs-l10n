# Names

### *class* Names

A collection of defined `name` objects in the workbook.

Excel's internal `_xlfn.*` and `_xlpm.*` names are excluded. Hidden
user-defined names remain available. The native collection is unchanged.
Iteration snapshots name objects so new internal names cannot shift it.

範例：

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.names
[<Name 'MyName': =Sheet1!$A$3>]
>>> book.sheets[0].names  # sheet scope only
```

#### Versionadded
在 0.9.0 版被加入.

#### add(name, refers_to)

Defines a new name for a range, constant, or formula (including a LAMBDA).

Full support for named constants and formulas requires classic xlwings, xlwings Lite, or xlwings Server. Google Sheets supports named ranges only; Office Scripts only returns named ranges in its snapshot.

* **參數:**
  * **name** (*str*) -- Specifies the text to use as the name. Names cannot include spaces and
    cannot be formatted as cell references.
  * **refers_to** (*str*) -- Describes what the name refers to, in English, using A1-style
    notation.
* **回傳型別:**
  [*Name*](name.md#xlwings.Name)

#### Versionadded
在 0.9.0 版被加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine beingused.

#### Versionadded
在 0.9.0 版被加入.

#### *property* count *: int*

Returns the number of objects in the collection.
