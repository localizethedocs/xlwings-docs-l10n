# Names

### *class* Names

A collection of defined `name` objects in the workbook.

Excel’s internal `_xlfn.*` and `_xlpm.*` names are excluded. Hidden
user-defined names remain available. The native collection is unchanged.
Iteration snapshots name objects so new internal names cannot shift it.

Examples:

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.names
[<Name 'MyName': =Sheet1!$A$3>]
>>> book.sheets[0].names  # sheet scope only
```

#### Versionadded
Added in version 0.9.0.

#### add(name, refers_to)

Defines a new name for a range, constant, or formula (including a LAMBDA).

Full support for named constants and formulas requires classic xlwings, xlwings Lite, or xlwings Server. Google Sheets supports named ranges only; Office Scripts only returns named ranges in its snapshot.

* **Parameters:**
  * **name** (*str*) – Specifies the text to use as the name. Names cannot include spaces and
    cannot be formatted as cell references.
  * **refers_to** (*str*) – Describes what the name refers to, in English, using A1-style
    notation.
* **Return type:**
  [*Name*](name.md#xlwings.Name)

#### Versionadded
Added in version 0.9.0.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine beingused.

#### Versionadded
Added in version 0.9.0.

#### *property* count *: int*

Returns the number of objects in the collection.
