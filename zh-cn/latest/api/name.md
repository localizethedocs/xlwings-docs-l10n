# Name

### *class* Name

The name object is a member of the `names` collection:

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet.names[0]  # or sheet.names['MyName']
<Name 'MyName': =Sheet1!$A$3>
```

#### Versionadded
在 0.9.0 版本加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.9.0 版本加入.

#### delete()

Deletes the name.

#### Versionadded
在 0.9.0 版本加入.

#### *property* name *: str*

Returns or sets the name of the name object.

#### Versionadded
在 0.9.0 版本加入.

#### *property* refers_to *: str | None*

Returns or sets the formula that the name is defined to refer to,
in A1-style notation, beginning with an equal sign.

Returns `None` when an older remote client supplies neither the
definition nor single-range coordinates for the name.

* **Setter type:**
  str

#### Versionadded
在 0.9.0 版本加入.

#### *property* refers_to_range *: [Range](range.md#xlwings.Range)*

Returns the Range object referred to by a Name object.

#### Versionadded
在 0.9.0 版本加入.
