# Name

### *class* Name(impl)

The name object is a member of the `names` collection:

```pycon
>>> import xlwings as xw
>>> sht = xw.books['Book1'].sheets[0]
>>> sht.names[0]  # or sht.names['MyName']
<Name 'MyName': =Sheet1!$A$3>
```

#### Versionadded
Added in version 0.9.0.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.9.0.

#### delete()

Deletes the name.

#### Versionadded
Added in version 0.9.0.

#### *property* name *: str*

Returns or sets the name of the name object.

#### Versionadded
Added in version 0.9.0.

#### *property* refers_to *: str*

Returns or sets the formula that the name is defined to refer to,
in A1-style notation, beginning with an equal sign.

#### Versionadded
Added in version 0.9.0.

#### *property* refers_to_range *: [Range](range.md#xlwings.Range)*

Returns the Range object referred to by a Name object.

#### Versionadded
Added in version 0.9.0.
