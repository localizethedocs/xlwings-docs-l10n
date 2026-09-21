# Note

### *class* Note(impl)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.24.2.

#### delete()

Delete the note.

#### Versionadded
Added in version 0.24.2.

#### *async* get_text()

Fetch the note’s text on demand.

Requires xlwings Lite.

#### *property* text *: str*

Gets or sets the text of a note. Keep in mind that the note must already
exist!

### Examples

```pycon
>>> sheet = xw.Book(...).sheets[0]
>>> sheet['A1'].note.text = 'mynote'
>>> sheet['A1'].note.text
>>> 'mynote'
```

#### Versionadded
Added in version 0.24.2.
