# Note

### *class* Note(impl)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.24.2 版被加入.

#### delete()

Delete the note.

#### Versionadded
在 0.24.2 版被加入.

#### *async* get_text()

Fetch the note's text on demand.

Requires xlwings Lite.

#### *property* text *: str*

Gets or sets the text of a note. Keep in mind that the note must already
exist!

### 範例

```pycon
>>> sheet = xw.Book(...).sheets[0]
>>> sheet['A1'].note.text = 'mynote'
>>> sheet['A1'].note.text
>>> 'mynote'
```

#### Versionadded
在 0.24.2 版被加入.
