# Characters

### *class* Characters(impl)

The characters object can be accessed as an attribute of the range or shape
object.

* `mysheet['A1'].characters`
* `mysheet.shapes[0].characters`

#### NOTE
On macOS, `characters` are currently not supported due to bugs/lack of
support in AppleScript.

#### Versionadded
Added in version 0.23.0.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj) of the engine
being used.

#### Versionadded
Added in version 0.23.0.

#### *property* font *: [Font](font.md#xlwings.main.Font)*

Returns or sets the text property of a `characters` object.

```pycon
>>> sheet['A1'].characters[1:3].font.bold = True
>>> sheet['A1'].characters[1:3].font.bold
True
```

#### Versionadded
Added in version 0.23.0.

#### *async* get_text()

Fetch the characters’ text on demand.

`None` if the shape holds no text.

Requires xlwings Lite.

#### *property* text *: str*

Returns or sets the text property of a `characters` object.

```pycon
>>> sheet['A1'].value = 'Python'
>>> sheet['A1'].characters[:3].text
Pyt
```

#### Versionadded
Added in version 0.23.0.
