# Note

See also [Notes](notes.md) for worksheet and workbook note collections.
Modern threaded comments use a separate [Comment](comment.md) model.

### *class* Note

Represents a cell Note.

Before the introduction of threaded comments, a Note was called a Comment.

### Examples

```python
# Assumes an existing workbook object named book.
cell = book.sheets[0]["A1"]
note = cell.add_note("Review this value")
print(note.text)
```

#### Versionadded
Added in version 0.24.2.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.24.2.

#### *property* author *: str*

The note’s author. In xlwings Lite, use `await get_author()`.

#### Versionadded
Added in version 0.37.5.

#### delete()

Delete the note.

#### Versionadded
Added in version 0.24.2.

#### *async* get_author()

Read the note’s author from Excel; return `None` if the note is gone.

Requires xlwings Lite.

#### Versionadded
Added in version 0.37.5.

* **Return type:**
  *str* | *None*

#### *async* get_location()

Read the note’s cell from Excel; return `None` if it is gone.

Requires xlwings Lite.

#### Versionadded
Added in version 0.37.5.

* **Return type:**
  [*Range*](range.md#xlwings.Range) | *None*

#### *async* get_text()

Fetch the note’s text on demand.

Requires xlwings Lite.

#### Versionadded
Added in version 0.37.0.

* **Return type:**
  *str* | *None*

#### *property* location *: [Range](range.md#xlwings.Range)*

The cell containing this note. In xlwings Lite, use `await get_location()`.

#### Versionadded
Added in version 0.37.5.

#### *property* text *: str*

Gets or sets the text of a note. Keep in mind that the note must already
exist!

### Examples

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet['A1'].note.text = 'mynote'
>>> sheet['A1'].note.text
'mynote'
```

#### Versionadded
Added in version 0.24.2.
