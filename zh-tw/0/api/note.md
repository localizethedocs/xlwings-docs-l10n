# Note

See also [Notes](notes.md) for worksheet and workbook note collections.
Modern threaded comments use a separate [Comment](comment.md) model.

### *class* Note

Represents a cell Note.

Before the introduction of threaded comments, a Note was called a Comment.

### 範例

```python
# Assumes an existing workbook object named book.
cell = book.sheets[0]["A1"]
note = cell.add_note("Review this value")
print(note.text)
```

#### Versionadded
在 0.24.2 版被加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.24.2 版被加入.

#### *property* author *: str*

The note's author. In xlwings Lite, use `await get_author()`.

#### Versionadded
在 0.37.5 版被加入.

#### delete()

Delete the note.

#### Versionadded
在 0.24.2 版被加入.

#### *async* get_author()

Read the note's author from Excel; return `None` if the note is gone.

Requires xlwings Lite.

#### Versionadded
在 0.37.5 版被加入.

* **回傳型別:**
  *str* | *None*

#### *async* get_location()

Read the note's cell from Excel; return `None` if it is gone.

Requires xlwings Lite.

#### Versionadded
在 0.37.5 版被加入.

* **回傳型別:**
  [*Range*](range.md#xlwings.Range) | *None*

#### *async* get_text()

Fetch the note's text on demand.

Requires xlwings Lite.

#### Versionadded
在 0.37.0 版被加入.

* **回傳型別:**
  *str* | *None*

#### *property* location *: [Range](range.md#xlwings.Range)*

The cell containing this note. In xlwings Lite, use `await get_location()`.

#### Versionadded
在 0.37.5 版被加入.

#### *property* text *: str*

Gets or sets the text of a note. Keep in mind that the note must already
exist!

### 範例

```pycon
>>> # Assumes an existing workbook object named book.
>>> sheet = book.sheets[0]
>>> sheet['A1'].note.text = 'mynote'
>>> sheet['A1'].note.text
'mynote'
```

#### Versionadded
在 0.24.2 版被加入.
