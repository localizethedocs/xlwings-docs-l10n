# Comment

### *class* Comment

A modern threaded comment attached to one cell, distinct from a [Note](note.md).

### 範例

```python
# Assumes an existing workbook object named book.
cell = book.sheets[0]["A1"]
comment = cell.add_comment("Review this value")
comment.text = "Review the updated value"
print(comment.author, comment.text)
comment.delete()
```

#### add_reply(text)

Add a plain-text reply.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* author *: str*

Author name. In xlwings Lite, use `await get_author()`.

#### *property* creation_date *: datetime | None*

Creation date, when Excel has one. In xlwings Lite, use `await get_creation_date()`.

#### delete()

Delete the thread and all its replies.

#### *async* get_author()

Read the author's name.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *async* get_creation_date()

Read the creation date.

Requires xlwings Lite.

* **回傳型別:**
  *datetime* | *None*

#### *async* get_location()

Read the comment's cell.

Requires xlwings Lite.

* **回傳型別:**
  [*Range*](range.md#xlwings.Range) | *None*

#### *async* get_replies()

Fetch replies.

Requires xlwings Lite.

* **回傳型別:**
  *list*[[*CommentReply*](comment_reply.md#xlwings.main.CommentReply)]

#### *async* get_resolved()

Read the resolution state.

Requires xlwings Lite.

* **回傳型別:**
  *bool* | *None*

#### *async* get_text()

Read the comment text, or `None` if it was deleted.

Requires xlwings Lite.

* **回傳型別:**
  *str* | *None*

#### *property* location *: [Range](range.md#xlwings.Range)*

The comment's cell. In xlwings Lite, use `await get_location()`.

#### reopen()

Reopen this thread.

#### *property* replies *: list[[CommentReply](comment_reply.md#xlwings.main.CommentReply)]*

Replies to this comment. In xlwings Lite, use `await get_replies()`.

#### resolve()

Resolve this thread.

#### *property* resolved *: bool*

Thread resolution state. In xlwings Lite, use `await get_resolved()`.

#### *property* text *: str*

Comment text. In xlwings Lite, use `await get_text()` to read it.
