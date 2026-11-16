# CommentReply

Replies belong to a [Comment](comment.md). Use `comment.add_reply(text)` to add a reply, and `await comment.get_replies()` in xlwings Lite to enumerate them.

### *class* CommentReply

A plain-text reply to a threaded comment.

### Examples

```python
# Assumes an existing workbook object named book.
comment = book.sheets[0]["A1"].add_comment("Review")
comment.add_reply("Checked")
print(comment.replies[0].text)
```

#### *async* get_text()

Read the reply text, or `None` if it was deleted.

Requires xlwings Lite.

* **Return type:**
  *str* | *None*

#### *property* text *: str*

Reply text. In xlwings Lite, use `await get_text()`.
