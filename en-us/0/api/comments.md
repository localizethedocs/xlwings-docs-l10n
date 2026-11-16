# Comments

### *class* Comments(parent, items=None)

Threaded comments on a worksheet or in a workbook.

In xlwings Lite, obtain the collection with `await sheet.get_comments()` or `await book.get_comments()`.

### Examples

```python
# Assumes an existing workbook object named book.
sheet = book.sheets[0]
sheet["A1"].add_comment("Review")
for comment in book.comments:
    print(comment.location.address, comment.text)
print(sheet.comments["A1"].text)
```

#### *property* count *: int*

Number of comments in this collection.
