# Notes

### *class* Notes(parent)

A collection of notes on one worksheet (`sheet.notes`) or across all worksheets in a workbook (`book.notes`).

Iterate over `sheet.notes` to inspect one worksheet, or `book.notes` to inspect every worksheet. If you already know the cell, access its note directly with `sheet["A1"].note`. Collections also support zero-based indexing and lookup by cell address (`sheet.notes`) or a single-cell [`Range`](range.md#xlwings.Range) (`book.notes`).

### 範例

```python
# Assumes an existing workbook object named book.
sheet = book.sheets[0]
other_sheet = book.sheets.add(after=sheet)
sheet["A1"].add_note("Review")
other_sheet["B2"].add_note("Check")

for note in sheet.notes:
    print(note.location.address, note.text)
for note in book.notes:
    print(note.location.sheet.name, note.location.address, note.text)
```

#### *property* count *: int*

Number of notes in this collection.
