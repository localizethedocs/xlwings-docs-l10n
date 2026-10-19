# Sheets

### *class* Sheets(impl)

A collection of all `sheet` objects:

```pycon
>>> import xlwings as xw
>>> xw.sheets  # active book
Sheets([<Sheet [Book1]Sheet1>, <Sheet [Book1]Sheet2>])
>>> xw.Book('Book1').sheets  # specific book
Sheets([<Sheet [Book1]Sheet1>, <Sheet [Book1]Sheet2>])
```

#### Versionadded
Added in version 0.9.0.

#### *property* active *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the active Sheet.

#### add(name=None, before=None, after=None)

Creates a new Sheet and makes it the active sheet.

* **Parameters:**
  * **name** (*str* *|* *None*) – Name of the new sheet. If None, will default to Excel’s default name.
  * **before** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) – An object that specifies the sheet before which the new sheet is
    added.
  * **after** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) – An object that specifies the sheet after which the new sheet is
    added.

* **Returns:**
  Added sheet object
* **Return type:**
  [*Sheet*](sheet.md#xlwings.Sheet)

#### *async* get_active()

Returns the active Sheet, fetched live from Excel.

Requires xlwings Lite.

#### Versionadded
Added in version 0.35.0.
