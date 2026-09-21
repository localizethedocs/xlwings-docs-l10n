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
在 0.9.0 版被加入.

#### *property* active *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the active Sheet.

#### add(name=None, before=None, after=None)

Creates a new Sheet and makes it the active sheet.

* **參數:**
  * **name** (*str* *|* *None*) -- Name of the new sheet. If None, will default to Excel's default name.
  * **before** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) -- An object that specifies the sheet before which the new sheet is
    added.
  * **after** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) -- An object that specifies the sheet after which the new sheet is
    added.

* **回傳:**
  Added sheet object
* **回傳型別:**
  [*Sheet*](sheet.md#xlwings.Sheet)

#### *async* get_active()

Returns the active Sheet, fetched live from Excel.

Requires xlwings Lite.

#### Versionadded
在 0.35.0 版被加入.
