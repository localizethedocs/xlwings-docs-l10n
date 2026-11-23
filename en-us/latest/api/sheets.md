# Sheets

### *class* Sheets

A collection of all `sheet` objects:

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.sheets
Sheets([<Sheet [Book1]Sheet1>, <Sheet [Book1]Sheet2>])
```

#### Versionadded
Added in version 0.9.0.

#### *property* active *: [Sheet](sheet.md#xlwings.Sheet)*

Returns the active Sheet.

#### add(name=None, before=None, after=None)

Creates a new Sheet and makes it the active sheet. Returns the added sheet.

* **Parameters:**
  * **name** (*str* *|* *None*) – Name of the new sheet. If None, will default to Excel’s default name.
  * **before** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) – An object that specifies the sheet before which the new sheet is
    added.
  * **after** ([*Sheet*](sheet.md#xlwings.Sheet) *|* *None*) – An object that specifies the sheet after which the new sheet is
    added.
* **Return type:**
  [*Sheet*](sheet.md#xlwings.Sheet)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.

#### *async* get_active()

Returns the active Sheet, fetched live from Excel.

Requires xlwings Lite.

#### Versionadded
Added in version 0.35.0.

* **Return type:**
  [*Sheet*](sheet.md#xlwings.Sheet)
