# Books

### *class* Books(impl)

A collection of all `book` objects:

```pycon
>>> import xlwings as xw
>>> xw.books  # active app
Books([<Book [Book1]>, <Book [Book2]>])
>>> xw.apps[10559].books  # specific app, get the PIDs via xw.apps.keys()
Books([<Book [Book1]>, <Book [Book2]>])
```

#### Versionadded
Added in version 0.9.0.

#### *property* active *: [Book](book.md#xlwings.Book)*

Returns the active Book.

#### add()

Creates a new Book. The new Book becomes the active Book. Returns a Book object.

#### *async* get_active()

Returns the active Book without pre-loading values.

This is the entry point to the async API, which reads on demand
(lazy loading) instead of snapshotting the whole workbook up front.

Requires xlwings Lite.

Use `await myrange.get_value()` to read cell values on demand.

#### Versionadded
Added in version 0.35.0.

#### open(fullname=None, update_links=None, read_only=None, format=None, password=None, write_res_password=None, ignore_read_only_recommended=None, origin=None, delimiter=None, editable=None, notify=None, converter=None, add_to_mru=None, local=None, corrupt_load=None, json=None)

Opens a Book if it is not open yet and returns it. If it is already open,
it doesn’t raise an exception but simply returns the Book object.

* **Parameters:**
  * **fullname** (*str* *|* *PathLike* *[**str* *]*  *|* *None*) – filename or fully qualified filename, e.g., `r'C:\path\to\file.xlsx'`
    or `'file.xlsm'`. Without a full path, it looks for the file in the
    current working directory.
  * **Parameters** (*Other*)
  * **see** – `xlwings.Book()`
