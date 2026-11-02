# BookAsync

### *class* BookAsync(fullname=None, update_links=None, read_only=None, format=None, password=None, write_res_password=None, ignore_read_only_recommended=None, origin=None, delimiter=None, editable=None, notify=None, converter=None, add_to_mru=None, local=None, corrupt_load=None, impl=None, json=None, mode=None, engine=None, \*\*kwargs)

Type hint for the async API in xlwings Lite scripts.

Annotate a script’s `book` parameter with `xw.BookAsync` to opt into the
async, on-demand API — no cell values are pre-loaded, and you read them via
`await myrange.get_value()`:

```python
import xlwings as xw
from xlwings import script

@script
async def myfunc(book: xw.BookAsync):
    data = await book.sheets[0]["A1:B2"].get_value()
```

At runtime it’s a plain `Book`; the annotation only signals
xlwings Lite to skip loading the values of the entire book up front.

#### Versionadded
Added in version 0.36.9.
