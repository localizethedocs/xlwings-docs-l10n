# Reports

### *class* Image(filename)

Use this class to provide images to either `render_template()`.

* **Parameters:**
  **filename** (*str* *or* *pathlib.Path object*) – The file name or path

### *class* Markdown(text, style=<MarkdownStyle> h1.font: .bold: True paragraph.blank_lines_after: 1 unordered_list.bullet_character: • unordered_list.blank_lines_after: 1 strong.bold: True emphasis.italic: True)

Markdown objects can be assigned to a single cell or shape via `myrange.value` or
`myshape.text`. They accept a string in Markdown format which will cause the text
in the cell to be formatted accordingly. They can also be used in
`mysheet.render_template()`.

#### NOTE
On macOS, formatting is currently not supported, but things like bullet
points will still work.

* **Parameters:**
  * **text** (*str*) – The text in Markdown syntax
  * **style** (*MarkdownStyle object* *,* *optional*) – The MarkdownStyle object defines how the text will be formatted.

### Examples

```pycon
>>> mysheet['A1'].value = Markdown("A text with *emphasis* and **strong** style.")
>>> myshape.text = Markdown("A text with *emphasis* and **strong** style.")
```

#### Versionadded
Added in version 0.23.0.

### *class* MarkdownStyle

`MarkdownStyle` defines how `Markdown` objects are being rendered in Excel cells
or shapes. Start by instantiating a `MarkdownStyle` object. Printing it will show
you the current (default) style:

```pycon
>>> style = MarkdownStyle()
>>> style
<MarkdownStyle>
h1.font: .bold: True
h1.blank_lines_after: 1
paragraph.blank_lines_after: 1
unordered_list.bullet_character: •
unordered_list.blank_lines_after: 1
strong.bold: True
emphasis.italic: True
```

You can override the defaults, e.g., to make `**strong**` text red instead of
bold, do this:

```pycon
>>> style.strong.bold = False
>>> style.strong.color = (255, 0, 0)
>>> style.strong
strong.color: (255, 0, 0)
```

#### Versionadded
Added in version 0.23.0.

### formatter(func)

Decorator

### render_template(template, output, book_settings=None, app=None, \*\*data)

This function requires xlwings `PRO`.

This is a convenience wrapper around
:meth:`mysheet.render_template <xlwings.Sheet.render_template>`

Writes the values of all key word arguments to the `output` file according to the
`template` and the variables contained in there (Jinja variable syntax).
Following variable types are supported:

strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames, pictures and
Matplotlib/Plotly figures.

* **Parameters:**
  * **template** – Path to your Excel template, e.g.
    `r'C:\Path\to\my_template.xlsx'`
  * **output** – Path to your Report, e.g. `r'C:\Path\to\my_report.xlsx'`
  * **book_settings** – A dict of `xlwings.Book` parameters, for details see:
    :attr:`xlwings.Book`.
    For example: `book_settings={'update_links': False}`.
  * **app** – By passing in an xlwings App instance, you can control where your report
    runs and configure things like `visible=False`. For details see
    :attr:`xlwings.App`. By default, it creates the report in the currently
    active instance of Excel.
  * **data** – All key/value pairs that are used in the template.

* **Returns:**
  xlwings Book

### Examples

In `my_template.xlsx`, put the following Jinja variables in two cells:
`{{ title }}` and `{{ df }}`

```pycon
>>> from xlwings.reports import render_template
>>> import pandas as pd
>>> df = pd.DataFrame(data=[[1,2],[3,4]])
>>> mybook = render_template('my_template.xlsx', 'my_report.xlsx',
                             title='MyTitle', df=df)
```

With many template variables it may be useful to collect the data first:

```pycon
>>> data = dict(title='MyTitle', df=df)
>>> mybook = render_template('my_template.xlsx', 'my_report.xlsx', **data)
```

If you need to handle external links or a password, use it like so:

```pycon
>>> mybook = render_template('my_template.xlsx', 'my_report.xlsx',
                             book_settings={'update_links': True,
                             'password': 'mypassword'}, **data)
```
