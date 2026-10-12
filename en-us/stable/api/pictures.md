# Pictures

### *class* Pictures(impl)

A collection of all `picture` objects on the specified sheet:

```pycon
>>> import xlwings as xw
>>> xw.books['Book1'].sheets[0].pictures
Pictures([<Picture 'Picture 1' in <Sheet [Book1]Sheet1>>,
          <Picture 'Picture 2' in <Sheet [Book1]Sheet1>>])
```

#### Versionadded
Added in version 0.9.0.

#### add(image, link_to_file=False, save_with_document=True, left=None, top=None, width=None, height=None, name=None, update=False, scale=None, format=None, anchor=None, export_options=None)

Adds a picture to the specified sheet.

* **Parameters:**
  * **image** (*str* *|* *PathLike* *[**str* *]*  *|* *Any*) – Either a filepath or a Matplotlib figure object.
  * **left** (*float* *|* *None*) – Left position in points, defaults to 0. If you use `top`/`left`, you
    must not provide a value for `anchor`.
  * **top** (*float* *|* *None*) – Top position in points, defaults to 0. If you use `top`/`left`,
    you must not provide a value for `anchor`.
  * **width** (*float* *|* *None*) – Width in points. Defaults to original width.
  * **height** (*float* *|* *None*) – Height in points. Defaults to original height.
  * **name** (*str* *|* *None*) – Excel picture name. Defaults to Excel standard name if not provided,
    e.g., ‘Picture 1’.
  * **update** (*bool*) – Replace an existing picture with the same name. Requires `name` to
    be set.
  * **scale** (*float* *|* *None*) – Scales your picture by the provided factor.
  * **format** (*str* *|* *None*) – Only used if image is a Matplotlib or Plotly plot. By default, the
    plot is inserted in the “png” format, but you may want to change this to
    a vector-based format like “svg” on Windows (may require Microsoft 365)
    or “eps” on macOS for better print quality. If you use `'vector'`, it
    will be using `'svg'` on Windows and `'eps'` on macOS. To find out which
    formats your version of Excel supports, see:
    [https://support.microsoft.com/en-us/topic/support-for-eps-images-has-been-turned-off-in-office-a069d664-4bcf-415e-a1b5-cbb0c334a840](https://support.microsoft.com/en-us/topic/support-for-eps-images-has-been-turned-off-in-office-a069d664-4bcf-415e-a1b5-cbb0c334a840)
  * **anchor** ([*Range*](range.md#xlwings.Range) *|* *None*) – The xlwings Range object of where you want to insert the picture. If
    you use `anchor`, you must not provide values for `top`/`left`.
    *New in version 0.24.3.*
  * **export_options** (*dict* *[**str* *,* *Any* *]*  *|* *None*) – For Matplotlib plots, this dictionary is passed on to
    `image.savefig()` with the following defaults:
    `{"bbox_inches": "tight", "dpi": 200}`, so if you want to leave the
    picture uncropped and increase dpi to 300, use:
    `export_options={"dpi": 300}`. For Plotly, the options are passed to
    `write_image()`. *New in version 0.27.7.*

### Examples

1. Picture

```pycon
>>> import xlwings as xw
>>> sht = xw.Book().sheets[0]
>>> sht.pictures.add(r'C:\path\to\file.png')
<Picture 'Picture 1' in <Sheet [Book1]Sheet1>>
```

1. Matplotlib

```pycon
>>> import matplotlib.pyplot as plt
>>> fig = plt.figure()
>>> plt.plot([1, 2, 3, 4, 5])
>>> sht.pictures.add(fig, name='MyPlot', update=True)
<Picture 'MyPlot' in <Sheet [Book1]Sheet1>>
```
