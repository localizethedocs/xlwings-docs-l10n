# Shapes

### *class* Shapes(impl)

A collection of all `shape` objects on the specified sheet:

```pycon
>>> import xlwings as xw
>>> xw.books['Book1'].sheets[0].shapes
Shapes([<Shape 'Oval 1' in <Sheet [Book1]Sheet1>>,
        <Shape 'Rectangle 1' in <Sheet [Book1]Sheet1>>])
```

#### Versionadded
Added in version 0.9.0.
