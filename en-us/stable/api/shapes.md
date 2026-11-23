# Shapes

### *class* Shapes

A collection of all `shape` objects on the specified sheet:

```pycon
>>> # Assumes an existing workbook object named book.
>>> book.sheets[0].shapes
Shapes([<Shape 'Oval 1' in Sheet1>,
        <Shape 'Rectangle 1' in Sheet1>])
```

#### Versionadded
Added in version 0.9.0.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* count *: int*

Returns the number of objects in the collection.
