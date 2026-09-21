# Apps

### *class* Apps(impl)

A collection of all `app` objects:

```pycon
>>> import xlwings as xw
>>> xw.apps
Apps([<Excel App 1668>, <Excel App 1644>])
```

#### *property* active *: [App](app.md#xlwings.App) | None*

Returns the active app.

#### Versionadded
Added in version 0.9.0.

#### add(\*\*kwargs)

Creates a new App. The new App becomes the active one. Returns an App object.

#### cleanup()

Removes Excel zombie processes (Windows-only). Note that this is
automatically called with `App.quit()` and `App.kill()` and when the Python
interpreter exits.

#### Versionadded
Added in version 0.30.2.

#### *property* count *: int*

Returns the number of apps.

#### Versionadded
Added in version 0.9.0.

#### keys()

Provides the PIDs of the Excel instances
that act as keys in the Apps collection.

#### Versionadded
Added in version 0.13.0.
