# PageSetup

### *class* PageSetup(impl)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
Added in version 0.24.2.

#### *property* print_area *: str | None*

Gets or sets the range address that defines the print area.

### Examples

```pycon
>>> mysheet.page_setup.print_area = '$A$1:$B$3'
>>> mysheet.page_setup.print_area
'$A$1:$B$3'
>>> mysheet.page_setup.print_area = None  # clear the print_area
```

#### Versionadded
Added in version 0.24.2.
