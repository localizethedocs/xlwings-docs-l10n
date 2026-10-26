# PageSetup

### *class* PageSetup(impl)

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### Versionadded
在 0.24.2 版被加入.

#### *property* print_area *: str | None*

Gets or sets the range address that defines the print area.

### 範例

```pycon
>>> mysheet.page_setup.print_area = '$A$1:$B$3'
>>> mysheet.page_setup.print_area
'$A$1:$B$3'
>>> mysheet.page_setup.print_area = None  # clear the print_area
```

#### Versionadded
在 0.24.2 版被加入.
