# PageSetup

### *class* PageSetup

Represents a PageSetup object.

#### Versionadded
在 0.24.2 版本加入.

#### *property* api *: Any*

Returns the native object (`pywin32` or `appscript` obj)
of the engine being used.

#### *property* print_area *: str | None*

Gets or sets the range address that defines the print area.

### 示例

```pycon
>>> mysheet.page_setup.print_area = '$A$1:$B$3'
>>> mysheet.page_setup.print_area
'$A$1:$B$3'
>>> mysheet.page_setup.print_area = None  # clear the print_area
```
