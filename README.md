# VBA_Set_Sheets_Protection
This VBA module will protect / unprotect all sheets from all opened workbooks based on user approval.

## User Settings
* Change the following settings before running the sub
```vba
Const bProtectMode As Boolean = False
Const bProtectWorkbook As Boolean = False
Const sPassword As String = "]"
Const bAllowFormattingColumns As Boolean = True
Const bAllowSorting As Boolean = True
Const bAllowFiltering As Boolean = True
```
