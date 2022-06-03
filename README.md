# VBA_Set_Sheets_Protection
This VBA module will protect / unprotect all sheets from all opened workbooks based on user approval.

![alt text](https://github.com/KanaszM/VBA_Set_Sheets_Protection/blob/main/ReadMe_Resources/Picture1.png)

## User Settings
* Change the following settings before running the sub
```vba
Const bProtectMode As Boolean = False
Const bProtectWorkbook As Boolean = False
Const sPassword As String = ""
Const bAllowFormattingColumns As Boolean = True
Const bAllowSorting As Boolean = True
Const bAllowFiltering As Boolean = True
```

* bProtectMode = Enable or disable sheet protection on the currently opened workbooks
* bProtectWorkbook = Enable or disable workbook protection on the currently opened workbooks
* sPassword = Set the protection password (the same password will be used for both sheet and workbook protection)
* bAllowFormattingColumns, bAllowSorting, bAllowFiltering = Additional sheet protection settings
