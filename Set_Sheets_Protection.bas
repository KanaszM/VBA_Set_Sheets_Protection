Attribute VB_Name = "Set_Sheets_Protection"
Sub Set_Sheets_Protection()

' User Settings
Const bProtectMode As Boolean = False
Const bProtectWorkbook As Boolean = False
Const sPassword As String = ""
Const bAllowFormattingColumns As Boolean = True
Const bAllowSorting As Boolean = True
Const bAllowFiltering As Boolean = True

' Optimization Start
Application.ScreenUpdating = False
Application.EnableEvents = False

' Iterate through each open workbook
Dim oWB As Workbook ' The current workbook stored from the iteration
For Each oWB In Application.Workbooks
    Dim sWBName As String: sWBName = oWB.Name
    
    If sWBName <> "PERSONAL.XLSB" Then ' Ignore your personal workbook
        Dim iWSCount As Integer: iWSCount = oWB.Worksheets.Count ' The number of sheets in the oWB
        
		' Prompt the user if the operation should continue on the current oWB
        Dim iAnswer As Integer
        iAnswer = MsgBox("Process file: " & sWBName & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Execution")
        If iAnswer = vbYes Then
            
			' Iterate through all sheets from oWB
            For iWS = 1 To iWSCount
                Dim oWS As Worksheet: Set oWS = oWB.Worksheets(iWS) ' The current sheet stored from the iteration
                
                oWS.DisplayPageBreaks = False ' For optimization purposes
                
				' Set the preferred sheet protection mode based on the user settings
                If bProtectMode Then
                    oWS.Protect _
                        Password:=sPassword, _
                        AllowFormattingColumns:=bAllowFormattingColumns, _
                        AllowSorting:=bAllowSorting, _
                        AllowFiltering:=bAllowFiltering
                Else
                    oWS.Unprotect sPassword
                End If
                
            Next iWS
            
			' Set the preffered workbook protection mode based on the user settings
            If bProtectWorkbook Then
                oWB.Protect sPassword, True
            Else
                oWB.Unprotect sPassword
            End If
            
            
        End If
    End If

    Set oWB = Nothing
Next oWB

' Optimization End
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

