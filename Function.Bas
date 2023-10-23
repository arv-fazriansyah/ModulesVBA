Sub CreateNewSheet()
    Sheets.Add
End Sub

Sub DeleteSheet(sheetName As String)
    On Error Resume Next
    Sheets(sheetName).Delete
    On Error GoTo 0
End Sub

Sub RenameSheet(oldName As String, newName As String)
    On Error Resume Next
    Sheets(oldName).Name = newName
    On Error GoTo 0
End Sub

Sub MoveSheet(sheetName As String, afterSheet As String)
    On Error Resume Next
    Sheets(sheetName).Move After:=Sheets(afterSheet)
    On Error GoTo 0
End Sub

Sub CopySheet(sheetName As String, newName As String)
    On Error Resume Next
    Sheets(sheetName).Copy Before:=Sheets(1)
    ActiveSheet.Name = newName
    On Error GoTo 0
End Sub

Sub ProtectSheet(sheetName As String, password As String)
    On Error Resume Next
    Sheets(sheetName).Protect Password:=password
    On Error GoTo 0
End Sub

Sub UnprotectSheet(sheetName As String, password As String)
    On Error Resume Next
    Sheets(sheetName).Unprotect Password:=password
    On Error GoTo 0
End Sub

Sub HideSheet(sheetName As String)
    On Error Resume Next
    Sheets(sheetName).Visible = xlSheetHidden
    On Error GoTo 0
End Sub

Sub UnhideSheet(sheetName As String)
    On Error Resume Next
    Sheets(sheetName).Visible = xlSheetVisible
    On Error GoTo 0
End Sub

Sub SelectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Select
    Next ws
End Sub

Sub SetTabColor(sheetName As String, tabColor As Long)
    On Error Resume Next
    Sheets(sheetName).Tab.Color = tabColor
    On Error GoTo 0
End Sub
