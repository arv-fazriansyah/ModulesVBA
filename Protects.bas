Sub ProtectSheets()
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("DATAUSER")

    Dim lastRowData As Long
    lastRowData = dataSheet.Cells(dataSheet.Rows.count, "AF").End(xlUp).row

    Dim i As Long
    Dim actionType As String
    Dim sheetNameValue As String
    Dim sheetPassword As String
    Dim targetSheet As Worksheet

    For i = 2 To lastRowData
        actionType = dataSheet.Cells(i, "AF").value
        sheetNameValue = dataSheet.Cells(i, "AG").value
        sheetPassword = dataSheet.Cells(i, "AH").value

        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(sheetNameValue)
        On Error GoTo 0

        If Not targetSheet Is Nothing Then
            If actionType = 1 Then
                ProtectSheet targetSheet, sheetPassword
            ElseIf actionType = 0 Then
                UnprotectSheet targetSheet, sheetPassword
            End If
        Else
            MsgBox "Lembar '" & sheetNameValue & "' tidak ditemukan.", vbExclamation
        End If
    Next i
End Sub

Sub ProtectSheet(sheet As Worksheet, password As String)
    sheet.Unprotect password
    sheet.Protect password
End Sub

Sub UnprotectSheet(sheet As Worksheet, password As String)
    sheet.Unprotect password
End Sub
