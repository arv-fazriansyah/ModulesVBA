Sub CopyFormulas()
    Dim sourceSheetName As String
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    sourceSheetName = "Sheet1"  ' Ganti dengan nama sheet sumber yang diinginkan
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Sheet sumber '" & sourceSheetName & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    
    For i = 2 To lastRow ' Anggap baris 1 adalah untuk header
        Set wsDestination = Nothing
        On Error Resume Next
        Set wsDestination = ThisWorkbook.Sheets(wsSource.Cells(i, "E").Value)
        On Error GoTo 0
        
        If Not wsDestination Is Nothing Then
            wsDestination.Range(wsSource.Cells(i, "F").Value).Formula = wsSource.Cells(i, "D").Formula
        Else
            MsgBox "Sheet tujuan '" & wsSource.Cells(i, "E").Value & "' tidak ditemukan.", vbExclamation
        End If
    Next i
End Sub
