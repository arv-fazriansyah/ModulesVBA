Sub CopyFormulas()
    Dim sourceSheetName As String
    Dim sourceColumnFormula As String
    Dim destinationColumnSheet As String
    Dim destinationColumnCell As String
    
    sourceSheetName = "Sheet1"  ' Ganti dengan nama sheet sumber yang diinginkan
    sourceColumnFormula = "D"    ' Kolom yang berisi formula
    destinationColumnSheet = "E"  ' Kolom yang berisi nama sheet tujuan
    destinationColumnCell = "F"  ' Kolom yang berisi sel tujuan
    
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Sheet sumber '" & sourceSheetName & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, sourceColumnFormula).End(xlUp).Row
    
    For i = 2 To lastRow
        Set wsDestination = Nothing
        On Error Resume Next
        Set wsDestination = ThisWorkbook.Sheets(wsSource.Cells(i, destinationColumnSheet).Value)
        On Error GoTo 0
        
        If Not wsDestination Is Nothing Then
            wsDestination.Range(wsSource.Cells(i, destinationColumnCell).Value).Formula = wsSource.Cells(i, sourceColumnFormula).Formula
        Else
            MsgBox "Sheet tujuan '" & wsSource.Cells(i, destinationColumnSheet).Value & "' tidak ditemukan.", vbExclamation
        End If
    Next i
End Sub
