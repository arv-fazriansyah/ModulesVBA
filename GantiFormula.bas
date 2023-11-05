Sub CopyFormulas()
    Dim sourceSheetName As String
    sourceSheetName = "Sheet1"  ' Ganti dengan nama sheet sumber yang diinginkan
    
    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set worksheet sumber
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Sheet sumber '" & sourceSheetName & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If
    
    ' Kolom-kolom yang bisa diatur
    Dim kolomFormula As String
    Dim kolomTujuanSheet As String
    Dim kolomTujuanCell As String
    
    kolomFormula = "D"  ' Kolom yang berisi formula
    kolomTujuanSheet = "E"  ' Kolom yang berisi nama sheet tujuan
    kolomTujuanCell = "F"  ' Kolom yang berisi sel tujuan
    
    ' Loop melalui setiap baris di kolom D
    lastRow = wsSource.Cells(wsSource.Rows.Count, kolomFormula).End(xlUp).Row
    
    Dim wsDestination As Worksheet
    Dim destinationSheetName As String
    Dim destinationCell As String
    
    For i = 2 To lastRow ' Anggap baris 1 adalah untuk header
        ' Dapatkan nama sheet tujuan dan sel tujuan dari kolom E dan F
        destinationSheetName = wsSource.Cells(i, kolomTujuanSheet).Value
        destinationCell = wsSource.Cells(i, kolomTujuanCell).Value
        
        ' Periksa apakah sheet tujuan ada
        Set wsDestination = Nothing
        On Error Resume Next
        Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)
        On Error GoTo 0
        
        If wsDestination Is Nothing Then
            MsgBox "Sheet tujuan '" & destinationSheetName & "' tidak ditemukan.", vbExclamation
        Else
            ' Salin formula dari kolom D ke sel yang ditentukan di sheet tujuan
            wsDestination.Range(destinationCell).Formula = wsSource.Cells(i, kolomFormula).Formula
        End If
    Next i
End Sub
