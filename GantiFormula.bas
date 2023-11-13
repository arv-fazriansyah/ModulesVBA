Sub CopyFormulasToDestinationSheet()
    Dim sourceSheetName As String
    Dim sourceColumnFormula As String
    Dim destinationColumnSheet As String
    Dim destinationColumnCell As String
    
    sourceSheetName = "DATAUSER"
    sourceColumnFormula = "H"
    destinationColumnSheet = "I"
    destinationColumnCell = "J"
    
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets(sourceSheetName)
    
    Dim separator As String
    ' Dapatkan tanda pemisah dari pengaturan regional pengguna
    separator = Application.International(xlListSeparator)
    
    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, sourceColumnFormula).End(xlUp).row
    
    Dim i As Long
    For i = 1 To lastRow
        ' Ambil formula dari kolom H
        Dim formulaValue As String
        formulaValue = sourceSheet.Cells(i, sourceColumnFormula).formula
        
        ' Sesuaikan tanda pemisah formula dengan tanda pemisah regional
        formulaValue = Replace(formulaValue, ";", separator)
        formulaValue = Replace(formulaValue, ",", separator)
        
        ' Ambil nama sheet tujuan dari kolom I
        Dim destSheetName As String
        destSheetName = sourceSheet.Cells(i, destinationColumnSheet).value
        
        ' Ambil sel tujuan dari kolom J
        Dim destCell As String
        destCell = sourceSheet.Cells(i, destinationColumnCell).value
        
        ' Cek jika nama sheet tujuan tidak kosong dan sel tujuan tidak kosong
        If destSheetName <> "" And destCell <> "" Then
            Dim destSheet As Worksheet
            On Error Resume Next
            Set destSheet = ThisWorkbook.Sheets(destSheetName)
            On Error GoTo 0
            
            If Not destSheet Is Nothing Then
                ' Tempelkan nilai sebagai teks ke sel tujuan pada sheet tujuan
                Application.DisplayAlerts = False
                destSheet.Range(destCell).value = formulaValue
                Application.DisplayAlerts = True
            End If
        End If
    Next i
    
End Sub
