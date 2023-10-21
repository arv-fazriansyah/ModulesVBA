Attribute VB_Name = "GantiFormula"
Sub UpdateFormulas()
    Dim databaseSheetName As String
    Dim formulaColumn As String
    Dim sheetTujuanColumn As String
    Dim cellTujuanColumn As String
    Dim wsDatabase As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Tentukan variabel-variabel untuk nama sheet "Database" dan kolom data.
    databaseSheetName = "Database"
    formulaColumn = "A"
    sheetTujuanColumn = "B"
    cellTujuanColumn = "C"
    
    ' Cek apakah sheet "Database" ada.
    If Not WorksheetExists(databaseSheetName) Then
        MsgBox "Sheet '" & databaseSheetName & "' tidak ditemukan!", vbExclamation, "Kesalahan"
        Exit Sub ' Menghentikan eksekusi skrip jika sheet "Database" tidak ditemukan.
    End If
    
    ' Ganti "Database" dengan nama sheet Anda yang berisi data formula, sheet tujuan, dan cell tujuan.
    Set wsDatabase = ThisWorkbook.Worksheets(databaseSheetName)
    
    ' Temukan jumlah baris data di sheet database (asumsi data dimulai dari baris 2).
    lastRow = wsDatabase.Cells(wsDatabase.Rows.Count, formulaColumn).End(xlUp).Row
    
    Dim sheetTujuanExists As Boolean
    sheetTujuanExists = True
    Dim missingSheetNames As String ' Ini akan digunakan untuk menyimpan nama-nama sheet yang tidak ditemukan.
    
    ' Loop melalui setiap baris data di sheet database.
    For i = 2 To lastRow
        Dim formula As String
        Dim sheetTujuan As String
        Dim cellTujuan As String
        
        ' Ambil formula, sheet tujuan, dan cell tujuan dari kolom A, B, dan C.
        formula = wsDatabase.Cells(i, formulaColumn).Value
        sheetTujuan = wsDatabase.Cells(i, sheetTujuanColumn).Value
        cellTujuan = wsDatabase.Cells(i, cellTujuanColumn).Value
        
        ' Cek apakah sheet tujuan sudah ada.
        If Not WorksheetExists(sheetTujuan) Then
            sheetTujuanExists = False
            missingSheetNames = missingSheetNames & sheetTujuan & ", "
        End If
    Next i
    
    If Not sheetTujuanExists Then
        ' Hapus koma dan spasi terakhir dari daftar nama sheet yang hilang.
        missingSheetNames = Left(missingSheetNames, Len(missingSheetNames) - 2)
        MsgBox "Sheet tujuan (" & missingSheetNames & ") tidak ditemukan!", vbExclamation, "Kesalahan"
    Else
        ' Lanjutkan dengan pembaruan formula jika semua sheet tujuan ada.
        For i = 2 To lastRow
            ' Ambil kembali formula, sheet tujuan, dan cell tujuan dari kolom A, B, dan C.
            formula = wsDatabase.Cells(i, formulaColumn).Value
            sheetTujuan = wsDatabase.Cells(i, sheetTujuanColumn).Value
            cellTujuan = wsDatabase.Cells(i, cellTujuanColumn).Value
            
            ' Coba menerapkan formula ke sel yang ditentukan.
            On Error Resume Next
            ThisWorkbook.Worksheets(sheetTujuan).Range(cellTujuan).formula = formula
            On Error GoTo 0
            
            ' Periksa kesalahan dan laporkan jika ditemukan.
            If Err.Number <> 0 Then
                MsgBox "Kesalahan pada baris " & i & ": " & Err.Description, vbExclamation, "Kesalahan"
                Err.Clear
            End If
        Next i
        
        ' Beri tahu pengguna bahwa proses telah selesai.
        MsgBox "Pembaruan formula telah selesai!", vbInformation, "Selesai"
    End If
End Sub

Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Worksheets(wsName) Is Nothing
    On Error GoTo 0
End Function

