Sub SumColor1()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim sumFormula As String
    Dim nextColorRow As Long
    Dim colorToCheck As Long
    Dim colorToSum As Long
    Dim columnsToFill As Variant
    Dim col As Variant
    Dim i As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual

    ' Pengaturan yang mudah diubah
    Set ws = ThisWorkbook.Sheets("RBK") ' Ganti dengan nama sheet yang sesuai
    startRow = 17                        ' Baris awal
    colorToCheck = RGB(237, 125, 49)     ' Warna yang diperiksa di kolom F
    colorToSum = RGB(189, 215, 238)      ' Warna yang akan dijumlahkan
    columnsToFill = Array("G", "O", "W", "AE", "AM", "AU", "BC", "BK", "BS", "CA", "CI", "CQ", "CY") ' Kolom yang akan diisi

    ' Menentukan endRow berdasarkan kolom B yang terisi
    endRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For currentRow = startRow To endRow
        ' Cek jika sel di kolom F berwarna dan tidak kosong
        If ws.Cells(currentRow, "F").Interior.Color = colorToCheck And _
           ws.Cells(currentRow, "F").value <> "" Then
            
            ' Cari baris yang berwarna RGB(237, 125, 49) berikutnya
            nextColorRow = currentRow + 1
            
            Do While nextColorRow <= endRow
                If ws.Cells(nextColorRow, "F").Interior.Color = colorToCheck Then
                    Exit Do
                End If
                nextColorRow = nextColorRow + 1
            Loop
            
            ' Buat formula SUM untuk setiap kolom yang ditentukan
            For Each col In columnsToFill
                sumFormula = "=SUM("
                For i = currentRow + 1 To nextColorRow - 1
                    If ws.Cells(i, col).Interior.Color = colorToSum Then
                        sumFormula = sumFormula & col & i & ","
                    End If
                Next i
                
                ' Hapus koma terakhir dan tambahkan kurung tutup
                If Len(sumFormula) > 5 Then
                    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
                    ws.Cells(currentRow, col).Formula = sumFormula
                End If
            Next col
        End If
    Next currentRow
    
    ' Memanggil subroutine lainnya
    SumColor2
    SumColor3
    SumColor4
    FormulaRBK2
    FormulaRBK3

    ' Mengaktifkan kembali pengaturan aplikasi
    Application.EnableEvents = True
    Application.EnableAnimations = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub SumColor2()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim sumFormula As String
    Dim nextColorRow As Long
    Dim colorToCheck As Long
    Dim colorToSum As Long
    Dim columnsToFill As Variant
    Dim col As Variant
    Dim i As Long

    ' Pengaturan yang mudah diubah
    Set ws = ThisWorkbook.Sheets("RBK") ' Ganti dengan nama sheet yang sesuai
    startRow = 17                        ' Baris awal
    colorToCheck = RGB(189, 215, 238)    ' Warna yang diperiksa di kolom F
    colorToSum = RGB(255, 255, 153)      ' Warna yang akan dijumlahkan
    columnsToFill = Array("G", "O", "W", "AE", "AM", "AU", "BC", "BK", "BS", "CA", "CI", "CQ", "CY") ' Kolom yang akan diisi

    ' Menentukan endRow berdasarkan kolom B yang terisi
    endRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    'MsgBox "Jumlah endRow: " & endRow

    For currentRow = startRow To endRow
        ' Cek jika sel di kolom F berwarna dan tidak kosong
        If ws.Cells(currentRow, "F").Interior.Color = colorToCheck And _
           ws.Cells(currentRow, "F").value <> "" Then
            
            ' Cari baris yang berwarna RGB(237, 125, 49) berikutnya
            nextColorRow = currentRow + 1
            
            Do While nextColorRow <= endRow
                If ws.Cells(nextColorRow, "F").Interior.Color = colorToCheck Then
                    Exit Do
                End If
                nextColorRow = nextColorRow + 1
            Loop
            
            ' Buat formula SUM untuk setiap kolom yang ditentukan
            For Each col In columnsToFill
                sumFormula = "=SUM("
                For i = currentRow + 1 To nextColorRow - 1
                    If ws.Cells(i, col).Interior.Color = colorToSum Then
                        sumFormula = sumFormula & col & i & ","
                    End If
                Next i
                
                ' Hapus koma terakhir dan tambahkan kurung tutup
                If Len(sumFormula) > 5 Then
                    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
                    ws.Cells(currentRow, col).Formula = sumFormula
                End If
            Next col
        End If
    Next currentRow
End Sub
Sub SumColor3()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim sumFormula As String
    Dim nextColorRow As Long
    Dim colorToCheck As Long
    Dim colorToSum As Long
    Dim columnsToFill As Variant
    Dim col As Variant
    Dim i As Long

    ' Pengaturan yang mudah diubah
    Set ws = ThisWorkbook.Sheets("RBK") ' Ganti dengan nama sheet yang sesuai
    startRow = 17                        ' Baris awal
    colorToCheck = RGB(255, 255, 153)     ' Warna yang diperiksa di kolom F
    colorToSum = RGB(217, 217, 217)      ' Warna yang akan dijumlahkan
    columnsToFill = Array("G", "O", "W", "AE", "AM", "AU", "BC", "BK", "BS", "CA", "CI", "CQ", "CY") ' Kolom yang akan diisi

    ' Menentukan endRow berdasarkan kolom B yang terisi
    endRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For currentRow = startRow To endRow
        ' Cek jika sel di kolom F berwarna dan tidak kosong
        If ws.Cells(currentRow, "F").Interior.Color = colorToCheck And _
           ws.Cells(currentRow, "F").value <> "" Then
            
            ' Cari baris yang berwarna RGB(255, 255, 153) berikutnya
            nextColorRow = currentRow + 1
            
            Do While nextColorRow <= endRow
                If ws.Cells(nextColorRow, "F").Interior.Color = colorToCheck Then
                    Exit Do
                End If
                nextColorRow = nextColorRow + 1
            Loop
            
            ' Buat formula SUM untuk setiap kolom yang ditentukan
            For Each col In columnsToFill
                sumFormula = "=SUM("
                For i = currentRow + 1 To nextColorRow - 1
                    If ws.Cells(i, col).Interior.Color = colorToSum Then
                        sumFormula = sumFormula & col & i & ","
                    End If
                Next i
                
                ' Hapus koma terakhir dan tambahkan kurung tutup
                If Len(sumFormula) > 5 Then
                    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
                    ws.Cells(currentRow, col).Formula = sumFormula
                End If
            Next col
        End If
    Next currentRow
End Sub
Sub SumColor4()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim sumFormula As String
    Dim nextColorRow As Long
    Dim colorToCheck As Long
    Dim colorToSum As Long
    Dim columnsToFill As Variant
    Dim col As Variant
    Dim i As Long

    ' Pengaturan yang mudah diubah
    Set ws = ThisWorkbook.Sheets("RBK") ' Ganti dengan nama sheet yang sesuai
    startRow = 17                        ' Baris awal
    colorToCheck = RGB(217, 217, 217)    ' Warna yang diperiksa di kolom F
    colorToSum = RGB(255, 255, 255)      ' Warna yang akan dijumlahkan
    columnsToFill = Array("G", "O", "W", "AE", "AM", "AU", "BC", "BK", "BS", "CA", "CI", "CQ", "CY") ' Kolom yang akan diisi

    ' Menentukan endRow berdasarkan kolom B yang terisi
    endRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For currentRow = startRow To endRow
        ' Cek jika sel di kolom F berwarna dan tidak kosong
        If ws.Cells(currentRow, "F").Interior.Color = colorToCheck And _
           ws.Cells(currentRow, "F").value <> "" Then
            
            ' Cari baris yang berwarna RGB(217, 217, 217) berikutnya
            nextColorRow = currentRow + 1
            
            Do While nextColorRow <= endRow
                If ws.Cells(nextColorRow, "F").Interior.Color = colorToCheck Then
                    Exit Do
                End If
                nextColorRow = nextColorRow + 1
            Loop
            
            ' Buat formula SUM untuk setiap kolom yang ditentukan
            For Each col In columnsToFill
                sumFormula = "=SUM("
                For i = currentRow + 1 To nextColorRow - 1
                    If ws.Cells(i, col).Interior.Color = colorToSum Then
                        sumFormula = sumFormula & col & i & ","
                    End If
                Next i
                
                ' Hapus koma terakhir dan tambahkan kurung tutup
                If Len(sumFormula) > 5 Then
                    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
                    ws.Cells(currentRow, col).Formula = sumFormula
                End If
            Next col
        End If
    Next currentRow
End Sub

Sub FormulaRBK2()
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long
    Dim checkColumn As String, formulaColumn As String
    Dim formulaCells As String
    Dim colorWhite As Long
    Dim i As Long
    Dim rngCheck As Range
    Dim formulas() As Variant
    
    ' Pengaturan yang mudah diubah
    Set ws = ThisWorkbook.Sheets("RBK") ' Sheet tempat data berada
    startRow = 21                        ' Baris awal
    endRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Menentukan endRow berdasarkan kolom B
    checkColumn = "F"                    ' Kolom yang diperiksa
    formulaColumn = "G"                  ' Kolom untuk menerapkan formula
    formulaCells = "O,W,AE,AM,AU,BC,BK,BS,CA,CI,CQ,CY" ' Kolom-kolom yang disertakan dalam formula SUM
    colorWhite = RGB(255, 255, 255)      ' Warna putih RGB 255

    ' Menyimpan hasil formula dalam array
    ReDim formulas(startRow To endRow)
    
    ' Loop dari baris startRow hingga endRow
    For i = startRow To endRow
        ' Setel range untuk kolom yang diperiksa
        Set rngCheck = ws.Range(checkColumn & i)
        
        ' Jika warna background rngCheck putih dan nilainya tidak kosong
        If rngCheck.Interior.Color = colorWhite And rngCheck.value <> "" Then
            ' Buat formula SUM berdasarkan kolom yang ditentukan
            formulas(i) = "=SUM(" & Replace(formulaCells, ",", i & ",") & i & ")"
        ElseIf rngCheck.Interior.Color = colorWhite Then
            ' Jika berwarna putih tetapi kosong, kosongkan formula
            formulas(i) = ""
        Else
            ' Jika bukan berwarna putih, biarkan nilai formula yang ada
            formulas(i) = ws.Range(formulaColumn & i).Formula ' Simpan formula yang ada
        End If
    Next i
    
    ' Menulis hasil ke Excel
    ws.Range(formulaColumn & startRow & ":" & formulaColumn & endRow).value = Application.Transpose(formulas)
    
End Sub

Sub FormulaRBK3()
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long
    Dim checkColumn As String
    Dim formulaColumns As Variant
    Dim colorWhite As Long
    Dim i As Long
    Dim j As Long
    
    ' Pengaturan yang mudah diubah
    Set ws = ThisWorkbook.Sheets("RBK") ' Sheet tempat data berada
    startRow = 21                        ' Baris awal
    endRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Menentukan endRow berdasarkan kolom B
    checkColumn = "F"                    ' Kolom yang diperiksa
    colorWhite = RGB(255, 255, 255)      ' Warna putih RGB 255
    formulaColumns = Array("O", "W", "AE", "AM", "AU", "BC", "BK", "BS", "CA", "CI", "CQ", "CY") ' Kolom tempat formula disimpan

    ' Loop dari baris startRow hingga endRow
    For i = startRow To endRow
        ' Periksa apakah warna background kolom F putih dan tidak kosong
        If ws.Range(checkColumn & i).Interior.Color = colorWhite And ws.Range(checkColumn & i).value <> "" Then
            ' Terapkan formula berdasarkan kolom yang ditentukan
            For j = LBound(formulaColumns) To UBound(formulaColumns)
                ' Hanya mengubah formula jika sel formula saat ini kosong
                If ws.Range(formulaColumns(j) & i).Formula = "" Then
                    Select Case formulaColumns(j)
                        Case "O"
                            ws.Range(formulaColumns(j) & i).Formula = "=H" & i & "*J" & i & "*L" & i & "*N" & i
                        Case "W"
                            ws.Range(formulaColumns(j) & i).Formula = "=P" & i & "*R" & i & "*T" & i & "*V" & i
                        Case "AE"
                            ws.Range(formulaColumns(j) & i).Formula = "=X" & i & "*Z" & i & "*AB" & i & "*AD" & i
                        Case "AM"
                            ws.Range(formulaColumns(j) & i).Formula = "=AF" & i & "*AH" & i & "*AJ" & i & "*AL" & i
                        Case "AU"
                            ws.Range(formulaColumns(j) & i).Formula = "=AN" & i & "*AP" & i & "*AR" & i & "*AT" & i
                        Case "BC"
                            ws.Range(formulaColumns(j) & i).Formula = "=AV" & i & "*AX" & i & "*AZ" & i & "*BB" & i
                        Case "BK"
                            ws.Range(formulaColumns(j) & i).Formula = "=BD" & i & "*BF" & i & "*BH" & i & "*BJ" & i
                        Case "BS"
                            ws.Range(formulaColumns(j) & i).Formula = "=BL" & i & "*BN" & i & "*BP" & i & "*BR" & i
                        Case "CA"
                            ws.Range(formulaColumns(j) & i).Formula = "=BT" & i & "*BV" & i & "*BX" & i & "*BZ" & i
                        Case "CI"
                            ws.Range(formulaColumns(j) & i).Formula = "=CB" & i & "*CD" & i & "*CF" & i & "*CH" & i
                        Case "CQ"
                            ws.Range(formulaColumns(j) & i).Formula = "=CJ" & i & "*CL" & i & "*CN" & i & "*CP" & i
                        Case "CY"
                            ws.Range(formulaColumns(j) & i).Formula = "=CR" & i & "*CT" & i & "*CV" & i & "*CX" & i
                    End Select
                End If
            Next j
        End If
    Next i
End Sub
