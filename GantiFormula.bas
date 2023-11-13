Sub CopyFormulas()
    Dim namaSheetSumber As String
    Dim rumusKolomSumber As String
    Dim namaSheetTujuanKolom As String
    Dim selTujuanKolom As String
    
    namaSheetSumber = "DATAUSER"
    rumusKolomSumber = "H"
    namaSheetTujuanKolom = "I"
    selTujuanKolom = "J"
    
    Dim sheetSumber As Worksheet
    On Error Resume Next
    Set sheetSumber = ThisWorkbook.Sheets(namaSheetSumber)
    On Error GoTo 0
    
    If sheetSumber Is Nothing Then
        MsgBox "Sheet sumber tidak ditemukan. Hubungi Admin!", vbExclamation
        Exit Sub
    End If
    
    Dim pemisah As String
    pemisah = Application.International(xlListSeparator)
    
    Dim barisTerakhir As Long
    barisTerakhir = sheetSumber.Cells(sheetSumber.Rows.count, rumusKolomSumber).End(xlUp).row
    
    Dim i As Long
    For i = 1 To barisTerakhir
        Dim nilaiRumus As String
        nilaiRumus = sheetSumber.Cells(i, rumusKolomSumber).formula
        nilaiRumus = Replace(nilaiRumus, ";", pemisah)
        nilaiRumus = Replace(nilaiRumus, ",", pemisah)
        
        Dim namaLembarTujuan As String
        namaLembarTujuan = sheetSumber.Cells(i, namaSheetTujuanKolom).value
        
        Dim selTujuan As String
        selTujuan = sheetSumber.Cells(i, selTujuanKolom).value
        
        If namaLembarTujuan <> "" And selTujuan <> "" Then
            Dim lembarTujuan As Worksheet
            On Error Resume Next
            Set lembarTujuan = ThisWorkbook.Sheets(namaLembarTujuan)
            On Error GoTo 0
            
            If Not lembarTujuan Is Nothing Then
                Dim password As String
                password = "" ' Ganti dengan kata sandi yang benar
                
                If password <> "" Then
                    On Error Resume Next
                    lembarTujuan.Unprotect password
                    On Error GoTo 0
                    If lembarTujuan.ProtectContents Then
                        MsgBox "Kata sandi salah!", vbExclamation
                        Exit Sub
                    End If
                ElseIf lembarTujuan.ProtectContents Then
                    MsgBox "Lembar terlindungi, masukkan kata sandi!", vbExclamation
                    Exit Sub
                End If
                
                Application.DisplayAlerts = False
                lembarTujuan.Range(selTujuan).value = nilaiRumus
                Application.DisplayAlerts = True
                
                Dim tautan As Variant
                tautan = ThisWorkbook.LinkSources(xlExcelLinks)
                
                If Not IsEmpty(tautan) Then
                    Dim j As Long
                    For j = 1 To UBound(tautan)
                        ThisWorkbook.BreakLink Name:=tautan(j), Type:=xlLinkTypeExcelLinks
                    Next j
                End If
                
                ' Hanya proteksi setelah selesai mengisi nilai
                If password <> "" Then
                    lembarTujuan.Protect password
                End If
            End If
        End If
    Next i
End Sub
