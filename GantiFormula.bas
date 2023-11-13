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
        MsgBox "Unduh ulang aplikasi dan hubungi Admin!", vbExclamation
        Exit Sub
    End If
    
    Dim pemisah As String
    ' Dapatkan pemisah daftar dari pengaturan regional pengguna
    pemisah = Application.International(xlListSeparator)
    
    Dim barisTerakhir As Long
    barisTerakhir = sheetSumber.Cells(sheetSumber.Rows.count, rumusKolomSumber).End(xlUp).row
    
    Dim i As Long
    For i = 1 To barisTerakhir
        ' Dapatkan rumus dari kolom H
        Dim nilaiRumus As String
        nilaiRumus = sheetSumber.Cells(i, rumusKolomSumber).formula
        
        ' Ganti pemisah rumus dengan pemisah regional
        nilaiRumus = Replace(nilaiRumus, ";", pemisah)
        nilaiRumus = Replace(nilaiRumus, ",", pemisah)
        
        ' Dapatkan nama lembar tujuan dari kolom I
        Dim namaLembarTujuan As String
        namaLembarTujuan = sheetSumber.Cells(i, namaSheetTujuanKolom).value
        
        ' Dapatkan sel tujuan dari kolom J
        Dim selTujuan As String
        selTujuan = sheetSumber.Cells(i, selTujuanKolom).value
        
        ' Periksa apakah nama lembar tujuan dan sel tidak kosong
        If namaLembarTujuan <> "" And selTujuan <> "" Then
            Dim lembarTujuan As Worksheet
            On Error Resume Next
            Set lembarTujuan = ThisWorkbook.Sheets(namaLembarTujuan)
            On Error GoTo 0
            
            If Not lembarTujuan Is Nothing Then
                ' Unprotect worksheet dengan kata sandi jika diperlukan
                Dim password As String
                password = "" ' Ganti dengan kata sandi yang benar
                
                ' Pengecekan password dan lembar terlindungi
                If password <> "" Then
                    On Error Resume Next
                    lembarTujuan.Unprotect password
                    On Error GoTo 0
                    If lembarTujuan.ProtectContents Then
                        MsgBox "Kata sandi salah!", vbExclamation
                        Exit Sub
                    End If
                ElseIf lembarTujuan.ProtectContents Then
                    MsgBox "Sheet terlindungi, masukan kata sandi!", vbExclamation
                    Exit Sub
                End If
                
                ' Tempelkan nilai sebagai teks ke sel tujuan di lembar tujuan
                Application.DisplayAlerts = False
                lembarTujuan.Range(selTujuan).value = nilaiRumus
                Application.DisplayAlerts = True

                ' Proses setelah paste formula
                If password <> "" Then
                    lembarTujuan.Protect password
                End If
            End If
        End If
    Next i

    ' Hapus tautan buku kerja eksternal
    Dim tautan As Variant
    tautan = ThisWorkbook.LinkSources(xlExcelLinks)
    
    If Not IsEmpty(tautan) Then
        For i = 1 To UBound(tautan)
            ThisWorkbook.BreakLink Name:=tautan(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub
