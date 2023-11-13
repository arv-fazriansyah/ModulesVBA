Sub CopyFormulas()
    ' Deklarasi variabel
    Dim namaSheetSumber As String
    Dim rumusKolomSumber As String
    Dim namaSheetTujuanKolom As String
    Dim selTujuanKolom As String
    
    ' Inisialisasi nilai variabel
    namaSheetSumber = "DATAUSER"
    rumusKolomSumber = "AA"
    namaSheetTujuanKolom = "AB"
    selTujuanKolom = "AC"
    
    ' Pesan kesalahan
    UpdateErrorMsg = "Download ulang Aplikasi, hubungi Admin"
    
    On Error GoTo RefreshError
    
    ' Mengatur sumber data
    Dim sheetSumber As Worksheet
    On Error Resume Next
    Set sheetSumber = ThisWorkbook.Sheets(namaSheetSumber)
    
    ' Memeriksa keberadaan lembar kerja
    If sheetSumber Is Nothing Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    ' Mendapatkan pemisah yang sesuai dengan pengaturan lokal
    Dim pemisah As String
    pemisah = Application.International(xlListSeparator)
    
    ' Menentukan baris terakhir
    Dim barisTerakhir As Long
    barisTerakhir = sheetSumber.Cells(sheetSumber.Rows.Count, rumusKolomSumber).End(xlUp).Row
    
    ' Iterasi baris
    Dim i As Long
    For i = 1 To barisTerakhir
        ' Mendapatkan rumus
        Dim nilaiRumus As String
        nilaiRumus = sheetSumber.Cells(i, rumusKolomSumber).Formula
        nilaiRumus = Replace(nilaiRumus, ";", pemisah)
        nilaiRumus = Replace(nilaiRumus, ",", pemisah)
        
        ' Mendapatkan nama lembar tujuan dan sel tujuan
        Dim namaLembarTujuan As String
        namaLembarTujuan = sheetSumber.Cells(i, namaSheetTujuanKolom).Value
        
        Dim selTujuan As String
        selTujuan = sheetSumber.Cells(i, selTujuanKolom).Value
        
        ' Memeriksa apakah data ada
        If namaLembarTujuan <> "" And selTujuan <> "" Then
            ' Mengatur lembar tujuan
            Dim lembarTujuan As Worksheet
            On Error Resume Next
            Set lembarTujuan = ThisWorkbook.Sheets(namaLembarTujuan)
            On Error GoTo 0
            
            ' Memeriksa keberadaan lembar tujuan
            If Not lembarTujuan Is Nothing Then
                ' Mendapatkan password lembar tujuan
                Dim PasswordSheetTujuan As String
                PasswordSheetTujuan = sheetSumber.Cells(i, "AD").Value
                
                ' Mengecek dan memproteksi lembar tujuan
                If PasswordSheetTujuan <> "" Then
                    On Error Resume Next
                    lembarTujuan.Unprotect PasswordSheetTujuan
                    On Error GoTo 0
                    If lembarTujuan.ProtectContents Then
                        MsgBox "Password lembar tujuan salah!", vbExclamation
                        Exit Sub
                    End If
                ElseIf lembarTujuan.ProtectContents Then
                    MsgBox "Lembar terlindungi. Masukkan password!", vbExclamation
                    Exit Sub
                End If
                
                ' Memasukkan nilai rumus ke lembar tujuan
                Application.DisplayAlerts = False
                lembarTujuan.Range(selTujuan).Value = nilaiRumus
                Application.DisplayAlerts = True
                
                ' Menghapus tautan Excel
                Dim tautan As Variant
                tautan = ThisWorkbook.LinkSources(xlExcelLinks)
                
                If Not IsEmpty(tautan) Then
                    Dim j As Long
                    For j = 1 To UBound(tautan)
                        ThisWorkbook.BreakLink Name:=tautan(j), Type:=xlLinkTypeExcelLinks
                    Next j
                End If
                
                ' Melindungi lembar setelah selesai memasukkan nilai
                If PasswordSheetTujuan <> "" Then
                    lembarTujuan.Protect PasswordSheetTujuan
                End If
            End If
        End If
    Next i
    Exit Sub
    
RefreshError:
    MsgBox UpdateErrorMsg, vbExclamation
End Sub
