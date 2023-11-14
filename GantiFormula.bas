Sub CopyFormulas()
    Dim namaSheetSumber As String
    Dim rumusKolomSumber As String
    Dim namaSheetTujuanKolom As String
    Dim selTujuanKolom As String
    Dim PasswordSheetTujuanKolom As String ' Variabel yang ditambahkan
    
    namaSheetSumber = "DATAUSER"
    rumusKolomSumber = "AA"
    namaSheetTujuanKolom = "AB"
    selTujuanKolom = "AC"
    PasswordSheetTujuanKolom = "AD" ' Kolom kata sandi
    
    Dim sheetSumber As Worksheet
    ' Penanganan kesalahan untuk lembar sumber
    On Error Resume Next
    Set sheetSumber = ThisWorkbook.Sheets(namaSheetSumber)
    On Error GoTo 0
    
    If sheetSumber Is Nothing Then
        MsgBox "Lembar sumber tidak ditemukan. Hubungi Admin!", vbExclamation
        Exit Sub
    End If
    
    Dim pemisah As String
    pemisah = Application.International(xlListSeparator)
    
    Dim barisTerakhir As Long
    barisTerakhir = sheetSumber.Cells(sheetSumber.Rows.Count, rumusKolomSumber).End(xlUp).Row
    
    Dim i As Long
    For i = 1 To barisTerakhir
        Dim nilaiRumus As String
        nilaiRumus = sheetSumber.Cells(i, rumusKolomSumber).Formula
        nilaiRumus = Replace(nilaiRumus, ";", pemisah)
        nilaiRumus = Replace(nilaiRumus, ",", pemisah)
        
        Dim namaLembarTujuan As String
        namaLembarTujuan = sheetSumber.Cells(i, namaSheetTujuanKolom).Value
        
        Dim selTujuan As String
        selTujuan = sheetSumber.Cells(i, selTujuanKolom).Value
        
        Dim passwordLembarTujuan As String
        passwordLembarTujuan = sheetSumber.Cells(i, PasswordSheetTujuanKolom).Value ' Ambil kata sandi
        
        If namaLembarTujuan <> "" And selTujuan <> "" Then
            Dim lembarTujuan As Worksheet
            On Error Resume Next
            Set lembarTujuan = ThisWorkbook.Sheets(namaLembarTujuan)
            On Error GoTo 0
            
            If Not lembarTujuan Is Nothing Then
                If passwordLembarTujuan <> "" Then
                    On Error Resume Next
                    lembarTujuan.Unprotect passwordLembarTujuan
                    On Error GoTo 0
                    If lembarTujuan.ProtectContents Then
                        MsgBox "Kata sandi lembar tujuan salah!", vbExclamation
                        Exit Sub
                    End If
                ElseIf lembarTujuan.ProtectContents Then
                    MsgBox "Lembar terlindungi. Masukkan kata sandi!", vbExclamation
                    Exit Sub
                End If
                
                Application.DisplayAlerts = False
                lembarTujuan.Range(selTujuan).Value = nilaiRumus
                Application.DisplayAlerts = True
                
                Dim tautan As Variant
                tautan = ThisWorkbook.LinkSources(xlExcelLinks)
                
                If Not IsEmpty(tautan) Then
                    Dim j As Long
                    For j = 1 To UBound(tautan)
                        ThisWorkbook.BreakLink Name:=tautan(j), Type:=xlLinkTypeExcelLinks
                    Next j
                End If
                
                ' Lindungi lembar tujuan setelah mengisi nilai
                If passwordLembarTujuan <> "" Then
                    lembarTujuan.Protect passwordLembarTujuan
                End If
            End If
        End If
    Next i
End Sub
