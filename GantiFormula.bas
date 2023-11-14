Sub CopyFormulas()
    Dim namaSheetSumber As String
    Dim rumusKolomSumber As String
    Dim namaSheetTujuanKolom As String
    Dim selTujuanKolom As String
    Dim PasswordSheetTujuanKolom As String
    
    namaSheetSumber = Env.DataBase
    rumusKolomSumber = "AA"
    namaSheetTujuanKolom = "AB"
    selTujuanKolom = "AC"
    PasswordSheetTujuanKolom = "AD"
    
    Dim sheetSumber As Worksheet
    On Error Resume Next
    Set sheetSumber = ThisWorkbook.Sheets(namaSheetSumber)
    On Error GoTo 0
    
    If sheetSumber Is Nothing Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
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
        
        Dim passwordLembarTujuan As String
        passwordLembarTujuan = sheetSumber.Cells(i, PasswordSheetTujuanKolom).value ' Ambil kata sandi
        
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
                        MsgBox "Kata sandi untuk " & lembarTujuan.Name & " salah!", vbExclamation
                        Exit Sub
                    End If
                ElseIf lembarTujuan.ProtectContents Then
                    MsgBox "Lembar " & lembarTujuan.Name & " terlindungi. Masukkan kata sandi!", vbExclamation
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
                
                ' Lindungi lembar tujuan setelah mengisi nilai
                If passwordLembarTujuan <> "" Then
                    lembarTujuan.Protect passwordLembarTujuan
                End If
            End If
        End If
    Next i
End Sub
