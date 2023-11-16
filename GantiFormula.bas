Sub CopyFormulas()
    Dim NamaSheetSumber As String
    Dim RumusKolomSumber As String
    Dim NamaSheetTujuanKolom As String
    Dim SelTujuanKolom As String
    Dim PasswordSheetTujuanKolom As String
    
    NamaSheetSumber = Env.DataBase
    RumusKolomSumber = "AA"
    NamaSheetTujuanKolom = "AB"
    SelTujuanKolom = "AC"
    PasswordSheetTujuanKolom = "AD"
    
    Dim SheetSumber As Worksheet
    On Error Resume Next
    Set SheetSumber = ThisWorkbook.Sheets(NamaSheetSumber)
    ' On Error GoTo 0 ' Hapus baris ini
    
    If SheetSumber Is Nothing Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    Dim pemisah As String
    pemisah = Application.International(xlListSeparator)
    
    Dim barisTerakhir As Long
    barisTerakhir = SheetSumber.Cells(SheetSumber.Rows.count, RumusKolomSumber).End(xlUp).row
    
    Dim i As Long
    For i = 2 To barisTerakhir
        Dim nilaiRumus As String
        nilaiRumus = SheetSumber.Cells(i, RumusKolomSumber).formula
        nilaiRumus = Replace(nilaiRumus, ";", pemisah)
        nilaiRumus = Replace(nilaiRumus, ",", pemisah)
        
        Dim namaLembarTujuan As String
        namaLembarTujuan = SheetSumber.Cells(i, NamaSheetTujuanKolom).value
        
        Dim selTujuan As String
        selTujuan = SheetSumber.Cells(i, SelTujuanKolom).value
        
        Dim passwordLembarTujuan As String
        passwordLembarTujuan = SheetSumber.Cells(i, PasswordSheetTujuanKolom).value ' Ambil kata sandi
        
        If namaLembarTujuan <> "" And selTujuan <> "" Then
            If WorksheetExists(namaLembarTujuan) Then
                Dim lembarTujuan As Worksheet
                Set lembarTujuan = ThisWorkbook.Sheets(namaLembarTujuan)
                
                If Not lembarTujuan Is Nothing Then
                    If passwordLembarTujuan <> "" Then
                        ' On Error Resume Next ' Hapus baris ini
                        lembarTujuan.Unprotect passwordLembarTujuan
                        ' On Error GoTo 0 ' Hapus baris ini
                        If lembarTujuan.ProtectContents Then
                            MsgBox "Kata sandi untuk " & lembarTujuan.Name & " salah!", vbExclamation
                            Exit Sub
                        End If
                    ElseIf lembarTujuan.ProtectContents Then
                        MsgBox "Lembar " & lembarTujuan.Name & " terlindungi. Masukkan kata sandi!", vbExclamation
                        Exit Sub
                    End If
                    
                    If RangeExists(lembarTujuan, selTujuan) Then
                        Application.DisplayAlerts = False
                        lembarTujuan.Range(selTujuan).value = nilaiRumus
                        Application.DisplayAlerts = True
                        
                        Dim tautan As Variant
                        tautan = ThisWorkbook.linkSources(xlExcelLinks)
                        
                        If Not IsEmpty(tautan) Then
                            Dim j As Long
                            For j = 1 To UBound(tautan)
                                ThisWorkbook.BreakLink Name:=tautan(j), Type:=xlLinkTypeExcelLinks
                            Next j
                        End If
                        
                        ' Lindungi lembar tujuan setelah mengisi nilai
                        If passwordLembarTujuan <> "" Then
                            'lembarTujuan.Protect passwordLembarTujuan
                        End If
                    Else
                        MsgBox "Kolom Sel Tujuan '" & selTujuan & "' tidak ditemukan di Lembar '" & lembarTujuan.Name & "'!", vbExclamation
                    End If
                End If
            Else
                MsgBox "Lembar Tujuan '" & namaLembarTujuan & "' tidak ditemukan!", vbExclamation
            End If
        End If
    Next i
End Sub

Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    ' On Error GoTo 0 ' Hapus baris ini
End Function

Function RangeExists(ws As Worksheet, rngAddress As String) As Boolean
    On Error Resume Next
    RangeExists = Not ws.Range(rngAddress) Is Nothing
    ' On Error GoTo 0 ' Hapus baris ini
End Function
