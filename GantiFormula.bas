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
    ' On Error GoTo 0 ' Hapus baris ini
    
    If sheetSumber Is Nothing Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    Dim pemisah As String
    pemisah = Application.International(xlListSeparator)
    
    Dim barisTerakhir As Long
    barisTerakhir = sheetSumber.Cells(sheetSumber.Rows.Count, rumusKolomSumber).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To barisTerakhir
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
