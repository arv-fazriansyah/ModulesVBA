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
    
    If SheetSumber Is Nothing Then
        Exit Sub
    End If
    
    Dim pemisah As String
    pemisah = Application.International(xlListSeparator)
    
    Dim barisTerakhir As Long
    barisTerakhir = SheetSumber.Cells(SheetSumber.Rows.Count, RumusKolomSumber).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To barisTerakhir
        Dim nilaiRumus As String
        nilaiRumus = SheetSumber.Cells(i, RumusKolomSumber).Formula2 ' Menggunakan .Formula2
        nilaiRumus = Replace(nilaiRumus, ";", pemisah)
        nilaiRumus = Replace(nilaiRumus, ",", pemisah)
        
        Dim namaLembarTujuan As String
        namaLembarTujuan = SheetSumber.Cells(i, NamaSheetTujuanKolom).Value
        
        Dim selTujuan As String
        selTujuan = SheetSumber.Cells(i, SelTujuanKolom).Value
        
        Dim passwordLembarTujuan As String
        passwordLembarTujuan = SheetSumber.Cells(i, PasswordSheetTujuanKolom).Value
        
        If namaLembarTujuan <> "" And selTujuan <> "" Then
            If WorksheetExists(namaLembarTujuan) Then
                Dim lembarTujuan As Worksheet
                Set lembarTujuan = ThisWorkbook.Sheets(namaLembarTujuan)
                
                If Not lembarTujuan Is Nothing Then
                    If passwordLembarTujuan <> "" Then
                        lembarTujuan.Unprotect passwordLembarTujuan
                        If lembarTujuan.ProtectContents Then
                            Exit Sub
                        End If
                    ElseIf lembarTujuan.ProtectContents Then
                        Exit Sub
                    End If
                    
                    If RangeExists(lembarTujuan, selTujuan) Then
                        Application.DisplayAlerts = False
                        lembarTujuan.Range(selTujuan).Formula2 = nilaiRumus ' Menggunakan .Formula2
                        Application.DisplayAlerts = True
                        
                        Dim tautan As Variant
                        tautan = ThisWorkbook.LinkSources(xlExcelLinks)
                        
                        If Not IsEmpty(tautan) Then
                            Dim j As Long
                            For j = 1 To UBound(tautan)
                                ThisWorkbook.BreakLink Name:=tautan(j), Type:=xlLinkTypeExcelLinks
                            Next j
                        End If
                        
                        If passwordLembarTujuan <> "" Then
                            lembarTujuan.Protect passwordLembarTujuan, UserInterfaceOnly:=True
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

Function WorksheetExists(SheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(SheetName) Is Nothing
End Function

Function RangeExists(ws As Worksheet, rngAddress As String) As Boolean
    On Error Resume Next
    RangeExists = Not ws.Range(rngAddress) Is Nothing
End Function
