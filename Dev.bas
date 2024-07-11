Function TesKoneksi() As Boolean
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xhr.Open "GET", "https://www.google.com", False
    xhr.Send
    TesKoneksi = (Err.Number = 0) And (xhr.Status = 200)
    ' On Error GoTo 0
End Function

Sub ShowForm()
    Application.Visible = False ' Sembunyikan jendela Excel

    ' Tampilkan UserForm login sebagai modal
    HalamanLogin.Show vbModal
End Sub

Sub HideForm()
    Application.Visible = True ' Tampilkan jendela Excel
    HalamanLogin.Hide ' Sembunyikan UserForm login
End Sub

Sub ClosedForm()
    If CloseMode = vbFormControlMenu Then
        ' Nonaktifkan peringatan
        Application.DisplayAlerts = False
        ' Tutup Excel
        Application.Quit
    End If
End Sub

Sub ClosedAllExcel()
    ' Kode ini akan dijalankan saat file Excel dibuka

    ' Nonaktifkan peringatan
    Application.DisplayAlerts = False

    ' Menutup semua workbook kecuali workbook ini
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close SaveChanges:=True
        End If
    Next wb

    ' Aktifkan kembali peringatan
    Application.DisplayAlerts = True
End Sub

Sub CekIPPublik()
    Dim HTTPReq As Object
    Dim ipAddress As String
    Set HTTPReq = CreateObject("MSXML2.XMLHTTP")
    
    On Error Resume Next
    ' Mengambil alamat IP publik dari layanan pihak ketiga
    HTTPReq.Open "GET", "https://api.ipify.org", False
    HTTPReq.Send

    ' Menyimpan alamat IP publik ke dalam variabel
    ipAddress = HTTPReq.responseText

    ' Menuliskan alamat IP publik ke sel A1 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F4").Value = ipAddress
End Sub

Sub CekIDPerangkat()
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim deviceId As String
    
    On Error Resume Next
    ' Membuat objek untuk mengakses layanan Windows Management Instrumentation (WMI)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct", , 48)
    
    ' Mendapatkan ID perangkat
    For Each objItem In colItems
        deviceId = objItem.IdentifyingNumber
        Exit For ' Hanya mengambil ID dari perangkat pertama yang ditemukan
    Next objItem
    
    ' Menyimpan ID perangkat ke dalam sel F4 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F5").Value = deviceId
End Sub

Sub CekWaktu()
    ' Mendapatkan waktu saat ini
    Dim currentTime As String
    Dim dayName As String
    Dim translatedDate As String
    
    On Error Resume Next
    
    ' Mendapatkan tanggal dan waktu dalam format standar
    currentTime = Format(Now, "dddd, dd-mm-yyyy hh:mm:ss")
    
    ' Mengubah nama hari ke bahasa Indonesia
    dayName = Format(Now, "dddd")
    Select Case dayName
        Case "Sunday"
            dayName = "Minggu"
        Case "Monday"
            dayName = "Senin"
        Case "Tuesday"
            dayName = "Selasa"
        Case "Wednesday"
            dayName = "Rabu"
        Case "Thursday"
            dayName = "Kamis"
        Case "Friday"
            dayName = "Jumat"
        Case "Saturday"
            dayName = "Sabtu"
    End Select
    
    ' Membuat timestamp dengan nama hari dalam bahasa Indonesia
    translatedDate = dayName & ", " & Format(Now, "dd-mm-yyyy hh:mm:ss")
    
    ' Menyimpan waktu saat ini ke dalam sel F3 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F3").Value = translatedDate
End Sub

Sub CekVersiOffice()
    Dim officeVersion As String
    
    On Error Resume Next
    ' Mendapatkan informasi versi Microsoft Office
    officeVersion = Application.Version
    
    ' Menyimpan informasi versi Office ke dalam sel F6 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F6").Value = officeVersion
End Sub

Sub CekNamaKomputer()
    ' Mendapatkan nama komputer
    Dim computerName As String
    
    On Error Resume Next
    computerName = Environ("COMPUTERNAME")
    
    ' Menyimpan nama komputer ke dalam sel F7 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F7").Value = computerName
End Sub

Sub HapusData()
    Dim ws As Worksheet
    Dim SheetName As String
    Dim dataRange As String
    Dim Pass As String
    Dim wrongPasswordMsg As String
    
    On Error Resume Next
    
    ' Konfigurasi
    SheetName = Env.DataBase
    dataRange = ThisWorkbook.Sheets(SheetName).Range("H2")
    Pass = ThisWorkbook.Sheets(SheetName).Range("G2")
    
    ' wrongPasswordMsg = "Silahkan Update Aplikasi Anda!"
    
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ' Periksa apakah lembar kerja ada
    If ws Is Nothing Then
        ' MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    ' Jika lembar kerja dilindungi, lakukan unprotect terlebih dahulu
    If ws.ProtectContents Then
        If Pass <> "" Then
            On Error Resume Next
            ws.Unprotect Pass
            If ws.ProtectContents Then
                ' MsgBox wrongPasswordMsg, vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    ' Tentukan rentang sel yang ingin dihapus
    Dim DeleteRange As Range
    Set DeleteRange = ws.Range(dataRange)
    
    ' Hapus data dari rentang sel yang ditentukan
    If Not DeleteRange Is Nothing Then
        DeleteRange.ClearContents ' Menghapus isi dari sel
    End If
    
    ' Setelah pembaruan, lindungi kembali lembar kerja jika password diberikan
    If Pass <> "" Then ws.Protect Pass
    
    ' Simpan workbook setelah menghapus data eksternal
    ThisWorkbook.Save
    Exit Sub
    
Error:
    ' MsgBox wrongPasswordMsg, vbExclamation
End Sub

Sub CopyFormulas()
    Dim NamaSheetSumber As String
    Dim RumusKolomSumber As String
    Dim NamaSheetTujuanKolom As String
    Dim SelTujuanKolom As String
    Dim PasswordSheetTujuanKolom As String
    Dim SheetSumber As Worksheet
    
    NamaSheetSumber = Env.DataBase
    RumusKolomSumber = "AA"
    NamaSheetTujuanKolom = "AB"
    SelTujuanKolom = "AC"
    PasswordSheetTujuanKolom = "AD"
    
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
        
        ' Ubah "#=" menjadi "=" di awal kalimat
        If Left(nilaiRumus, 2) = "#=" Then
            nilaiRumus = "=" & Mid(nilaiRumus, 3)
        End If
        
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
