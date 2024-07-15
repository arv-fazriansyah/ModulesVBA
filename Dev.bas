Function TesKoneksi() As Boolean
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Set timeout values (in milliseconds)
    Dim timeout As Long
    timeout = 1000 ' 1 seconds timeout
    
    ' Open and send the request
    xhr.setTimeouts timeout, timeout, timeout, timeout
    xhr.Open "GET", "https://www.google.com", False
    xhr.Send
    
    ' Check if the request was successful
    TesKoneksi = (Err.Number = 0) And (xhr.Status = 200)
End Function

Public Sub Expired()
    Dim ws As Worksheet
    Dim TanggalExpired As Variant
    
    On Error Resume Next

    ' Menentukan sheet DATAUSER
    Set ws = ThisWorkbook.Sheets("DATAUSER")
    
    ' Mendapatkan nilai dari cell H2
    TanggalExpired = ws.Range("H2").Value
    
    ' Menghilangkan tanda # jika ada
    TanggalExpired = Replace(TanggalExpired, "#", "")

    ' Memeriksa apakah TanggalExpired adalah tanggal yang valid
    If IsDate(TanggalExpired) Then
        TanggalExpired = CDate(TanggalExpired)
    Else
        'MsgBox "Tanggal tidak valid atau kosong di H2", vbExclamation
        Exit Sub
    End If

    ' Memeriksa apakah masa trial sudah habis
    If TanggalExpired <= Date Then
        MsgBox "Masa Trial sudah habis!" & vbNewLine & _
               "File ini akan terhapus otomatis", vbInformation
        With ThisWorkbook
            .ChangeFileAccess xlReadOnly
            Kill .FullName
            .Close False
        End With
    Else
        'MsgBox "Masa Trial tinggal " & TanggalExpired - Date & " Hari lagi", vbInformation
    End If
End Sub

Sub AutoHideColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DATAUSER")
    
    ' Menyembunyikan kolom AA sampai ZZ
    ws.columns("AA:ZZ").Hidden = True
End Sub

Public Sub TextVersion()
    Dim dataText As String

    On Error Resume Next
    dataText = ThisWorkbook.Sheets("DEV").Range("J3").Value

    If dataText = "" Then
        dataText = ThisWorkbook.Sheets("DATAUSER").Range("E2").Value
        
        If dataText = "" Then
            dataText = "Masukan Username, kemudian Update!!"
        End If
    End If

    With HalamanLogin.LabelVersion
        .Tag = dataText
        .Caption = dataText
        .ForeColor = RGB(255, 255, 255) ' Atur warna teks putih
    End With
End Sub

Sub TimeButton()
    Static LastRunTime As Date
    Dim TimeDelay As Date
    
    TimeDelay = Now() - LastRunTime
    If TimeDelay < TimeValue("00:02:00") Then
        MsgBox "Maaf, Anda hanya dapat menjalankan fungsi ini sekali dalam dua menit.", vbExclamation
        Exit Sub
    End If
    LastRunTime = Now()
End Sub
    
Sub ShowForm()
    Application.ScreenUpdating = False
    If Workbooks.Count > 1 Then
        Windows(ThisWorkbook.Name).Visible = False
    Else
        Application.Visible = False
    End If
    ' Tampilkan UserForm login sebagai modal
    HalamanLogin.Show vbModeless 'vbModal
End Sub

Sub HideForm()
    HalamanLogin.Hide ' Sembunyikan UserForm login
    'Application.ScreenUpdating = True
    Application.Visible = True ' Tampilkan jendela Excel
    Windows(ThisWorkbook.Name).Visible = True
    Berhenti = True
End Sub

Sub ClosedForm()
    If CloseMode = vbFormControlMenu Then
        ' Nonaktifkan peringatan
        Application.DisplayAlerts = False
        ' Tutup Excel
        'Application.Quit
        ThisWorkbook.Close SaveChanges:=False
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

Function CekWaktu(tanggal As Date) As String
    Dim formattedDate As String
    
    ' Menggunakan fungsi Text untuk memformat tanggal dalam bahasa Indonesia
    formattedDate = Application.WorksheetFunction.text(tanggal, "[$-0421]DDDD, DD MMMM YYYY hh:mm:ss")
    
    CekWaktu = formattedDate
End Function

Sub SimpanWaktu()
    Dim waktuSekarang As String
    waktuSekarang = CekWaktu(Now)
    
    On Error Resume Next
    ' Menyimpan waktu saat ini ke dalam sel F3 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F3").Value = waktuSekarang
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
    Dim sheetName As String
    Dim dataRange As String
    Dim Pass As String
    Dim wrongPasswordMsg As String
    
    On Error Resume Next
    
    ' Konfigurasi
    sheetName = Env.DataBase
    dataRange = ThisWorkbook.Sheets(sheetName).Range("H2")
    Pass = ThisWorkbook.Sheets(sheetName).Range("G2")
    
    ' wrongPasswordMsg = "Silahkan Update Aplikasi Anda!"
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    
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
    Dim deleteRange As Range
    Set deleteRange = ws.Range(dataRange)
    
    ' Hapus data dari rentang sel yang ditentukan
    If Not deleteRange Is Nothing Then
        deleteRange.ClearContents ' Menghapus isi dari sel
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
    barisTerakhir = SheetSumber.Cells(SheetSumber.Rows.Count, RumusKolomSumber).End(xlUp).row
    
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

Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
End Function

Function RangeExists(ws As Worksheet, rngAddress As String) As Boolean
    On Error Resume Next
    RangeExists = Not ws.Range(rngAddress) Is Nothing
End Function

Sub ProtectSheets()
    Dim DataSheet As Worksheet
    
    On Error Resume Next
    
    Set DataSheet = ThisWorkbook.Sheets("DATAUSER")

    Dim lastRowData As Long
    lastRowData = DataSheet.Cells(DataSheet.Rows.Count, "AF").End(xlUp).row

    Dim i As Long
    Dim actionType As String
    Dim sheetNameValue As String
    Dim sheetPassword As String
    Dim TargetSheet As Worksheet

    For i = 2 To lastRowData
        actionType = DataSheet.Cells(i, "AF").Value
        sheetNameValue = DataSheet.Cells(i, "AG").Value
        sheetPassword = DataSheet.Cells(i, "AH").Value

        On Error Resume Next
        Set TargetSheet = ThisWorkbook.Sheets(sheetNameValue)
        On Error GoTo 0

        If Not TargetSheet Is Nothing Then
            If actionType = 1 Then
                ProtectSheet TargetSheet, sheetPassword
            ElseIf actionType = 0 Then
                UnprotectSheet TargetSheet, sheetPassword
            End If
        Else
            ' MsgBox "Lembar '" & sheetNameValue & "' tidak ditemukan.", vbExclamation
        End If
    Next i
End Sub

Sub ProtectSheet(sheet As Worksheet, Password As String)
    sheet.Unprotect Password
    sheet.Protect Password
End Sub

Sub UnprotectSheet(sheet As Worksheet, Password As String)
    sheet.Unprotect Password
End Sub

Sub SendData()
    Dim url As String
    Dim HTTPReq As Object
    Dim JSONString As String
    Dim RangeData As Range
    Dim DataArray As Variant
    Dim i As Long
    Dim j As Long
    Dim sheetName As String
    Dim StartColumn As String
    Dim EndColumn As String

    On Error Resume Next

    ' URL for Google Sheets REST API
    url = "https://" & SubPath & "." & Author & ".eu.org/" & "send"

    ' Set the sheet name and data range
    sheetName = "DEV"
    StartColumn = "AA"
    EndColumn = "ZZ"

    ' Set the data range from Excel
    With ThisWorkbook.Sheets(sheetName)
        Set RangeData = .Range(StartColumn & "3:" & EndColumn & .Cells(.Rows.Count, StartColumn).End(xlUp).row)
    End With

    ' Convert the data range to an array
    DataArray = RangeData.Value

    ' Create the JSON string
    JSONString = "{""values"": ["
    For i = 1 To UBound(DataArray)
        JSONString = JSONString & "["
        For j = 1 To UBound(DataArray, 2)
            JSONString = JSONString & """" & DataArray(i, j) & """"
            If j <> UBound(DataArray, 2) Then
                JSONString = JSONString & ","
            End If
        Next j
        JSONString = JSONString & "]"
        If i <> UBound(DataArray) Then
            JSONString = JSONString & ","
        End If
    Next i
    JSONString = JSONString & "]}"

    ' Show the JSON string in a message box
    'MsgBox JSONString, vbInformation, "JSON Data"

    ' Create the WinHttpRequest object
    Set HTTPReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Send the POST request
    HTTPReq.Open "POST", url, False
    HTTPReq.setRequestHeader "Content-Type", "application/json"
    HTTPReq.Send JSONString

    ' Show the result message
    ' MsgBox "Data has been successfully sent to Google Sheets.", vbInformation
End Sub

