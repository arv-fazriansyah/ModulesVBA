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

    ' Mengambil alamat IP publik dari layanan pihak ketiga
    HTTPReq.Open "GET", "https://api.ipify.org", False
    HTTPReq.Send

    ' Menyimpan alamat IP publik ke dalam variabel
    ipAddress = HTTPReq.responseText

    ' Menuliskan alamat IP publik ke sel A1 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F3").Value = ipAddress
End Sub

Sub CekIDPerangkat()
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim deviceId As String
    
    ' Membuat objek untuk mengakses layanan Windows Management Instrumentation (WMI)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct", , 48)
    
    ' Mendapatkan ID perangkat
    For Each objItem In colItems
        deviceId = objItem.IdentifyingNumber
        Exit For ' Hanya mengambil ID dari perangkat pertama yang ditemukan
    Next objItem
    
    ' Menyimpan ID perangkat ke dalam sel F4 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F4").Value = deviceId
End Sub

Sub CekWaktu()
    ' Mendapatkan waktu saat ini
    Dim currentTime As String
    currentTime = Format(Now, "dddd, dd-mm-yyyy hh:mm:ss")
    
    ' Menyimpan waktu saat ini ke dalam sel F5 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F5").Value = currentTime
End Sub

Sub CekVersiOffice()
    Dim officeVersion As String
    
    ' Mendapatkan informasi versi Microsoft Office
    officeVersion = Application.Version
    
    ' Menyimpan informasi versi Office ke dalam sel F6 di sheet "DEV"
    ThisWorkbook.Sheets("DEV").Range("F6").Value = officeVersion
End Sub

Sub CekNamaKomputer()
    ' Mendapatkan nama komputer
    Dim computerName As String
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
    
    wrongPasswordMsg = "Silahkan Update Aplikasi!"
    
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ' Periksa apakah lembar kerja ada
    If ws Is Nothing Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    ' Jika lembar kerja dilindungi, lakukan unprotect terlebih dahulu
    If ws.ProtectContents Then
        If Pass <> "" Then
            On Error Resume Next
            ws.Unprotect Pass
            If ws.ProtectContents Then
                MsgBox wrongPasswordMsg, vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    ' Tentukan rentang sel yang ingin dihapus
    On Error GoTo Error
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
    MsgBox wrongPasswordMsg, vbExclamation
End Sub
