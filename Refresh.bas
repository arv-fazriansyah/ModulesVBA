Sub GsheetData()
    Dim ws As Worksheet
    Dim sheetName As String, startCell As String
    Dim url As String, key As String, gid As String
    Dim password As String

    ' Validasi Input
    key = "14V7IxlKuEXi7275zO2gxK2I47h6IlIL2UU82FUSrBNM" ' Ganti dengan kunci Google Sheets yang valid
    gid = "0" ' Ganti dengan ID grup Google Sheets yang valid
    sheetName = "Sheet1" ' Ganti dengan nama worksheet yang Anda inginkan
    startCell = "A1" ' Ganti dengan sel awal yang Anda inginkan
    password = "ADMIN" ' Ganti dengan kata sandi perlindungan (jika diperlukan)

    ' Pesan-pesan MsgBox
    Dim internetErrorMsg As String
    internetErrorMsg = "Tidak ada koneksi internet. Pastikan Anda terhubung ke internet dan coba lagi."

    Dim wrongPasswordMsg As String
    wrongPasswordMsg = "Kata sandi yang dimasukkan salah. Data tidak dapat diperbarui."

    Dim updateCompleteMsg As String
    updateCompleteMsg = "Update selesai."

    Dim updateErrorMsg As String
    updateErrorMsg = "Terjadi kesalahan saat melakukan update data: "

    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox internetErrorMsg, vbExclamation
        Exit Sub
    End If

    ' Mengecek atau membuat worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Buat worksheet baru jika tidak ada
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    Else
        ' Unprotect worksheet dengan kata sandi jika diperlukan
        If password <> "" Then
            On Error Resume Next
            ws.Unprotect password
            On Error GoTo 0
            If ws.ProtectContents Then
                MsgBox wrongPasswordMsg, vbExclamation
                Exit Sub
            End If
        End If
    End If

    ' Hapus tabel kueri jika ada
    If ws.QueryTables.Count > 0 Then
        ws.QueryTables(1).Delete
    End If

    ' Hapus isi worksheet
    ws.Cells.Clear

    ' Buat URL untuk mengambil data dari Google Sheets
    url = "https://spreadsheets.google.com/tq?tqx=out:html&key=" & key & "&gid=" & gid & "&tx=tx"
    
    ' Set QueryTable dan mengambil data dari Google Sheets
    On Error GoTo RefreshError
    With ws.QueryTables.Add(Connection:="URL;" & url, Destination:=ws.Range(startCell))
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .BackgroundQuery = False
        .Refresh
    End With
    On Error GoTo 0

    ' Proteksi worksheet jika password diberikan
    If password <> "" Then
        ws.Protect password
    End If

    ' Hapus semua koneksi data dalam workbook
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    ' Tampilkan pesan ketika proses selesai
    MsgBox updateCompleteMsg, vbInformation
    Exit Sub

RefreshError:
    MsgBox updateErrorMsg & Err.Description, vbExclamation
End Sub

Function IsInternetConnected() As Boolean
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    xhr.Open "GET", "https://www.google.com", False
    xhr.send
    If Err.Number <> 0 Then
        IsInternetConnected = False
    Else
        IsInternetConnected = (xhr.Status = 200)
    End If
    On Error GoTo 0
End Function
