Sub GsheetData()
    Dim ws As Worksheet
    Dim sheetName As String, startCell As String
    Dim URL As String, Path As String, Password As String, Author As String

    ' Konfigurasi
    Author = "fazriansyah"
    Path = "token"
    Password = ""

    ' Konfigurasi worksheet
    sheetName = "Sheet1"
    startCell = "A1"

    ' Pesan Kesalahan
    Dim internetErrorMsg As String
    Dim wrongPasswordMsg As String
    Dim updateErrorMsg As String

    internetErrorMsg = "Tidak ada koneksi internet. Pastikan Anda terhubung ke internet dan coba lagi."
    wrongPasswordMsg = "Kata sandi yang dimasukkan salah. Data tidak dapat diperbarui."
    updateErrorMsg = "Terjadi kesalahan saat melakukan update data: "

    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox internetErrorMsg, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' Membuat worksheet jika tidak ada atau menghapus lembar yang dilindungi tanpa password
    If ws Is Nothing Or (ws.ProtectContents And Password = "") Then
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    ElseIf Password <> "" Then
        On Error Resume Next
        ws.Unprotect Password
        On Error GoTo 0
        If ws.ProtectContents Then
            MsgBox wrongPasswordMsg, vbExclamation
            Exit Sub
        End If
    End If

    ' Menghapus tabel kueri jika ada
    If ws.QueryTables.count > 0 Then
        ws.QueryTables(1).Delete
    End If

    ' Menghapus isi worksheet
    ws.Cells.Clear

    ' Membuat URL untuk mengambil data dari Google Sheets
    URL = "https://data." & Author & ".eu.org/" & Path

    On Error GoTo RefreshError
    ' Menyiapkan QueryTable dan mengambil data dari Google Sheets
    With ws.QueryTables.Add(Connection:="URL;" & URL, Destination:=ws.Range(startCell))
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .RefreshStyle = xlInsertDeleteCells
        .HasAutoFormat = True
        .TablesOnlyFromHTML = False
        .SaveData = True
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With

    ' Melindungi worksheet jika password diberikan
    If Password <> "" Then ws.Protect Password

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    ' Menampilkan pesan setelah proses selesai
    ShowRefreshMessage
    Exit Sub

RefreshError:
    MsgBox updateErrorMsg & Err.Description, vbExclamation
End Sub

Function IsInternetConnected() As Boolean
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xhr.Open "GET", "https://www.google.com", False
    xhr.send
    IsInternetConnected = (Err.Number = 0) And (xhr.Status = 200)
    On Error GoTo 0
End Function

Sub ShowRefreshMessage()
    MsgBox "Data telah berhasil diperbarui.", vbInformation, "Informasi"
End Sub
