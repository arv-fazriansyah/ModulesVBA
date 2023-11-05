Sub GsheetData()
    Dim ws As Worksheet
    Dim sheetName As String, startCell As String
    Dim url As String, key As String, gid As String, user As String, password As String

    ' Konfigurasi Google Sheets
    key = "14V7IxlKuEXi7275zO2gxK2I47h6IlIL2UU82FUSrBNM"
    gid = "0"
    user = "20206687"
    
    ' Konfigurasi worksheet
    sheetName = "Sheet1"
    startCell = "A1"
    password = "ADMIN"

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

    ' Membuat worksheet jika tidak ada
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    ' Unprotect worksheet dengan kata sandi jika diperlukan
    ElseIf password <> "" Then
        On Error Resume Next
        ws.Unprotect password
        On Error GoTo 0
        If ws.ProtectContents Then
            MsgBox wrongPasswordMsg, vbExclamation
            Exit Sub
        End If
    End If

    ' Menghapus tabel kueri jika ada
    If ws.QueryTables.Count > 0 Then
        ws.QueryTables(1).Delete
    End If

    ' Menghapus isi worksheet
    ws.Cells.Clear

    ' Membuat URL untuk mengambil data dari Google Sheets
    url = "https://docs.google.com/spreadsheets/u/0/d/" & key & "/gviz/tq?tqx=out:html&gid=" & gid & "&tq=SELECT+*+WHERE+B%3D" & user

    On Error GoTo RefreshError
    ' Menyiapkan QueryTable dan mengambil data dari Google Sheets
    With ws.QueryTables.Add(Connection:="URL;" & url, Destination:=ws.Range(startCell))
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .RefreshStyle = xlInsertDeleteCells
        .HasAutoFormat = True
        .TablesOnlyFromHTML = False
        .SaveData = True
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With
    On Error GoTo 0

    ' Melindungi worksheet jika password diberikan
    If password <> "" Then ws.Protect password

    ' Menghapus semua koneksi data dalam workbook
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
    MsgBox "Hay", vbInformation, "Informasi"
End Sub
