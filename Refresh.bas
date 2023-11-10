Sub GsheetData()
    Dim ws As Worksheet
    Dim SheetName As String, URL As String, Path As String, Password As String, Author As String
    Dim InternetErrorMsg As String, UpdateErrorMsg As String
    Dim searchValue As String

    ' Konfigurasi
    Author = "fazriansyah"
    Path = "token"
    Password = ""
    SheetName = "DATAUSER"
    searchValue = HalamanLogin.TextBoxUsername.value

    ' Pesan Kesalahan
    InternetErrorMsg = "Tidak ada koneksi internet."
    UpdateErrorMsg = "Download ulang Aplikasi, hubungi Admin"

    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox InternetErrorMsg, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' Hapus lembar jika sudah ada
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    ' Membuat worksheet baru
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = SheetName

    ' Membuat URL untuk mengambil data dari Google Sheets
    URL = "https://data." & Author & ".eu.org/" & Path

    On Error GoTo RefreshError
    ' Menyiapkan QueryTable dan mengambil data dari Google Sheets
    With ws.QueryTables.Add(Connection:="URL;" & URL, Destination:=ws.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With

    ' Hanya menampilkan baris
    Dim i As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    Application.ScreenUpdating = False

    For i = lastRow To 2 Step -1 ' Dimulai dari baris kedua
        If ws.Cells(i, 1).value <> searchValue Then
            ws.Rows(i).Delete
        End If
    Next i

    Application.ScreenUpdating = True

    ' Melindungi worksheet jika password diberikan
    If Password <> "" Then ws.Protect Password

    ' Menghapus semua koneksi data dalam workbook
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    ' Menampilkan pesan setelah proses selesai
    ShowRefreshMessage
    Exit Sub

RefreshError:
    MsgBox UpdateErrorMsg, vbExclamation
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
    Dim MessageUpdate As String
    MessageUpdate = ThisWorkbook.Sheets("DATAUSER").Range("B2").value
    MsgBox MessageUpdate, vbInformation, "Informasi"
End Sub
