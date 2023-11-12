Sub GsheetData()
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String
    Dim Password As String, Author As String
    Dim searchValue As String
    Dim InternetErrorMsg As String, UpdateErrorMsg As String

    ' Konfigurasi
    Author = "fazriansyah"
    SheetNameData = "DATAUSER"
    PathData = "AKfycbxeCKT0HoO-SMxBbCv_mA3g5fMkI1Ke6119G8KfDWdlVn7zf3boXtAJ3qadMHvlFpscsg"
    searchValue = "20206687" ' HalamanLogin.TextBoxUsername.Value

    ' Pesan Kesalahan
    InternetErrorMsg = "Tidak ada koneksi internet."
    UpdateErrorMsg = "Download ulang Aplikasi, hubungi Admin"

    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox InternetErrorMsg, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ' Coba menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False ' Matikan peringatan penghapusan lembar kerja
    ThisWorkbook.Sheets(SheetNameData).Delete
    Application.DisplayAlerts = True ' Hidupkan peringatan penghapusan lembar kerja
    On Error GoTo 0

    ' Membuat lembar kerja baru
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = SheetNameData
    
    ' Membuat URL untuk mengambil data dari Google Sheets
    Dim URLDAT As String, URLFOR As String
    URLDAT = "https://data." & Author & ".eu.org/" & PathData
    
    On Error GoTo RefreshError
    ' Menyiapkan QueryTable dan mengambil data dari Google Sheets - Data
    With wsData.QueryTables.Add(Connection:="URL;" & URLDAT, Destination:=wsData.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With

    ' Hanya menampilkan baris
    Dim i As Long
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.count, 1).End(xlUp).row

    Application.ScreenUpdating = False

    For i = lastRow To 2 Step -1 ' Dimulai dari baris kedua
        If wsData.Cells(i, 2).value <> searchValue Then
            wsData.Rows(i).Delete
        End If
    Next i

    Application.ScreenUpdating = True
    
    PathFormula = wsData.Range("F2").value
    URLFOR = "https://data." & Author & ".eu.org/" & PathFormula
    
    On Error GoTo RefreshError
    If PathFormula <> "" Then
        ' Menyiapkan QueryTable dan mengambil data dari Google Sheets - Data
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("H1"))
            .Refresh BackgroundQuery:=False
        End With
    End If

    ' Melindungi worksheet jika password diberikan
    If Password <> "" Then wsData.Protect Password

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    ' Menampilkan pesan setelah proses selesai
    Dim MessageUpdate As String
        MessageUpdate = wsData.Range("D2").value
        MsgBox MessageUpdate, vbInformation, "Informasi"
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
