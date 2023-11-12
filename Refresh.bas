Sub GsheetDataUpdate()
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String
    Dim Password As String, Author As String
    Dim SearchValue As String
    Dim InternetErrorMsg As String, UpdateErrorMsg As String
    
    ' Pesan kesalahan
    InternetErrorMsg = "Tidak ada koneksi internet."
    UpdateErrorMsg = "Download ulang Aplikasi, hubungi Admin"
    
    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox InternetErrorMsg, vbExclamation
        Exit Sub
    End If
    
    On Error GoTo RefreshError

    ' Konfigurasi data
    Author = "fazriansyah"
    SheetNameData = "DATAUSER"
    PathData = "AKfycbxeCKT0HoO-SMxBbCv_mA3g5fMkI1Ke6119G8KfDWdlVn7zf3boXtAJ3qadMHvlFpscsg"
    SearchValue = "20206687"
    'HalamanLogin.TextBoxUsername.Value
    'ThisWorkbook.Sheets(SheetNameData).Range("B2").value

    ' Menghapus lembar kerja yang sudah ada jika ada
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetNameData).Delete
    Application.DisplayAlerts = True

    ' Membuat lembar kerja baru
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = SheetNameData
    
    ' Membuat URL untuk mengambil data
    Dim URLDAT As String, URLFOR As String
    URLDAT = "https://data." & Author & ".eu.org/" & PathData
    
    ' Menyiapkan QueryTable dan mengambil data
    With wsData.QueryTables.Add(Connection:="URL;" & URLDAT, Destination:=wsData.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With

    ' Hanya menampilkan baris
    Dim i As Long
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.count, 1).End(xlUp).row
    Application.ScreenUpdating = False
    For i = lastRow To 2 Step -1 ' Dimulai dari baris kedua
        If wsData.Cells(i, 2).value <> SearchValue Then
            wsData.Rows(i).Delete
        End If
    Next i
    Application.ScreenUpdating = True
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").value
    URLFOR = "https://data." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        On Error Resume Next
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
