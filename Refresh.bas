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
    Author = Env.Author
    SheetNameData = Env.DataBase
    PathData = Env.Token
    Password = ""
    SearchValue = "20208081" ' ThisWorkbook.Sheets(SheetNameData).Range("B2").Value
    
    If SearchValue = "" Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    ' Menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(SheetNameData).Delete
    On Error GoTo 0
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
    
    ' Hanya menampilkan baris SearchValue
    Dim rng As Range
    Set rng = wsData.UsedRange
    rng.AutoFilter Field:=2, Criteria1:="<>" & SearchValue
    rng.Offset(1, 0).Resize(rng.Rows.count - 1, rng.Columns.count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    wsData.AutoFilterMode = False
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").value
    URLFOR = "https://data." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        On Error Resume Next
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("H1"))
            .Refresh BackgroundQuery:=False
        End With
        On Error GoTo 0
    End If

    ' Melindungi worksheet jika password diberikan
    If Password <> "" Then
        wsData.Protect Password
    End If

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    ' Pengaturan lainnya:
    ' Disini

    ' Menampilkan pesan setelah proses selesai
    Dim MessageUpdate As String
    MessageUpdate = wsData.Range("D2").value

    If MessageUpdate = "" Then
        MsgBox "Username tidak terdaftar!", vbExclamation
    Else
        MsgBox MessageUpdate, vbInformation, "Informasi"
    End If
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

Sub GsheetDataLogin()
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
    Author = Env.Author
    SheetNameData = Env.DataBase
    PathData = Env.Token
    SearchValue = HalamanLogin.TextBoxUsername.value
    
    If SearchValue = "" Or SearchValue = "Username" Then
        MsgBox "Masukkan Username terlebih dahulu!", vbExclamation
        Exit Sub
    End If
    
    ' Menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(SheetNameData).Delete
    On Error GoTo 0
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

    ' Hanya menampilkan baris SearchValue
    Dim rng As Range
    Set rng = wsData.UsedRange
    rng.AutoFilter Field:=2, Criteria1:="<>" & SearchValue
    rng.Offset(1, 0).Resize(rng.Rows.count - 1, rng.Columns.count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    wsData.AutoFilterMode = False
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").value
    URLFOR = "https://data." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        On Error Resume Next
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("H1"))
            .Refresh BackgroundQuery:=False
        End With
        On Error GoTo 0
    End If

    ' Melindungi worksheet jika password diberikan
    If Password <> "" Then
        wsData.Protect Password
    End If

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    ' Pengaturan lainnya:
    ' Disini

    ' Menampilkan pesan setelah proses selesai
    Dim MessageUpdate As String
    MessageUpdate = wsData.Range("D2").value

    If MessageUpdate = "" Then
        MsgBox "Username tidak terdaftar!", vbExclamation
    Else
        MsgBox MessageUpdate, vbInformation, "Informasi"
    End If
    Exit Sub

RefreshError:
    MsgBox UpdateErrorMsg, vbExclamation
End Sub
