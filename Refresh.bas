Sub GsheetDataUpdate()
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String
    Dim password As String, Author As String
    Dim SearchValue As String
    Dim internetErrorMsg As String, updateErrorMsg As String
    
    ' Pesan kesalahan
    internetErrorMsg = "Tidak ada koneksi internet."
    updateErrorMsg = "Download ulang Aplikasi, hubungi Admin"
    
    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox internetErrorMsg, vbExclamation
        Exit Sub
    End If
    
    On Error GoTo RefreshError

    ' Konfigurasi data
    Author = Env.Author
    SheetNameData = Env.DataBase
    PathData = Env.Token
    SearchValue = "20208081" ' ThisWorkbook.Sheets(SheetNameData).Range("B2").value
    
    If SearchValue = "" Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    '"20208081"
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
        .WebSelectionType = xlAllTables ' Memilih semua tabel dari halaman web
        .WebFormatting = xlWebFormattingNone ' Tidak melakukan pemformatan web
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
            .WebSelectionType = xlAllTables ' Memilih semua tabel dari halaman web
            .WebFormatting = xlWebFormattingNone ' Tidak melakukan pemformatan web
            .Refresh BackgroundQuery:=False
        End With
    End If

    ' Melindungi worksheet jika password diberikan
    If password <> "" Then wsData.Protect password

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    ' Menampilkan pesan setelah proses selesai
    Dim MessageUpdate As String
    MessageUpdate = wsData.Range("D2").value

    If MessageUpdate = "" Then
        MsgBox "Username tidak terdaftar!", vbExclamation
        Exit Sub
    Else
        MsgBox MessageUpdate, vbInformation, "Informasi"
    End If
Exit Sub

RefreshError:
    MsgBox updateErrorMsg, vbExclamation
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
    Dim password As String, Author As String
    Dim SearchValue As String
    Dim internetErrorMsg As String, updateErrorMsg As String
    
    ' Pesan kesalahan
    internetErrorMsg = "Tidak ada koneksi internet."
    updateErrorMsg = "Download ulang Aplikasi, hubungi Admin"
    
    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox internetErrorMsg, vbExclamation
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
    
    '"20206687"
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
    If password <> "" Then wsData.Protect password

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    If MessageUpdate = "" Then
        MsgBox "Username tidak terdaftar!", vbExclamation
        Exit Sub
    Else
        MsgBox MessageUpdate, vbInformation, "Informasi"
    End If
Exit Sub

RefreshError:
    MsgBox updateErrorMsg, vbExclamation
End Sub
