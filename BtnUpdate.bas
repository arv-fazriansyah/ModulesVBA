Sub GsheetDataUpdateDownload()
    ' Tombol Download
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String, SubPath As String
    Dim Password As String, Author As String
    Dim SearchValue As String
    Dim UpdateErrorMsg As String
    
    On Error Resume Next

    ' Konfigurasi data
    Author = Env.Author
    SubPath = Env.SubPath
    PathData = Env.Token
    SheetNameData = Env.DataBase
    SearchValue = ThisWorkbook.Sheets(SheetNameData).Range("A2").Value
    
    ' Menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetHidden
    ThisWorkbook.Sheets(SheetNameData).Delete
    Application.DisplayAlerts = True

    ' Membuat lembar kerja baru
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = SheetNameData
    
    ' Hidden sheet DATAUSER
    'ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetVeryHidden
    
    ' Membuat URL untuk mengambil data
    Dim URLDAT As String, URLFOR As String
    URLDAT = "https://" & SubPath & "." & Author & ".eu.org/" & PathData
    
    ' Menyiapkan QueryTable dan mengambil data
    With wsData.QueryTables.Add(Connection:="URL;" & URLDAT, Destination:=wsData.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With
    
    ' Hanya menampilkan baris SearchValue
    Dim rng As Range
    Set rng = wsData.UsedRange
    rng.AutoFilter Field:=1, Criteria1:="<>" & SearchValue
    If rng.columns(1).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
        rng.offset(1, 0).Resize(rng.Rows.Count - 1, rng.columns.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    wsData.AutoFilterMode = False
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").Value
    URLFOR = "https://" & SubPath & "." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("AA1"))
            .Refresh BackgroundQuery:=False
        End With
    End If

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    ' Pengaturan lainnya:
    Dev.HapusData
    Dev.AutoHideColumns
    
    ' Melindungi worksheet jika password diberikan
    Password = wsData.Range("G2").Value
    If Password <> "" Then
        wsData.Protect Password
    End If
    Exit Sub
End Sub

Sub GsheetDataUpdate()
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String, SubPath As String
    Dim Password As String, Author As String
    Dim SearchValue As String
    Dim UpdateErrorMsg As String
    Static LastRunTime As Date
    Dim TimeDelay As Date
    
    ' Pesan kesalahan
    UpdateErrorMsg = "Download ulang Aplikasi, hubungi Admin"
    
    ' Mengecek koneksi internet
    If Not Dev.TesKoneksi Then
        MsgBox "Tidak ada koneksi internet.", vbExclamation
        Exit Sub
    End If
    
    TimeDelay = Now() - LastRunTime
    If TimeDelay < TimeValue("00:05:00") Then
        MsgBox "Maaf, Anda hanya dapat menjalankan fungsi ini sekali dalam 5 menit.", vbExclamation
        Exit Sub
    End If
    LastRunTime = Now()
    
    On Error Resume Next

    ' Konfigurasi data
    Author = Env.Author
    SubPath = Env.SubPath
    PathData = Env.Token
    SheetNameData = Env.DataBase
    SearchValue = ThisWorkbook.Sheets(SheetNameData).Range("A2").Value
    
    If SearchValue = "" Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    ' Menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetHidden
    ThisWorkbook.Sheets(SheetNameData).Delete
    Application.DisplayAlerts = True

    ' Membuat lembar kerja baru
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = SheetNameData
    
    ' Hidden sheet DATAUSER
    'ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetVeryHidden
    
    ' Membuat URL untuk mengambil data
    Dim URLDAT As String, URLFOR As String
    URLDAT = "https://" & SubPath & "." & Author & ".eu.org/" & PathData
    
    ' Menyiapkan QueryTable dan mengambil data
    On Error GoTo RefreshError
    With wsData.QueryTables.Add(Connection:="URL;" & URLDAT, Destination:=wsData.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With
    
    ' Hanya menampilkan baris SearchValue
    Dim rng As Range
    Set rng = wsData.UsedRange
    rng.AutoFilter Field:=1, Criteria1:="<>" & SearchValue
    If rng.columns(1).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
        rng.offset(1, 0).Resize(rng.Rows.Count - 1, rng.columns.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    wsData.AutoFilterMode = False
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").Value
    URLFOR = "https://" & SubPath & "." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        On Error Resume Next
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("AA1"))
            .Refresh BackgroundQuery:=False
        End With
    End If

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    ' Pengaturan lainnya:
    ThisWorkbook.Sheets("DEV").Range("F8").Value = "UPDATE"
    Dev.CekIDPerangkat
    Dev.CekIPPublik
    Dev.CekNamaKomputer
    Dev.CekVersiOffice
    Dev.SimpanWaktu
    'Dev.AutoHideColumns
    Dev.HapusData
    Dev.CopyFormulas
    Dev.ProtectSheets
    'Dev.SendData
    
    ' Melindungi worksheet jika password diberikan
    Password = wsData.Range("G2").Value
    If Password <> "" Then
        wsData.Protect Password
    End If

    ' Menampilkan pesan setelah proses selesai
    Dim MessageUpdate As String
    MessageUpdate = wsData.Range("D2").Value

    If MessageUpdate = "" Then
        MsgBox "Username tidak terdaftar!", vbExclamation
    Else
        MsgBox MessageUpdate, vbInformation, "Informasi"
    End If
    Exit Sub

RefreshError:
    MsgBox UpdateErrorMsg, vbExclamation
End Sub

Sub GsheetDataLoginUpdate()
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String, SubPath As String
    Dim Password As String, Author As String
    Dim SearchValue As String
    Dim UpdateErrorMsg As String
    
    ' Pesan kesalahan
    UpdateErrorMsg = "Download ulang Aplikasi, hubungi Admin"
    
    ' Mengecek koneksi internet
    If Not Dev.TesKoneksi Then
        MsgBox "Tidak ada koneksi internet.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next

    ' Konfigurasi data
    Author = Env.Author
    SubPath = Env.SubPath
    PathData = Env.Token
    SheetNameData = Env.DataBase
    SearchValue = HalamanLogin.TextBoxUsername.Value
    
    If SearchValue = "" Or SearchValue = "Username" Then
        MsgBox "Masukkan Username terlebih dahulu!", vbExclamation
        Exit Sub
    End If
    
    ' Menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetHidden
    ThisWorkbook.Sheets(SheetNameData).Delete
    Application.DisplayAlerts = True

    ' Membuat lembar kerja baru
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = SheetNameData
    
    ' Hidden sheet DATAUSER
    ' ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetVeryHidden
    
    ' Membuat URL untuk mengambil data
    Dim URLDAT As String, URLFOR As String
    URLDAT = "https://" & SubPath & "." & Author & ".eu.org/" & PathData
    
    ' Menyiapkan QueryTable dan mengambil data
    On Error GoTo RefreshError
    With wsData.QueryTables.Add(Connection:="URL;" & URLDAT, Destination:=wsData.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With

    ' Hanya menampilkan baris SearchValue
    Dim rng As Range
    Set rng = wsData.UsedRange
    rng.AutoFilter Field:=1, Criteria1:="<>" & SearchValue
    If rng.columns(1).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
        rng.offset(1, 0).Resize(rng.Rows.Count - 1, rng.columns.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    wsData.AutoFilterMode = False
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").Value
    URLFOR = "https://" & SubPath & "." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        On Error Resume Next
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("AA1"))
            .Refresh BackgroundQuery:=False
        End With
    End If

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    ' Pengaturan lainnya:
    Dev.AutoHideColumns
    Dev.HapusData
    Dev.CopyFormulas
    Dev.ProtectSheets
    
    ' Melindungi worksheet jika password diberikan
    Password = wsData.Range("G2").Value
    If Password <> "" Then
        wsData.Protect Password
    End If
    
    ' Menampilkan pesan setelah proses selesai
    Dim MessageUpdate As String
    MessageUpdate = wsData.Range("D2").Value

    If MessageUpdate = "" Then
        MsgBox "Username tidak terdaftar!", vbExclamation
    Else
        MsgBox MessageUpdate, vbInformation, "Informasi"
    End If
    Exit Sub

RefreshError:
    MsgBox UpdateErrorMsg, vbExclamation
End Sub

Sub GsheetDataLogin()
    Dim wsData As Worksheet
    Dim SheetNameData As String
    Dim PathData As String, PathFormula As String, SubPath As String
    Dim Password As String, Author As String
    Dim SearchValue As String
    
    On Error Resume Next

    ' Konfigurasi data
    Author = Env.Author
    SubPath = Env.SubPath
    PathData = Env.Token
    SheetNameData = Env.DataBase
    SearchValue = HalamanLogin.TextBoxUsername.Value
    
    ' Menghapus lembar kerja yang sudah ada jika ada
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetHidden
    ThisWorkbook.Sheets(SheetNameData).Delete
    Application.DisplayAlerts = True

    ' Membuat lembar kerja baru
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = SheetNameData
    
    ' Hidden sheet DATAUSER
    'ThisWorkbook.Sheets(SheetNameData).Visible = xlSheetVeryHidden
    
    ' Membuat URL untuk mengambil data
    Dim URLDAT As String, URLFOR As String
    URLDAT = "https://" & SubPath & "." & Author & ".eu.org/" & PathData
    
    ' Menyiapkan QueryTable dan mengambil data
    With wsData.QueryTables.Add(Connection:="URL;" & URLDAT, Destination:=wsData.Range("A1"))
        .Refresh BackgroundQuery:=False
    End With

    ' Hanya menampilkan baris SearchValue
    Dim rng As Range
    Set rng = wsData.UsedRange
    rng.AutoFilter Field:=1, Criteria1:="<>" & SearchValue
    If rng.columns(1).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
        rng.offset(1, 0).Resize(rng.Rows.Count - 1, rng.columns.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    wsData.AutoFilterMode = False
    
    ' Membuat URL untuk mengambil data
    PathFormula = wsData.Range("F2").Value
    URLFOR = "https://" & SubPath & "." & Author & ".eu.org/" & PathFormula
    
    ' Menyiapkan QueryTable dan mengambil data
    If PathFormula <> "" Then
        With wsData.QueryTables.Add(Connection:="URL;" & URLFOR, Destination:=wsData.Range("AA1"))
            .Refresh BackgroundQuery:=False
        End With
    End If

    ' Menghapus semua koneksi data dalam workbook
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    ' Pengaturan lainnya:
    ThisWorkbook.Sheets("DEV").Range("F8").Value = "LOGIN"
    
    ' Melindungi worksheet jika password diberikan
    Password = wsData.Range("G2").Value
    If Password <> "" Then
        wsData.Protect Password
    End If

    Exit Sub
End Sub
