Sub GsheetData()
    Dim ws As Worksheet
    Dim key As String, gid As String, sheetName As String, startCell As String
    Dim qt As QueryTable, url As String
    Dim password As String
    
    ' Atur variabel berikut sesuai kebutuhan Anda
    key = "14V7IxlKuEXi7275zO2gxK2I47h6IlIL2UU82FUSrBNM"
    gid = "0"
    sheetName = "Sheet1" ' Ganti dengan nama worksheet yang Anda inginkan
    startCell = "A1" ' Ganti dengan sel awal yang Anda inginkan
    password = "ADMIN" ' Ganti dengan kata sandi perlindungan (jika diperlukan)
    
    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox "Tidak ada koneksi internet. Pastikan Anda terhubung ke internet dan coba lagi.", vbExclamation
        Exit Sub
    End If
    
    ' Memeriksa apakah worksheet ada
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Worksheet tidak ada, buat worksheet baru
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    Else
        ' Worksheet ada, periksa apakah dilindungi
        If ws.ProtectContents Then
            ' Worksheet dilindungi, coba melepas perlindungan dengan password
            On Error Resume Next
            ws.Unprotect password
            On Error GoTo 0
            If ws.ProtectContents Then
                MsgBox "Kata sandi salah.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If ws.QueryTables.Count > 0 Then ws.QueryTables(1).Delete
    ws.Cells.Clear
    
    url = "https://spreadsheets.google.com/tq?tqx=out:html&key=" & key & "&gid=" & gid
    
    Set qt = ws.QueryTables.Add(Connection:="URL;" & url, Destination:=ws.Range(startCell))
    
    With qt
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .BackgroundQuery = False
        .Refresh
    End With
    
    ' Melindungi kembali worksheet dengan password
    If password <> "" Then
        ws.Protect password
    End If
    
    ' Pesan dialog ketika proses selesai
    MsgBox "Proses selesai.", vbInformation
End Sub

Function IsInternetConnected() As Boolean
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    xhr.Open "GET", "https://www.google.com", False
    xhr.send
    IsInternetConnected = (xhr.Status = 200)
    On Error GoTo 0
End Function
