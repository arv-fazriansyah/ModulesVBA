Sub GsheetData()
    Dim qt As QueryTable, url As String
    Dim key As String, gid As String, sheetName As String, startCell As String, password As String
    Dim ws As Worksheet
    Dim isProtected As Boolean
    
    ' Atur variabel berikut sesuai kebutuhan Anda
    key = "14V7IxlKuEXi7275zO2gxK2I47h6IlIL2UU82FUSrBNM"
    gid = "0"
    sheetName = "Sheet1" ' Ganti dengan nama worksheet yang Anda inginkan
    startCell = "A1" ' Ganti dengan sel awal yang Anda inginkan
    password = "ADMIN" ' Ganti dengan kata sandi proteksi worksheet Anda
    
    ' Mengecek koneksi internet
    If Not IsInternetConnected() Then
        MsgBox "Tidak ada koneksi internet. Pastikan Anda terhubung ke internet dan coba lagi.", vbExclamation
        Exit Sub
    End If
    
    ' Mengambil referensi ke worksheet yang ditentukan
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' Mengecek apakah worksheet telah diproteksi
    isProtected = ws.ProtectContents
    
    ' Unprotect worksheet jika telah diproteksi
    If isProtected Then
        ws.Unprotect password:=password
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
    
    ' Menghapus koneksi data setelah refresh
    qt.Delete
    
    ' Proteksi worksheet kembali jika sebelumnya telah diproteksi
    If isProtected Then
        ws.Protect password:=password
    End If
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

