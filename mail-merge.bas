Sub MailMerged()
    ' ====== KONFIGURASI ======
    Const TEMPLATE_ID As String = "1LH72QD8PkerC3Q70h61uoi8WrIAjAKRwfG8xZ8b42cw"
    Const SHEET_NAME As String = "MAIL"
    Const OUTPUT_FOLDER As String = "GENERATE RBK 2025"
    Const TEMPLATE_NAME As String = "Cover.docx"
    Const TEMP_FOLDER As String = "tempDownload"
    Const DATA_RANGE As String = "A1:H50"
    Const FIELD_NAMA As String = "COVER <<up_sekolah>> <<up_kecamtan>>"
    ' ==========================

    Dim ws As Worksheet, fso As Object, http As Object, stream As Object
    Dim wordApp As Object, wordDoc As Object
    Dim dataArray As Variant, headers() As String, headerMap As Object
    Dim basePath As String, tempPath As String, outPath As String
    Dim filePath As String, pdfPath As String, fileName As String
    Dim i As Long, j As Long, count As Long

    ' OPTIMASI
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    dataArray = ws.Range(DATA_RANGE).value
    basePath = ThisWorkbook.Path & "\"
    tempPath = basePath & TEMP_FOLDER
    outPath = basePath & OUTPUT_FOLDER & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(tempPath) Then fso.CreateFolder tempPath
    If Not fso.FolderExists(outPath) Then fso.CreateFolder outPath

    ' DOWNLOAD TEMPLATE
    filePath = tempPath & "\" & TEMPLATE_NAME
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://docs.google.com/document/d/" & TEMPLATE_ID & "/export?format=doc", False
    http.Send
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        With stream
            .Type = 1: .Open: .Write http.ResponseBody
            .SaveToFile filePath, 2: .Close
        End With
    Else
        MsgBox "Gagal download template.", vbCritical: Exit Sub
    End If

    ' SIAPKAN HEADER
    Set headerMap = CreateObject("Scripting.Dictionary")
    ReDim headers(1 To UBound(dataArray, 2))
    For j = 1 To UBound(dataArray, 2)
        headers(j) = Trim(dataArray(1, j))
        headerMap(headers(j)) = j
    Next

    ' BUKA WORD
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    ' PROSES TIAP BARIS
    For i = 2 To UBound(dataArray, 1)
        If Trim(dataArray(i, 1)) <> "" Then
            Set wordDoc = wordApp.Documents.Open(filePath)
            For j = 1 To UBound(headers)
                With wordDoc.Content.Find
                    .Text = "<<" & headers(j) & ">>"
                    .Replacement.Text = dataArray(i, j)
                    .Execute Replace:=2
                End With
            Next

            ' SUSUN NAMA FILE
            fileName = ""
            For Each part In Split(FIELD_NAMA, " ")
                If Left(part, 2) = "<<" And Right(part, 2) = ">>" Then
                    Dim h: h = Mid(part, 3, Len(part) - 4)
                    If headerMap.exists(h) Then fileName = fileName & " " & dataArray(i, headerMap(h))
                Else
                    fileName = fileName & " " & part
                End If
            Next
            fileName = Trim(SanitizeFileName(fileName))

            ' EKSPOR PDF
            pdfPath = outPath & fileName & ".pdf"
            wordDoc.ExportAsFixedFormat pdfPath, 17
            wordDoc.Close False
            count = count + 1
        End If
    Next

    wordApp.Quit: Set wordApp = Nothing

    ' HAPUS FOLDER TEMP
    On Error Resume Next: If fso.FolderExists(tempPath) Then fso.DeleteFolder tempPath, True

    ' KEMBALIKAN SETTING
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

    MsgBox "Berhasil membuat " & count & " file PDF.", vbInformation
End Sub

Function SanitizeFileName(nama As String) As String
    Dim invalidChars: invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim c: For Each c In invalidChars: nama = Replace(nama, c, ""): Next
    SanitizeFileName = Trim(nama)
End Function
