Sub MailMerged()

    ' ================== KONFIGURASI ==================
    Const TEMPLATE_ID As String = "1LH72QD8PkerC3Q70h61uoi8WrIAjAKRwfG8xZ8b42cw"
    Const SHEET_NAME As String = "MAIL"
    Const OUTPUT_FOLDER_NAME As String = "GENERATE RBK 2025"
    Const TEMPLATE_NAME As String = "Cover"
    Const TEMP_FOLDER_NAME As String = "tempDownload"
    Const DATA_RANGE As String = "A1:H50"
    Const FIELD_NAMA As String = "COVER <<up_sekolah>> <<up_kecamtan>>"
    ' ==================================================

    Dim fso As Object, http As Object, stream As Object
    Dim wordApp As Object, wordDoc As Object
    Dim ws As Worksheet
    Dim basePath As String, tempFolder As String, outputFolder As String
    Dim filePath As String, pdfPath As String
    Dim downloadURL As String
    Dim i As Long, j As Long
    Dim dataRange As Range, dataArray As Variant
    Dim headers() As String
    Dim processedCount As Long: processedCount = 0
    Dim headerIndex As Object

    ' OPTIMASI EXCEL
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Set dataRange = ws.Range(DATA_RANGE)
    dataArray = dataRange.value

    basePath = ThisWorkbook.Path & "\"
    tempFolder = basePath & TEMP_FOLDER_NAME
    outputFolder = basePath & OUTPUT_FOLDER_NAME & "\"

    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    ' Download template Word dari Google Drive
    downloadURL = "https://docs.google.com/document/d/" & TEMPLATE_ID & "/export?format=doc"
    filePath = tempFolder & "\" & TEMPLATE_NAME

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", downloadURL, False
    http.Send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        With stream
            .Type = 1
            .Open
            .Write http.ResponseBody
            .SaveToFile filePath, 2
            .Close
        End With
    Else
        MsgBox "Gagal mendownload file template.", vbCritical
        Exit Sub
    End If

    ' Ambil header baris pertama
    ReDim headers(1 To UBound(dataArray, 2))
    Set headerIndex = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(dataArray, 2)
        headers(j) = Trim(dataArray(1, j))
        headerIndex(headers(j)) = j
    Next j

    ' Ambil placeholder nama file dari FIELD_NAMA
    Dim rawParts() As String
    rawParts = Split(FIELD_NAMA, " ")

    ' Buka Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    ' Proses data
    For i = 2 To UBound(dataArray, 1)
        If Trim(dataArray(i, 1)) <> "" Then
            Set wordDoc = wordApp.Documents.Open(filePath)

            ' Ganti semua placeholder
            For j = 1 To UBound(dataArray, 2)
                With wordDoc.Content.Find
                    .Text = "<<" & headers(j) & ">>"
                    .Replacement.Text = dataArray(i, j)
                    .Forward = True
                    .Wrap = 1
                    .Execute Replace:=2
                End With
            Next j

            ' Susun nama file dari FIELD_NAMA
            Dim namaFile As String, part, cleanHeader As String
            namaFile = ""
            For Each part In rawParts
                If Left(part, 2) = "<<" And Right(part, 2) = ">>" Then
                    cleanHeader = Mid(part, 3, Len(part) - 4)
                    If headerIndex.exists(cleanHeader) Then
                        namaFile = namaFile & " " & dataArray(i, headerIndex(cleanHeader))
                    End If
                Else
                    namaFile = namaFile & " " & part
                End If
            Next part

            namaFile = Trim(namaFile)
            namaFile = SanitizeFileName(namaFile)

            pdfPath = outputFolder & namaFile & ".pdf"
            wordDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
            wordDoc.Close False
            processedCount = processedCount + 1
        End If
    Next i

    wordApp.Quit
    Set wordApp = Nothing

    ' Bersihkan folder temp
    On Error Resume Next
    If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
    On Error GoTo 0

    ' Kembalikan setting Excel
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

    MsgBox "Berhasil membuat " & processedCount & " file PDF.", vbInformation

End Sub

' Hapus karakter ilegal dari nama file
Function SanitizeFileName(ByVal nama As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        nama = Replace(nama, invalidChars(i), "")
    Next i
    SanitizeFileName = Trim(nama)
End Function
