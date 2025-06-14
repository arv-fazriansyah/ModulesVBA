Sub MailMerged()

    ' ================== KONFIGURASI (UBAH DI SINI) ==================
    Const TEMPLATE_ID As String = "1LH72QD8PkerC3Q70h61uoi8WrIAjAKRwfG8xZ8b42cw"
    Const SHEET_NAME As String = "MAIL"
    Const OUTPUT_FOLDER_NAME As String = "GENERATE RBK 2025"
    Const TEMPLATE_NAME As String = "Cover.docx"
    Const TEMP_FOLDER_NAME As String = "tempDownload"
    Const FILE_NAME_PREFIX As String = "RBK_"
    Const DATA_RANGE As String = "A1:H50" ' BATAS RANGE
    ' ================================================================

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

    ' OPTIMALKAN SETTING EXCEL
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    ' Inisialisasi objek
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Set dataRange = ws.Range(DATA_RANGE)
    dataArray = dataRange.value ' simpan data dalam array

    basePath = ThisWorkbook.Path & "\"
    tempFolder = basePath & TEMP_FOLDER_NAME
    outputFolder = basePath & OUTPUT_FOLDER_NAME & "\"

    ' Buat folder sementara jika belum ada
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    ' Unduh template Word
    downloadURL = "https://docs.google.com/document/d/" & TEMPLATE_ID & "/export?format=doc"
    filePath = tempFolder & "\" & TEMPLATE_NAME

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", downloadURL, False
    http.Send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        With stream
            .Type = 1 ' Binary
            .Open
            .Write http.responseBody
            .SaveToFile filePath, 2 ' Overwrite
            .Close
        End With
    Else
        MsgBox "Gagal mendownload file template dari Google Docs.", vbCritical
        Exit Sub
    End If

    ' Simpan header di array
    ReDim headers(1 To UBound(dataArray, 2))
    For j = 1 To UBound(dataArray, 2)
        headers(j) = Trim(dataArray(1, j))
    Next j

    ' Buka Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    ' Proses tiap baris
    For i = 2 To UBound(dataArray, 1)
        If Trim(dataArray(i, 1)) <> "" Then ' Cek kolom A tidak kosong
            Set wordDoc = wordApp.Documents.Open(filePath)

            For j = 1 To UBound(dataArray, 2)
                With wordDoc.Content.Find
                    .Text = "<<" & headers(j) & ">>"
                    .Replacement.Text = dataArray(i, j)
                    .Forward = True
                    .Wrap = 1 ' wdFindContinue
                    .Execute Replace:=2 ' wdReplaceAll
                End With
            Next j

            ' Simpan PDF
            pdfPath = outputFolder & FILE_NAME_PREFIX & Format(Now, "yyyymmdd_hhnnss") & "_" & i & ".pdf"
            wordDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
            wordDoc.Close False

            processedCount = processedCount + 1
        End If
    Next i

    wordApp.Quit
    Set wordApp = Nothing

    ' Bersihkan folder sementara
    On Error Resume Next
    If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
    On Error GoTo 0

    ' Kembalikan pengaturan Excel
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

    MsgBox "PDF berhasil dibuat sebanyak " & processedCount & " file.", vbInformation

End Sub

