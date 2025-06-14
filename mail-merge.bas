Sub DownloadEditAndExportPDF()
    Dim fso As Object, http As Object, stream As Object
    Dim wordApp As Object, wordDoc As Object
    Dim ws As Worksheet
    Dim basePath As String, tempFolder As String, outputFolder As String
    Dim filePath As String, pdfPath As String
    Dim i As Long, lastCol As Long
    Dim header As String, value As String
    Dim downloadURL As String

    ' Inisialisasi objek
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Sheets("mail")
    basePath = ThisWorkbook.Path & "\"
    tempFolder = basePath & "RBKdownload\"
    outputFolder = basePath & "GENERATE RBK 2025\"
    filePath = tempFolder & "template.docx"
    downloadURL = "https://docs.google.com/document/d/1O6gAYOr3B4CNybingzGrXNAPcsFR2MWUBAwRBR9T-nE/export?format=doc"

    ' Buat folder jika belum ada
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    ' Download file Word
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", downloadURL, False
    http.Send
    If http.Status <> 200 Then
        MsgBox "Gagal mendownload file.", vbCritical
        Exit Sub
    End If

    ' Simpan ke file sementara
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1
        .Open
        .Write http.responseBody
        .SaveToFile filePath, 2 ' Overwrite
        .Close
    End With

    ' Buka dan edit dokumen Word
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Open(filePath)
    wordApp.Visible = False

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        header = Trim(ws.Cells(1, i).Value)
        value = ws.Cells(2, i).Value
        With wordDoc.Content.Find
            .Text = "<<" & header & ">>"
            .Replacement.Text = value
            .Wrap = 1 ' wdFindContinue
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next i

    ' Ekspor ke PDF
    pdfPath = outputFolder & "RBK_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    wordDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17

    ' Tutup Word dan bersihkan
    wordDoc.Close False
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing

    ' Hapus folder RBKdownload
    If fso.FolderExists(ThisWorkbook.Path & "\RBKdownload") Then
        fso.DeleteFolder ThisWorkbook.Path & "\RBKdownload", True
    End If

    MsgBox "PDF berhasil dibuat di: " & pdfPath, vbInformation
End Sub
