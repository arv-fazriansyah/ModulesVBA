Sub MailMerged()
    Dim fso As Object, http As Object, stream As Object
    Dim wordApp As Object, wordDoc As Object
    Dim ws As Worksheet
    Dim basePath As String, tempFolder As String, tempFolderClean As String
    Dim outputFolder As String, filePath As String, pdfPath As String
    Dim downloadURL As String
    Dim header As String, value As String
    Dim i As Long, lastCol As Long

    ' Inisialisasi objek
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Sheets("mail")
    basePath = ThisWorkbook.Path & "\"

    ' Path folder RBKdownload
    tempFolder = basePath & "RBKdownload\"
    tempFolderClean = basePath & "RBKdownload" ' tanpa backslash, untuk delete
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder

    ' Path folder output
    outputFolder = basePath & "GENERATE RBK 2025\"
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    ' Download file template
    downloadURL = "https://docs.google.com/document/d/1O6gAYOr3B4CNybingzGrXNAPcsFR2MWUBAwRBR9T-nE/export?format=doc"
    filePath = tempFolder & "template.docx"

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
        MsgBox "Gagal mendownload file", vbCritical
        Exit Sub
    End If

    ' Buka dokumen Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(filePath)

    ' Ambil data dari sheet "mail" baris 2
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        header = Trim(ws.Cells(1, i).value)
        value = ws.Cells(2, i).value

        With wordDoc.Content.Find
            .Text = "<<" & header & ">>"
            .Replacement.Text = value
            .Forward = True
            .Wrap = 1 ' wdFindContinue
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next i

    ' Simpan dokumen sebagai PDF
    pdfPath = outputFolder & "RBK_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
    wordDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17

    ' Tutup Word
    wordDoc.Close False
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing

    ' Bersihkan folder sementara
    On Error Resume Next
    If fso.FolderExists(tempFolderClean) Then
        fso.DeleteFolder tempFolderClean, True
    End If

    MsgBox "PDF berhasil dibuat!", vbInformation
End Sub

